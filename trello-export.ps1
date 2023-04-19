# Prompt for Trello Board to export and API keys to use
# Get your key and generate a token here: https://trello.com/app-key
$apiKey = Read-Host -Prompt "Enter Trello API key (personal key)"
$apiToken = Read-Host -Prompt "Enter Trello API token"
# Find the boardkey by opening the trello board and adding '.json' to the URL.
$boardKey = Read-Host -Prompt "Enter Board key (24 char)"
$baseUrl = "https://api.trello.com/1"
$exportFolder = "c:\temp\" #boardId will be appended

enum Change {
    Add = 0
    Remove = 1
}

function Start-DownloadAttachment {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ApiKey,
        [Parameter(Mandatory = $true)]
        [string] $ApiToken,
        [Parameter(Mandatory = $true)]
        [string] $Id,
        [Parameter(Mandatory = $true)]
        [string] $Url,
        [Parameter(Mandatory = $true)]
        [string] $Path
    )
    $dest = Join-Path -Path $Path -ChildPath "Attachments\${id}\"
    if (-not (Test-Path -Path $dest -PathType Container)) {
        New-Item -ItemType Directory -Path $dest | Out-Null
    }
    $destFile = (Join-path -Path $dest -ChildPath (Split-Path $url -leaf))
    $oldProgressPreference = $progressPreference 
    $progressPreference = 'SilentlyContinue' 
    Invoke-WebRequest -Uri $url -Headers @{Authorization = "OAuth oauth_consumer_key=`"$ApiKey`", oauth_token=`"$ApiToken`""} -OutFile $destFile
    $progressPreference = $oldProgressPreference
    return $destFile
}

# Function to get the board name from the Trello API
function Get-BoardName {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ApiKey,
        [Parameter(Mandatory = $true)]
        [string] $ApiToken,
        [Parameter(Mandatory = $true)]
        [string] $BoardKey
    )

    $uri = "${baseUrl}/boards/${BoardKey}?key=${ApiKey}&token=${ApiToken}"
    $response = Invoke-RestMethod -Method Get -Uri $uri
    $boardName = $response.name

    return $boardName
}

# Function to get the board Url from the Trello API
function Get-BoardUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ApiKey,
        [Parameter(Mandatory = $true)]
        [string] $ApiToken,
        [Parameter(Mandatory = $true)]
        [string] $BoardKey
    )

    $uri = "${baseUrl}/boards/${BoardKey}?key=${ApiKey}&token=${ApiToken}"
    $response = Invoke-RestMethod -Method Get -Uri $uri
    $boardUrl = $response.url

    return $boardUrl
}

# Function to export checklist actions to separate files
function Export-TrelloChecklistActions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ApiKey,
        [Parameter(Mandatory=$true)]
        [string]$ApiToken,
        [Parameter(Mandatory=$true)]
        [string]$BoardKey,
        [Parameter(Mandatory=$true)]
        [string]$CardId,
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $checklistsUrl = "${baseUrl}/cards/$($card.id)/checklists?key=$ApiKey&token=$ApiToken"
    $checklists = Invoke-RestMethod -Uri $checklistsUrl -Method Get

    $actionUri = "${baseUrl}/cards/$($card.id)/actions?key=${ApiKey}&token=${ApiToken}&filter=updateCheckItemStateOnCard,updateChecklist"
    $actionsResponse = Invoke-RestMethod -Method Get -Uri $actionUri
    
    foreach ($checklist in $checklists) {
        $checklistsUrl = "${baseUrl}/checklists/$($checklist.id)?key=$ApiKey&token=$ApiToken"
        $checklist = Invoke-RestMethod -Uri $checklistsUrl -Method Get

        if ($checklist) {            
            foreach ($checkitem in $checklist.checkItems) {

                $checkItemUpdates = $actionsResponse | where-object {$_.data.checkItem.id -eq $checkitem.id}
                $lastUpdate = $null
                if($checkItemUpdates.Count -gt 1) {
                    $lastUpdate = ($checkItemUpdates | Sort-Object -Property date)[0]
                }
                else {
                    $lastUpdate = $checkItemUpdates
                }

                $author = $lastUpdate.memberCreator.fullName ?? "Anonymous"
                $created = $lastUpdate.date
                if($null -ne $created) {
                    $created = $created.ToUniversalTime()
                }

                $revision = [Ordered]@{
                    Author = $author
                    Time = $null -eq $created ? $null : (Get-Date -Date $created -Format "yyyy-MM-ddTHH:mm:ssZ")
                    Index = 0
                    Fields = @(
                        [Ordered]@{
                            ReferenceName = "System.Title"
                            Value = $checkitem.name
                        },
                        [Ordered]@{
                            ReferenceName = "System.State"
                            Value = switch -regex ($checkitem.state) {
                                "^(?i)complete$" { "Closed" }
                                default { "New" }
                            }
                        }
                    )
                    Links = @()
                    Attachments = @()
                    AttachmentReferences = $card.attachments.Count -gt 0
                }

                #Link to parent User Story
                $parentLink = [Ordered]@{
                    Change = [Change]::Add
                    TargetOriginId = $cardId
                    TargetWiId = 0
                    WiType = "System.LinkTypes.Hierarchy-Reverse"
                }
                $revision.Links += $parentLink
                
                $outputObject = [Ordered]@{
                    Type = "Task"
                    WiId = -1
                    OriginId = $checkitem.id
                    Revisions = @($revision)
                }

                # Write the output object to a file
                $fileName = "$($checkitem.id).json"
                $outputObject | ConvertTo-Json -Depth 99 | Out-File (Join-Path -Path $Path -ChildPath $fileName) -Force
            }
        }
    }
}

function Export-TrelloAttachment {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ApiKey,
        [Parameter(Mandatory = $true)]
        [string] $ApiToken,
        [Parameter(Mandatory = $true)]
        [PSCustomObject] $CardData,
        [Parameter(Mandatory = $true)]
        [PSCustomObject] $Action,
        [Parameter(Mandatory = $true)]
        [ref] $revisionIndex,
        [Parameter(Mandatory = $true)]
        [string] $Path
    )
    # Get the attachments currently on card
    $currentAttachments = Invoke-RestMethod "$baseUrl/cards/$($card.id)/attachments?key=${apiKey}&token=${apiToken}"
    if($currentAttachments.Count -gt 0)
    {                    
        $addAttachmentAuthor = $action.memberCreator.fullName
        $addAttachmentDate = $action.date.ToUniversalTime()
        $attachmentId = $action.data.attachment.id
        $attachmentUrl = $action.data.attachment.url
        
        #Match attachment URI
        if($attachmentUrl -match "^https:\/\/trello\.com\/1\/cards\/[a-zA-Z0-9]*\/attachments\/[a-zA-Z0-9]*\/download\/.*$") {
            #Check if attachment is deleted
            if($currentAttachments.id -match $attachmentId)
            {
                $filePath = Start-DownloadAttachment $ApiKey $ApiToken $attachmentId $attachmentUrl $Path
            }

            $attachmentRevision = [Ordered]@{
                Author = $addAttachmentAuthor
                Time = Get-Date -Date $addAttachmentDate -Format "yyyy-MM-ddTHH:mm:ssZ"
                Index = $revisionIndex.Value
                Fields = @( 
                    [Ordered]@{
                        ReferenceName = "System.History"
                        Value = "Added attachment $($action.data.attachment.name)"
                    }
                )
                Attachments = @(
                    [Ordered]@{
                        Change = [Change]::Add
                        FilePath = $filePath
                        Comment = "Imported from Trello"
                        AttOriginId = $attachmentId
                    }
                )
                AttachmentReferences = $card.attachments.Count -gt 0
            }
            $cardData.Revisions += $attachmentRevision
            $revisionIndex.Value++
        }
        elseif($attachmentUrl -match "^https:\/\/trello\.com\/c\/[a-zA-Z0-9]*\/.*$") {
            # A link to another card
        }
        else {
            # A Hyperlink
            Write-Warning "Manually add hyperlink $attachmentUrl to story `"$($card.name)`""
        }
    }
}

function Export-TrelloComment {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject] $CardData,
        [Parameter(Mandatory = $true)]
        [PSCustomObject] $Action,
        [Parameter(Mandatory = $true)]
        [ref] $revisionIndex
    )

    $commentAuthor = $action.memberCreator.fullName
    $commentDate = $action.date.ToUniversalTime()

    if ($action.data.text) {
        $commentRevision = [Ordered]@{
            Author = $commentAuthor
            Time = Get-Date -Date $commentDate -Format "yyyy-MM-ddTHH:mm:ssZ"
            Index = $revisionIndex.Value
            Fields = @( 
                [Ordered]@{
                    ReferenceName = "System.History"
                    Value = $action.data.text
                }
            )
        }
    }
    $cardData.Revisions += $commentRevision
    $revisionIndex.Value++
}

# Function to export Trello board data to JSON
function Export-TrelloBoardData {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ApiKey,
        [Parameter(Mandatory = $true)]
        [string] $ApiToken,
        [Parameter(Mandatory = $true)]
        [string] $BoardKey,
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $uri = "${baseUrl}/boards/${BoardKey}/cards?key=${ApiKey}&token=${ApiToken}&members=true&member_fields=all&actions=commentCard,addAttachmentToCard&action_memberCreator_fields=all&action_member_fields=all&attachments=true&attachment_fields=all&fields=all"
    $response = Invoke-RestMethod -Method Get -Uri $uri
    
    foreach ($card in $response) {
        $revisionIndex = 1
        $cardData = [Ordered]@{
            Type = "User Story"
            OriginId = $card.id
            WiId = -1
            Revisions = @()
        }

        $cardListUri = "${baseUrl}/cards/$($card.id)/list?key=${APIKey}&token=${APIToken}"
        $cardListResponse = Invoke-RestMethod -Method Get -Uri $cardListUri

        $createCardAction = "${baseUrl}/cards/$($card.id)/actions?key=${ApiKey}&token=${ApiToken}&filter=createCard"
        $createCardActionResponse = Invoke-RestMethod -Method Get -Uri $createCardAction
        if($createCardActionResponse) {
            $author = $createCardActionResponse[0].memberCreator.fullName ?? "Anonymous"
            $createdDate = $createCardActionResponse[0].date.ToUniversalTime()
        }
            
        $tags = @()
        if ($null -ne $card.labels -and $card.labels.Count -gt 0) {
            $tags += $card.labels.name
        }
        if ($null -ne $card.due) {
            $tags += $card.due
        }
        if ($null -ne $card.members -and $card.labels.Count -gt 0) {
            $tags += $card.members.fullName
        }
        if($card.idList -match "On Hold")
        {
            $tags += "On Hold"
        }

        $descriptionRevision = [Ordered]@{
            Author = $author
            Time = Get-Date -Date $createdDate -Format "yyyy-MM-ddTHH:mm:ssZ"
            Index = 0
            Fields = @(
                [Ordered]@{
                    ReferenceName = "System.Title"
                    Value = $card.name
                },
                [Ordered]@{
                    ReferenceName = "System.Tags"
                    Value = $tags -join ","
                }
                [Ordered]@{
                    ReferenceName = "System.State"
                    Value = switch -regex ($cardListResponse.name) {
                        "(?i)removed" { "Removed" }
                        "(?i)todo" { "New" }
                        "(?i)done" { "Closed" }
                        default { "Active" }
                    }
                }
            )
            Links = @()
            Attachments = @()
            AttachmentReferences = $card.attachments.Count -gt 0
        }

        #Use only the latest update to the description field
        if ($card.desc) {
            $descriptionAuthorUri = "${baseUrl}/cards/$($card.id)/actions?key=${ApiKey}&token=${apiToken}&filter=updateCard:desc&limit=1"
            $descriptionAuthorResponse = Invoke-RestMethod -Method Get -Uri $descriptionAuthorUri
            $author = $descriptionAuthorResponse.memberCreator.fullName

            $descriptionRevision.Fields += [Ordered]@{
                ReferenceName = "System.Description"
                Value = $card.desc
            }
        }
        $cardData.Revisions += $descriptionRevision

        #Parse historical actions on the card
        foreach ($action in $card.actions | Where-Object { $_.type -in @("commentCard","addAttachmentToCard") } | Sort-Object -Property date) {            
            #action is a comment
            if($action.type -eq "commentCard") {
                Export-TrelloComment -CardData $cardData -Action $action -RevisionIndex ([ref]$revisionIndex)
            }
            #action is an attachment
            elseif ($action.type -eq "addAttachmentToCard") {
                Export-TrelloAttachment -CardData $cardData -Action $action -RevisionIndex ([ref]$revisionIndex) -ApiKey $ApiKey -ApiToken $ApiToken -Path $Path
            }
        }

        # Write the data to a file with the card ID as the filename
        $fileName = "$($card.id).json"
        $cardData | ConvertTo-Json -Depth 99 | Out-File (Join-Path -Path $Path -ChildPath $fileName) -Force

        if($card.badges.checkItems -gt 0) {
            Export-TrelloChecklistActions -ApiKey $ApiKey -ApiToken $ApiToken -BoardKey $BoardKey -CardId $($card.id) -Path $Path
        }
    }
}

# Start export of board
Write-Output "Exporting $(Get-BoardName -BoardKey $boardKey -ApiKey $apiKey -ApiToken $apiToken)..."
$boardid = ((Get-BoardUrl -BoardKey $boardKey -ApiKey $apiKey -ApiToken $apiToken) -split '/')[-1]
$boardExportPath = (Join-Path -Path $exportFolder -ChildPath $boardid)
if (-not (Test-Path -Path $boardExportPath -PathType Container)) {
    New-Item -ItemType Directory -Path $boardExportPath | Out-Null
    New-Item -ItemType Directory -Path (Join-Path -Path $boardExportPath -ChildPath "Attachments") | Out-Null
}

Export-TrelloBoardData $boardKey -ApiKey $apiKey -ApiToken $apiToken -Path $boardExportPath