# Introduction
This PowerShell script exports all the cards from a Trello board and converts each card into a Azure DevOps 'User Story' work item. If a checklist has been added to the card, each item on the checklist is converted into a 'Task' work item. Comments on cards are preserved as comments on the work item, and attachments are also preserved. However, (hyper)links are currently not implemented.

An Agile Template for Azure DevOps is expected.

# Export
To export the cards from Trello using this script, follow these steps:

* Log in to Trello
* Create a 'Personal Key' and 'API Token' on https://trello.com/app-key
* Open your board and append '.json' to the URL to get the board key
* Run the PowerShell script and input the above values when prompted.

# Import to DevOps
The output format is compatible with the format defined by https://github.com/solidify/jira-azuredevops-migrator. You can use their `wi-import` tool to import the exported data to Azure DevOps.