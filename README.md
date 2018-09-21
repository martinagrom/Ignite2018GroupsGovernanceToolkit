# Ignite18
Groups Governance Toolkit 
Demos for Ignite18 by atwork.at
Contributors: Christoph Wilfing @CWilfing, Toni Pohl @atwork, Martina Grom @magrom

# Description
These code samples include a bunch of Azure Function PowerShell scripts for groups governance in an Office 365 tenant.
The functions require an app in the Azure Actice Directory with permissions to read and write users and groups in the Microsoft Graph provider and the tenant id. These information must be present in the App Settings of the Function App in the keys AppId, AppSecret and TenantID. Access to Azure Table Storage and queues is provided by the bindings. Some functions execute a Logic App for further processing.
See more about the How-To and the functionality at blog.atwork.at end of September.

# Quick overview
Currently, each function has a starting function, the odd function number, as f1, f3, f5 and f7, and a second function that does the work for one item. These are the even function numbers, as f2, f4, f6 and f8. The functions f3, f5 and f7 are time-triggered, while f1 is triggered by an HTTP call.

# Run these functions
Start these functions to provision a new group or to create reports:
f1 - creates a new group or team, could be called from a Flow per HTTP
f3 - creates the report "ownerless" - runs once a day
f5 - creates the report "externalguests" - runs once a day
f7 - creates the statistics table for PowerBI - runs once a day
