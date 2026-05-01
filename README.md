# Exchange Online Self-Service Delegation Manager

A robust, interactive PowerShell utility designed to streamline and standardize Exchange Online administrative tasks. 

This menu-driven tool allows IT Administrators and Information Systems Engineers to rapidly manage Mailbox, Calendar, and Distribution Group permissions without needing to manually construct complex ExchangeOnlineManagement cmdlets or navigate the Microsoft 365/Entra ID portals.

## 🚀 Features

* **Interactive Menu System:** Clean, easy-to-navigate console interface for high-speed ticket resolution.
* **Intelligent Error Handling:** Built-in validation checks to ensure target identities exist before executing configuration changes. Safely catches and handles non-terminating Exchange pipeline errors.
* **Mailbox Delegation:** View, grant, and revoke *Full Access*, *Send As*, and *Send on Behalf* permissions.
* **Calendar Management:** Granular control over folder-level permissions (e.g., Editor, Reviewer, PublishingAuthor) with built-in reference sheets.
* **Distribution Groups:** View, add, and remove members, translating raw GUIDs and aliases into readable Display Names.
* **Automated Dependency Management:** Automatically checks for and installs the required `ExchangeOnlineManagement` V3 module if it is not present on the local machine.

## 📋 Prerequisites

To run this script, you must have:
* **PowerShell 5.1** or later.
* **Permissions:** An Entra ID account with the *Exchange Administrator* or *Global Administrator* role.
* **Execution Policy:** Your local PowerShell execution policy must allow the running of scripts (e.g., `Set-ExecutionPolicy RemoteSigned`).

## 🛠️ Usage

Download or clone the script to your local machine. Execute the script from your terminal by passing your admin User Principal Name (UPN) to the `-Username` parameter.

```powershell
.\exchange-script.ps1 -Username admin@yourdomain.com
