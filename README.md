# AD-GPO-DriveMapAudit
Efficiently audit and document Active Directory group memberships and GPO-driven mapped drives with this comprehensive PowerShell analysis tool.

## Active Directory Group and GPO Drive Mapping Analysis Script

This PowerShell script is designed to analyze and document the group membership and drive mapping policies within an Active Directory environment. It provides a comprehensive overview of security groups and their members, as well as the configuration of Group Policy Objects (GPOs) related to mapped drives.

#### Features:

- **Active Directory and GPO Integration**: Leverages the Active Directory and Group Policy modules to interact with domain entities and policies.
- **Domain and OU Specification**: Allows customization of the target Organizational Unit (OU) and the specific GPO to analyze.
- **Group Membership Retrieval**: Enumerates all security groups within the specified OU and lists their members, focusing on enabled user accounts.
- **Drive Mapping Analysis**: Extracts and reports on the drive mapping configurations defined within the specified GPO, including path, label, letter, and action.
- **Action Mapping**: Translates action codes within GPOs to human-readable actions (e.g., Create, Update, Delete, Replace).
- **Comprehensive Reporting**: Merges group membership data with drive mapping details to provide a full picture of access and resources.
- **Export Capabilities**: Outputs the analysis to either an Excel spreadsheet or a CSV file, including a timestamp for versioning.
- **Module Dependency Handling**: Checks for and attempts to install the 'ImportExcel' module if not present, falling back to CSV export if necessary.

This script is a valuable tool for administrators seeking to audit and document the current state of group memberships and drive mapping policies, ensuring that resources are allocated correctly and access is granted appropriately within the domain.
