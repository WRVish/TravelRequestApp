Information1. Security & Governance
External User Audit: Filter and display only "External Users" (Guest accounts) to verify who from outside the organization has access.
API Endpoint: Filter SiteUsers where LoginName contains #ext#.
Sharing Links Report: Detect items shared via "Anyone with the link" or "Specific People" links (which bypass standard group permissions).
API Endpoint: /_api/web/lists/getByTitle('...')/GetSharingInformation
Orphaned User Cleanup: Identify users listed in the "User Information List" who no longer exist in Azure AD or have been disabled.
API Endpoint: Cross-reference SiteUsers with /_api/SP.UserProfiles.PeopleManager/GetPropertiesFor (or handle 404s).
2. Remediation & Action (Write Capabilities)
Restore Inheritance: Allow admins to "Reset" a list or item with unique permissions back to inheriting from the parent, fixing broken permission limitations.
API Endpoint: /_api/web/lists/getByTitle('...')/ResetRoleInheritance
Grant Permissions: Add a "Add User" feature to quickly grant access directly from the report without going to the settings page.
API Endpoint: /_api/web/roleassignments/addroleassignment
Clone Permissions: "Copy" permissions from one user to another (e.g., when a new employee joins a team).
Logic: Read User A's roles -> Write User B's roles.
3. Deep Dive Insights
Item-Level Scan: Go deeper than Lists. Scan for individual files or folders inside libraries that have broken inheritance.
API Endpoint: /_api/web/lists/getbytitle('...')/items?$select=HasUniqueRoleAssignments
"Check Permissions" Simulator: Replicate the native "Check Permissions" feature but for multiple users at once (Bulk Check).
API Endpoint: /_api/web/GetUserEffectivePermissions
4. Modernization
Hub Site Association: Show permissions inherited from a Hub Site (if applicable).
API Endpoint: /_api/site/HubSiteId

