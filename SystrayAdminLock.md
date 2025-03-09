# Restricting NinjaOne System Tray Scripts to Admin-Only Access

This guide covers how to create scripts for the NinjaOne system tray and configure some to execute exclusively for administrators as allowed by a checkbox inside NinjaOne, enhancing security and control over script execution.

## Step 1: Access Global Custom Fields
Navigate to **Settings > Administration > Global Custom Fields** in the NinjaOne interface to begin setting up the necessary configurations.

## Step 2: Define a New Global Custom Field
Set up a global custom field with the following details to track admin status:

| Field             | Value        |
|-------------------|--------------|
| Custom Field Type | Check box    |
| Label            | AdminStatus  |
| Definition Scope | Device       |

### Configure Permissions
Assign the appropriate permissions to control access to this field:

| Role         | Permission   |
|--------------|--------------|
| Technician   | Editable     |
| Automation   | Read/Write   |
| API          | None         |

## Restricting System Tray Scripts
Incorporate the following PowerShell code into any system tray script you want to limit to admin users only:

```powershell
$AdminStatusFieldName = "AdminStatus"
$MessageTitle = "Access Denied"
$MessageBody = "This script requires administrative privileges. Contact your admin for assistance."

$AdminStatus = Ninja-Property-Get $AdminStatusFieldName
if ($AdminStatus -ne 1) {
    $Session = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName
    if ($Session) {
        $Username = $Session.Split('\')[1]
        Invoke-Expression "msg $($Username) /TIME:30 '$MessageTitle - $MessageBody'"
        Write-Output "Message sent to $($Username): Admin access required."
    } else {
        Write-Output "No active user session detected to notify."
    }
} else {
    ## Insert Script to run here!
    Write-Output "Admin access granted. Running admin script."
}
```

## Safeguarding the Admin Restriction

To prevent the `AdminStatus` field from being left enabled accidentally, set up an automated process to enforce its restricted state. 
Depending on your preferences, configure this as an automation policy or a scheduled task. 
Execute the following script hourly to automatically disable the field if it’s been overlooked:

```powershell
$AdminStatusFieldName = "AdminStatus"
$AdminStatus = Ninja-Property-Get $AdminStatusFieldName
if ($AdminStatus -ne 0) {
    Ninja-Property-Set $AdminStatusFieldName 0
    Write-Output "AdminStatus has been successfully disabled."
}
else {
    Write-Output "AdminStatus is already in a disabled state."
}
```
## Configuring the System Tray for Admin-Only Scripts

Next, let’s configure the NinjaOne system tray to clearly distinguish and organize scripts reserved for admin use. This setup ensures they’re both easily recognizable and securely managed.

### Steps:
1. Go to **Administration > Branding > Systray** in the NinjaOne interface.
2. Either create a new system tray configuration or edit an existing one.
3. Add the following elements to structure your admin-only scripts:

| Menu Item Type    | Details                     |
|-------------------|-----------------------------|
| Separator         | (Creates a visual break)   |
| Group             | Label: "Admin Only Scripts" |
| Automation        | Your admin-specific scripts |

### Explanation:
- **Separator**: Inserts a dividing line in the tray menu to enhance visual separation.
- **Group**: Establishes a labeled section titled "Admin Only Scripts" to categorize restricted scripts.
- **Automation**: Nest your admin-only automations (e.g., scripts with the `AdminStatus` check) under the "Admin Only Scripts" group. This nesting ensures these scripts appear as submenu items beneath the group label, keeping them organized and clearly tied to their admin-only purpose.

This configuration not only isolates admin scripts visually in the system tray but also reinforces their restricted access through the `AdminStatus` check, providing a seamless experience for technicians.
