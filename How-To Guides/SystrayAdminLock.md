# Restricting NinjaOne System Tray Scripts to 'Admin-Only' Access

This guide explains how to create scripts for the NinjaOne system tray and configure some to run exclusively for administrators, using a checkbox within NinjaOne to control access. This approach enhances security and oversight for script execution. However, this method is best suited for restricting scripts that you’d prefer end users not interact with, rather than for critical security measures. It’s ideal for minor administrative tasks where accidental access wouldn’t pose a significant security risk, not for safeguarding highly sensitive operations.

**Disclaimer**: I am not responsible for any actions you take based on this guide or the outcomes that result from implementing these configurations. Use at your own discretion and ensure they align with your organization’s security policies.

**Resources and Credits**:  
- All icons and fonts referenced in this guide can be found at [Google Fonts - Material Symbols Outlined](https://fonts.google.com/icons?icon.size=24&icon.color=%23e3e3e3).  
- Special thanks to the [NinjaOne Stream](https://www.youtube.com/watch?v=qBhk0awc3-c) for inspiration and insights.  
- Shoutout to JT (MrDrProfessorJT) and Trevor (StrikerTS) for sharing that resource, and to Joseph for the inspiration!

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

See below for a visual guide on the systray setup, and the message a user will see if they dont have the permissions to run this.

<img src="https://github.com/user-attachments/assets/6fa5af98-e91d-49a6-b1e8-24e474745bb1" alt="Admin Only Scripts System Tray Example" width="300"> <img src="https://github.com/user-attachments/assets/3e9481d7-bfbc-4106-90d4-dc4b03d5a6c7" alt="Invalid rights message" width="300">


