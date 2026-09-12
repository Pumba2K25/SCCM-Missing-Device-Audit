# SCCM Missing Device Audit

PowerShell tool for auditing a list of Windows endpoints against **Microsoft Configuration Manager (SCCM/MECM)**, **Active Directory**, and a **single sequential ping**, then producing a team-friendly Excel workbook plus a technical CSV.

I originally built this for a physical device-reconciliation project where a plain device list was not enough to tell technicians what was still active, who likely owned it, or which systems were worth physically searching for.

## What it does

For each unique device in the input CSV, the script can:

- Match the device in SCCM / Configuration Manager
- Match the computer object in Active Directory
- Check whether the AD computer account is enabled
- Pull AD last-logon information
- Pull SCCM client / active status
- Pull SCCM last-active time
- Pull the SCCM username when available
- Resolve the user to a real name in Active Directory
- Resolve the user's manager in Active Directory
- Send **exactly one sequential ping per device**
- Generate a recommendation such as:
  - `Locate - Online now`
  - `Locate - Active in AD`
  - `Locate - Active in SCCM`
  - `Retirement Review - Not in AD or SCCM`
  - `Retirement Review - AD disabled / SCCM inactive`

## Output

The script creates:

- `Lost_Device_Audit_YYYYMMDD_HHmm.csv`
- `Lost_Device_Audit_YYYYMMDD_HHmm.xlsx`

The Excel workbook contains:

- **Dashboard**
- **By Manager**
- **Online**
- **Offline**
- **Retirement Review**
- **Manager Summary**

## Requirements

- Windows PowerShell 5.1
- Active Directory PowerShell module
- Configuration Manager PowerShell module
- Access to your SCCM / MECM environment
- Read access to Active Directory
- Microsoft Excel desktop application

The script uses Excel COM automation for workbook creation, so Excel must be installed on the computer running the script.

## Input CSV

The script attempts to automatically identify a device-name column.

Supported examples include:

- `Device`
- `Device Name`
- `Computer Name`
- `ComputerName`
- `Hostname`
- `Workstation`
- `Machine Name`
- `PC Name`
- `Asset Name`
- `Endpoint`
- `Name`

Optional columns can also be supplied for owner, manager, and notes.

Example:

```csv
Device,Owner,Manager,Notes
PC-001,jdoe,,Missing from physical inventory
PC-002,asmith,,Check old nursing station
PC-003,,,No owner listed
```

See [`examples/Example-DeviceList.csv`](examples/Example-DeviceList.csv).

## Usage

If a ConfigMgr PowerShell drive already exists:

```powershell
.\Invoke-MissingDeviceAudit.ps1 -InputCsv .\devices.csv
```

Or specify the site code and SMS Provider / site server:

```powershell
.\Invoke-MissingDeviceAudit.ps1 `
    -InputCsv .\devices.csv `
    -SiteCode "ABC" `
    -SiteServer "SCCM01.contoso.local"
```

You can also specify the output workbook path:

```powershell
.\Invoke-MissingDeviceAudit.ps1 `
    -InputCsv .\devices.csv `
    -OutputWorkbook C:\Reports\MissingDeviceAudit.xlsx
```

If `-InputCsv` is omitted, the script uses the newest usable CSV in the script folder.

## Network behavior

This was intentionally written to be conservative.

For every device, it performs:

- One ping
- Sequentially
- No retries
- No parallel pinging

The default ping timeout is 750 ms and can be changed:

```powershell
.\Invoke-MissingDeviceAudit.ps1 `
    -InputCsv .\devices.csv `
    -PingTimeoutMs 1000
```

## Safety / privacy

Do **not** commit real inventory exports or generated audit reports to a public repository.

The included `.gitignore` excludes CSV, Excel, logs, and common local configuration / secret files. The example CSV is explicitly allowed.

Before using this in your environment, review the code and adjust it for your organization's requirements.

## Notes

- SCCM and AD queries are read-only.
- A device being offline does not automatically mean it is retired.
- `LastLogonDate` and SCCM activity fields should be treated as supporting signals, not absolute proof of physical presence.
- User resolution depends on the quality of SCCM username data and your AD attributes.

## Feedback

This started as an internal utility and grew into a larger audit workflow. Suggestions, issues, and pull requests are welcome.
