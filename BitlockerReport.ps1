<# 
AD BitLocker Escrow Inventory (GUI)
- Scans domain (or a specific OU) for Computer objects
- Excludes Windows Server by default
- Counts msFVE-RecoveryInformation children (recovery keys)
- Uses newest whenCreated as 'EncryptionDate'
- Shows GUI, allows CSV export
#>

[CmdletBinding()]
param(
  [string]$SearchBase,

  # Exclude servers by default (OperatingSystem -notlike '*Server*')
  [switch]$IncludeServers,

  # Ignore stale computer accounts: only include lastLogonDate within N days (0 = no filter)
  [int]$MaxLastLogonAgeDays = 0
)

function Ensure-Module {
  if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "ActiveDirectory module not found. Install RSAT or import the module before running."
  }
  Import-Module ActiveDirectory -ErrorAction Stop
}

function Get-Computers {
  $props = 'OperatingSystem','lastLogonDate'
  if ($SearchBase) {
    Get-ADComputer -SearchBase $SearchBase -Filter * -Properties $props
  } else {
    Get-ADComputer -Filter * -Properties $props
  }
}

function Is-ServerOS {
  param([string]$os)
  return ($os -match 'Server')
}

function Get-DeviceRow {
  param([Microsoft.ActiveDirectory.Management.ADComputer]$Computer)

  # Query recovery objects under the computer
  $recoveryObjs = Get-ADObject -SearchBase $Computer.DistinguishedName `
                               -LDAPFilter "(objectClass=msFVE-RecoveryInformation)" `
                               -Properties whenCreated `
                               -ErrorAction SilentlyContinue
  $count = ($recoveryObjs | Measure-Object).Count

  # Latest escrow date (proxy for encryption/rotation time)
  $latestDate = $null
  if ($count -gt 0) {
    $latestDate = ($recoveryObjs | Sort-Object whenCreated -Descending | Select-Object -First 1).whenCreated
  }

  [PSCustomObject]@{
    ComputerName        = $Computer.Name
    OperatingSystem     = $Computer.OperatingSystem
    LastLogonDate       = $Computer.lastLogonDate
    HasRecoveryKeyInAD  = [bool]($count -gt 0)
    RecoveryKeyCountAD  = $count
    EncryptionDate      = $latestDate
    DistinguishedName   = $Computer.DistinguishedName
  }
}

function Build-Report {
  Write-Verbose "Enumerating computers…"
  $all = Get-Computers

  if (-not $IncludeServers) {
    $all = $all | Where-Object { -not (Is-ServerOS $_.OperatingSystem) }
  }

  if ($MaxLastLogonAgeDays -gt 0) {
    $cutoff = (Get-Date).AddDays(-$MaxLastLogonAgeDays)
    $all = $all | Where-Object { $_.lastLogonDate -ge $cutoff }
  }

  $rows = foreach ($c in $all) { Get-DeviceRow -Computer $c }

  # Summary to console
  $total = $rows.Count
  $withKey = ($rows | Where-Object HasRecoveryKeyInAD).Count
  $withoutKey = $total - $withKey
  Write-Host ("Summary: {0} of {1} devices have BitLocker recovery keys in AD. ({2} without)" -f $withKey, $total, $withoutKey)

  # Return sorted rows
  $rows | Sort-Object ComputerName
}

function Show-Gui {
  param([System.Collections.IEnumerable]$Data)

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "BitLocker AD Escrow Inventory"
  $form.StartPosition = "CenterScreen"
  $form.Size = New-Object System.Drawing.Size(1200,700)

  $label = New-Object System.Windows.Forms.Label
  $label.AutoSize = $true
  $label.Location = New-Object System.Drawing.Point(10,10)
  $form.Controls.Add($label)

  $filterBox = New-Object System.Windows.Forms.TextBox
  $filterBox.PlaceholderText = "Filter (ComputerName / OS / DN)…"
  $filterBox.Width = 400
  $filterBox.Location = New-Object System.Drawing.Point(10,35)
  $form.Controls.Add($filterBox)

  $exportBtn = New-Object System.Windows.Forms.Button
  $exportBtn.Text = "Export CSV…"
  $exportBtn.Location = New-Object System.Drawing.Point(420,32)
  $form.Controls.Add($exportBtn)

  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Location = New-Object System.Drawing.Point(10,70)
  $grid.Size = New-Object System.Drawing.Size(1160,580)
  $grid.ReadOnly = $true
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.AutoSizeColumnsMode = 'AllCells'
  $grid.RowHeadersVisible = $false
  $grid.SelectionMode = 'FullRowSelect'
  $form.Controls.Add($grid)

  # Data binding
  $dt = New-Object System.Data.DataTable
  foreach ($col in 'ComputerName','OperatingSystem','LastLogonDate','HasRecoveryKeyInAD','RecoveryKeyCountAD','EncryptionDate','DistinguishedName') {
    [void]$dt.Columns.Add($col)
  }
  foreach ($row in $Data) {
    [void]$dt.Rows.Add(
      $row.ComputerName,
      $row.OperatingSystem,
      $row.LastLogonDate,
      $row.HasRecoveryKeyInAD,
      $row.RecoveryKeyCountAD,
      $row.EncryptionDate,
      $row.DistinguishedName
    )
  }

  $dv = New-Object System.Data.DataView($dt)
  $grid.DataSource = $dv

  # Summary text
  $total = $dt.Rows.Count
  $withKey = ($dt.Select("HasRecoveryKeyInAD = true")).Count
  $label.Text = "Devices with keys in AD: $withKey / $total    (filter by typing above)"

  # Filter logic (simple contains across a few columns)
  $filterBox.Add_TextChanged({
    $q = ($filterBox.Text -replace "'","''")
    if ([string]::IsNullOrWhiteSpace($q)) {
      $dv.RowFilter = ""
    } else {
      $dv.RowFilter = "ComputerName LIKE '%$q%' OR OperatingSystem LIKE '%$q%' OR DistinguishedName LIKE '%$q%'"
    }
  })

  # Export
  $exportBtn.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV files (*.csv)|*.csv"
    $sfd.FileName = ("BitLocker_AD_Escrow_Inventory_{0:yyyyMMdd_HHmm}.csv" -f (Get-Date))
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      # Export current view (respect filter)
      $out = $sfd.FileName
      $sb = New-Object System.Text.StringBuilder
      $headers = ($dt.Columns | ForEach-Object ColumnName) -join ","
      [void]$sb.AppendLine($headers)
      foreach ($r in $dv.ToTable().Rows) {
        $line = @(
          $r.ComputerName,
          $r.OperatingSystem,
          (Get-Date $r.LastLogonDate -ErrorAction SilentlyContinue).ToString("s"),
          $r.HasRecoveryKeyInAD,
          $r.RecoveryKeyCountAD,
          (Get-Date $r.EncryptionDate -ErrorAction SilentlyContinue).ToString("s"),
          $r.DistinguishedName
        ) -replace ',', ';'  # make CSV robust if commas appear
        [void]$sb.AppendLine(($line -join ","))
      }
      [System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.Encoding]::UTF8)
      [System.Windows.Forms.MessageBox]::Show("Exported to:`n$out","Export CSV")
    }
  })

  [void]$form.ShowDialog()
}

try {
  Ensure-Module
  $data = Build-Report
  Show-Gui -Data $data
}
catch {
  Write-Error $_
}
