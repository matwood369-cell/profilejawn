
'=========================================================
' Halo MLG Converter - FIXED VBS
' Refract JSON -> Hayha format
' Uses SHIPPING only for the address.
' No constant blue PowerShell console behind the program.
'=========================================================
Option Explicit

Dim shell, fso, temp, psPath, outFile, ps1, cmd
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

temp = shell.ExpandEnvironmentStrings("%TEMP%")
psPath = temp & "\HaloMLGConverter_FIXED.ps1"

ps1 = ""
ps1 = ps1 & "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf
ps1 = ps1 & "Add-Type -AssemblyName System.Drawing" & vbCrLf
ps1 = ps1 & "[System.Windows.Forms.Application]::EnableVisualStyles()" & vbCrLf
ps1 = ps1 & "$ErrorActionPreference = 'Stop'" & vbCrLf
ps1 = ps1 & "" & vbCrLf

ps1 = ps1 & "function New-GuidText { return ([guid]::NewGuid().ToString()) }" & vbCrLf
ps1 = ps1 & "function StateCode([string]$s) {" & vbCrLf
ps1 = ps1 & "  if ([string]::IsNullOrWhiteSpace($s)) { return '' }" & vbCrLf
ps1 = ps1 & "  $map = @{" & vbCrLf
ps1 = ps1 & "    'Alabama'='AL';'Alaska'='AK';'Arizona'='AZ';'Arkansas'='AR';'California'='CA';'Colorado'='CO';'Connecticut'='CT';'Delaware'='DE';'Florida'='FL';'Georgia'='GA';" & vbCrLf
ps1 = ps1 & "    'Hawaii'='HI';'Idaho'='ID';'Illinois'='IL';'Indiana'='IN';'Iowa'='IA';'Kansas'='KS';'Kentucky'='KY';'Louisiana'='LA';'Maine'='ME';'Maryland'='MD';" & vbCrLf
ps1 = ps1 & "    'Massachusetts'='MA';'Michigan'='MI';'Minnesota'='MN';'Mississippi'='MS';'Missouri'='MO';'Montana'='MT';'Nebraska'='NE';'Nevada'='NV';'New Hampshire'='NH';'New Jersey'='NJ';" & vbCrLf
ps1 = ps1 & "    'New Mexico'='NM';'New York'='NY';'North Carolina'='NC';'North Dakota'='ND';'Ohio'='OH';'Oklahoma'='OK';'Oregon'='OR';'Pennsylvania'='PA';'Rhode Island'='RI';" & vbCrLf
ps1 = ps1 & "    'South Carolina'='SC';'South Dakota'='SD';'Tennessee'='TN';'Texas'='TX';'Utah'='UT';'Vermont'='VT';'Virginia'='VA';'Washington'='WA';'West Virginia'='WV';'Wisconsin'='WI';'Wyoming'='WY'" & vbCrLf
ps1 = ps1 & "  }" & vbCrLf
ps1 = ps1 & "  if ($map.ContainsKey($s)) { return $map[$s] }" & vbCrLf
ps1 = ps1 & "  if ($s.Length -eq 2) { return $s.ToUpper() }" & vbCrLf
ps1 = ps1 & "  return $s" & vbCrLf
ps1 = ps1 & "}" & vbCrLf
ps1 = ps1 & "" & vbCrLf

ps1 = ps1 & "$form = New-Object System.Windows.Forms.Form" & vbCrLf
ps1 = ps1 & "$form.Text = 'Halo MLG Converter'" & vbCrLf
ps1 = ps1 & "$form.Size = New-Object System.Drawing.Size(680,450)" & vbCrLf
ps1 = ps1 & "$form.StartPosition = 'CenterScreen'" & vbCrLf
ps1 = ps1 & "$form.BackColor = [System.Drawing.Color]::FromArgb(3,10,5)" & vbCrLf
ps1 = ps1 & "$form.ForeColor = [System.Drawing.Color]::FromArgb(115,255,115)" & vbCrLf
ps1 = ps1 & "$form.FormBorderStyle = 'FixedDialog'" & vbCrLf
ps1 = ps1 & "$form.MaximizeBox = $false" & vbCrLf

ps1 = ps1 & "$title = New-Object System.Windows.Forms.Label" & vbCrLf
ps1 = ps1 & "$title.Text = 'HALO MLG CONVERTER'" & vbCrLf
ps1 = ps1 & "$title.Font = New-Object System.Drawing.Font('Consolas',26,[System.Drawing.FontStyle]::Bold)" & vbCrLf
ps1 = ps1 & "$title.AutoSize = $true" & vbCrLf
ps1 = ps1 & "$title.Location = New-Object System.Drawing.Point(105,22)" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($title)" & vbCrLf

ps1 = ps1 & "$sub = New-Object System.Windows.Forms.Label" & vbCrLf
ps1 = ps1 & "$sub.Text = 'Refract JSON  ->  Hayha format | shipping only'" & vbCrLf
ps1 = ps1 & "$sub.Font = New-Object System.Drawing.Font('Consolas',11)" & vbCrLf
ps1 = ps1 & "$sub.AutoSize = $true" & vbCrLf
ps1 = ps1 & "$sub.Location = New-Object System.Drawing.Point(140,72)" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($sub)" & vbCrLf

ps1 = ps1 & "$global:ImportPath = ''" & vbCrLf

ps1 = ps1 & "$log = New-Object System.Windows.Forms.TextBox" & vbCrLf
ps1 = ps1 & "$log.Multiline = $true" & vbCrLf
ps1 = ps1 & "$log.ReadOnly = $true" & vbCrLf
ps1 = ps1 & "$log.ScrollBars = 'Vertical'" & vbCrLf
ps1 = ps1 & "$log.BackColor = [System.Drawing.Color]::Black" & vbCrLf
ps1 = ps1 & "$log.ForeColor = [System.Drawing.Color]::FromArgb(115,255,115)" & vbCrLf
ps1 = ps1 & "$log.Font = New-Object System.Drawing.Font('Consolas',10)" & vbCrLf
ps1 = ps1 & "$log.Location = New-Object System.Drawing.Point(35,190)" & vbCrLf
ps1 = ps1 & "$log.Size = New-Object System.Drawing.Size(600,165)" & vbCrLf
ps1 = ps1 & "$log.Text = 'READY > Click IMPORT and pick your Refract JSON file.'" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($log)" & vbCrLf
ps1 = ps1 & "function Log([string]$t) { $log.AppendText([Environment]::NewLine + $t) }" & vbCrLf

ps1 = ps1 & "$btnImport = New-Object System.Windows.Forms.Button" & vbCrLf
ps1 = ps1 & "$btnImport.Text = 'IMPORT'" & vbCrLf
ps1 = ps1 & "$btnImport.Font = New-Object System.Drawing.Font('Consolas',18,[System.Drawing.FontStyle]::Bold)" & vbCrLf
ps1 = ps1 & "$btnImport.Location = New-Object System.Drawing.Point(95,112)" & vbCrLf
ps1 = ps1 & "$btnImport.Size = New-Object System.Drawing.Size(200,55)" & vbCrLf
ps1 = ps1 & "$btnImport.BackColor = [System.Drawing.Color]::FromArgb(8,42,14)" & vbCrLf
ps1 = ps1 & "$btnImport.ForeColor = [System.Drawing.Color]::FromArgb(130,255,130)" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($btnImport)" & vbCrLf

ps1 = ps1 & "$btnExport = New-Object System.Windows.Forms.Button" & vbCrLf
ps1 = ps1 & "$btnExport.Text = 'EXPORT'" & vbCrLf
ps1 = ps1 & "$btnExport.Font = New-Object System.Drawing.Font('Consolas',18,[System.Drawing.FontStyle]::Bold)" & vbCrLf
ps1 = ps1 & "$btnExport.Location = New-Object System.Drawing.Point(380,112)" & vbCrLf
ps1 = ps1 & "$btnExport.Size = New-Object System.Drawing.Size(200,55)" & vbCrLf
ps1 = ps1 & "$btnExport.BackColor = [System.Drawing.Color]::FromArgb(8,42,14)" & vbCrLf
ps1 = ps1 & "$btnExport.ForeColor = [System.Drawing.Color]::FromArgb(130,255,130)" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($btnExport)" & vbCrLf

ps1 = ps1 & "$btnImport.Add_Click({" & vbCrLf
ps1 = ps1 & "  $ofd = New-Object System.Windows.Forms.OpenFileDialog" & vbCrLf
ps1 = ps1 & "  $ofd.Title = 'Select Refract JSON file'" & vbCrLf
ps1 = ps1 & "  $ofd.Filter = 'JSON / Hayha files (*.json;*.hayha)|*.json;*.hayha|All files (*.*)|*.*'" & vbCrLf
ps1 = ps1 & "  if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {" & vbCrLf
ps1 = ps1 & "    $global:ImportPath = $ofd.FileName" & vbCrLf
ps1 = ps1 & "    Log ('IMPORT > ' + $global:ImportPath)" & vbCrLf
ps1 = ps1 & "  }" & vbCrLf
ps1 = ps1 & "})" & vbCrLf

ps1 = ps1 & "$btnExport.Add_Click({" & vbCrLf
ps1 = ps1 & "  if ([string]::IsNullOrWhiteSpace($global:ImportPath) -or -not (Test-Path -LiteralPath $global:ImportPath)) {" & vbCrLf
ps1 = ps1 & "    [System.Windows.Forms.MessageBox]::Show('Import a Refract JSON file first.','Halo MLG Converter') | Out-Null" & vbCrLf
ps1 = ps1 & "    return" & vbCrLf
ps1 = ps1 & "  }" & vbCrLf
ps1 = ps1 & "  $sfd = New-Object System.Windows.Forms.SaveFileDialog" & vbCrLf
ps1 = ps1 & "  $sfd.Title = 'Save converted Hayha file'" & vbCrLf
ps1 = ps1 & "  $sfd.FileName = 'refract_converted.hayha'" & vbCrLf
ps1 = ps1 & "  $sfd.Filter = 'Hayha files (*.hayha)|*.hayha|JSON files (*.json)|*.json|All files (*.*)|*.*'" & vbCrLf
ps1 = ps1 & "  if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }" & vbCrLf
ps1 = ps1 & "  try {" & vbCrLf
ps1 = ps1 & "    $raw = Get-Content -LiteralPath $global:ImportPath -Raw -Encoding UTF8" & vbCrLf
ps1 = ps1 & "    $profiles = $raw | ConvertFrom-Json" & vbCrLf
ps1 = ps1 & "    if ($null -eq $profiles) { throw 'No profiles found in the imported file.' }" & vbCrLf
ps1 = ps1 & "    $groupId = New-GuidText" & vbCrLf
ps1 = ps1 & "    $out = @()" & vbCrLf
ps1 = ps1 & "    foreach ($p in @($profiles)) {" & vbCrLf
ps1 = ps1 & "      $s = $p.shipping" & vbCrLf
ps1 = ps1 & "      $pay = $p.payment" & vbCrLf
ps1 = ps1 & "      if ($null -eq $s) { throw 'Missing shipping object on one profile.' }" & vbCrLf
ps1 = ps1 & "      if ($null -eq $pay) { throw 'Missing payment object on one profile.' }" & vbCrLf
ps1 = ps1 & "      $out += [ordered]@{" & vbCrLf
ps1 = ps1 & "        name = [string]$p.name" & vbCrLf
ps1 = ps1 & "        shipping = [ordered]@{" & vbCrLf
ps1 = ps1 & "          firstName = [string]$s.firstName" & vbCrLf
ps1 = ps1 & "          lastName = [string]$s.lastName" & vbCrLf
ps1 = ps1 & "          email = [string]$p.email" & vbCrLf
ps1 = ps1 & "          phone = [string]$s.phone" & vbCrLf
ps1 = ps1 & "          address = [string]$s.address1" & vbCrLf
ps1 = ps1 & "          address2 = [string]$s.address2" & vbCrLf
ps1 = ps1 & "          country = [string]$s.country" & vbCrLf
ps1 = ps1 & "          state = (StateCode ([string]$s.province))" & vbCrLf
ps1 = ps1 & "          city = [string]$s.city" & vbCrLf
ps1 = ps1 & "          zipCode = [string]$s.postalCode" & vbCrLf
ps1 = ps1 & "        }" & vbCrLf
ps1 = ps1 & "        sameAsBilling = $true" & vbCrLf
ps1 = ps1 & "        cardInfo = [ordered]@{" & vbCrLf
ps1 = ps1 & "          cardNumber = [string]$pay.num" & vbCrLf
ps1 = ps1 & "          holder = [string]$pay.name" & vbCrLf
ps1 = ps1 & "          expMonth = [string]$pay.month" & vbCrLf
ps1 = ps1 & "          expYear = [int]$pay.year" & vbCrLf
ps1 = ps1 & "          cvv = [string]$pay.cvv" & vbCrLf
ps1 = ps1 & "        }" & vbCrLf
ps1 = ps1 & "        id = (New-GuidText)" & vbCrLf
ps1 = ps1 & "        groupId = $groupId" & vbCrLf
ps1 = ps1 & "        encrypted = $false" & vbCrLf
ps1 = ps1 & "      }" & vbCrLf
ps1 = ps1 & "    }" & vbCrLf
ps1 = ps1 & "    $jsonOut = $out | ConvertTo-Json -Depth 20" & vbCrLf
ps1 = ps1 & "    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)" & vbCrLf
ps1 = ps1 & "    [System.IO.File]::WriteAllText($sfd.FileName, $jsonOut, $utf8NoBom)" & vbCrLf
ps1 = ps1 & "    Log ('EXPORT > ' + $sfd.FileName)" & vbCrLf
ps1 = ps1 & "    Log ('DONE > Converted ' + @($out).Count + ' profile(s).')" & vbCrLf
ps1 = ps1 & "    [System.Windows.Forms.MessageBox]::Show(('Conversion complete! Converted ' + @($out).Count + ' profile(s).'),'Halo MLG Converter') | Out-Null" & vbCrLf
ps1 = ps1 & "  } catch {" & vbCrLf
ps1 = ps1 & "    Log ('ERROR > ' + $_.Exception.Message)" & vbCrLf
ps1 = ps1 & "    [System.Windows.Forms.MessageBox]::Show(('Export failed: ' + $_.Exception.Message),'Halo MLG Converter') | Out-Null" & vbCrLf
ps1 = ps1 & "  }" & vbCrLf
ps1 = ps1 & "})" & vbCrLf

ps1 = ps1 & "[void]$form.ShowDialog()" & vbCrLf

Set outFile = fso.CreateTextFile(psPath, True)
outFile.Write ps1
outFile.Close

cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File " & Chr(34) & psPath & Chr(34)
shell.Run cmd, 0, False
