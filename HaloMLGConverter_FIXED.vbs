'=========================================================
' Halo MLG Converter - Bidirectional VBS
' Refract JSON <-> Hayha format
' Refract uses shipping only when converting to Hayha.
'=========================================================
Option Explicit

Dim shell, fso, temp, psPath, outFile, ps1, cmd
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

temp = shell.ExpandEnvironmentStrings("%TEMP%")
psPath = temp & "\HaloMLGConverter_Bidirectional.ps1"

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
ps1 = ps1 & "function TextValue($value) { if ($null -eq $value) { return '' }; return [string]$value }" & vbCrLf
ps1 = ps1 & "function RequireProperty($object, [string]$propertyName, [string]$profileLabel) {" & vbCrLf
ps1 = ps1 & "  if ($null -eq $object) { throw ('Missing required object for ' + $profileLabel + '.') }" & vbCrLf
ps1 = ps1 & "  $property = $object.PSObject.Properties[$propertyName]" & vbCrLf
ps1 = ps1 & "  if ($null -eq $property -or $null -eq $property.Value) { throw ('Missing ' + $propertyName + ' on ' + $profileLabel + '.') }" & vbCrLf
ps1 = ps1 & "  return $property.Value" & vbCrLf
ps1 = ps1 & "}" & vbCrLf
ps1 = ps1 & "" & vbCrLf

ps1 = ps1 & "$form = New-Object System.Windows.Forms.Form" & vbCrLf
ps1 = ps1 & "$form.Text = 'Halo MLG Converter'" & vbCrLf
ps1 = ps1 & "$form.Size = New-Object System.Drawing.Size(680,490)" & vbCrLf
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
ps1 = ps1 & "$sub.Text = 'Profile format converter | Refract uses shipping address only'" & vbCrLf
ps1 = ps1 & "$sub.Font = New-Object System.Drawing.Font('Consolas',10)" & vbCrLf
ps1 = ps1 & "$sub.AutoSize = $true" & vbCrLf
ps1 = ps1 & "$sub.Location = New-Object System.Drawing.Point(62,72)" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($sub)" & vbCrLf

ps1 = ps1 & "$modeLabel = New-Object System.Windows.Forms.Label" & vbCrLf
ps1 = ps1 & "$modeLabel.Text = 'CONVERSION:'" & vbCrLf
ps1 = ps1 & "$modeLabel.Font = New-Object System.Drawing.Font('Consolas',11,[System.Drawing.FontStyle]::Bold)" & vbCrLf
ps1 = ps1 & "$modeLabel.AutoSize = $true" & vbCrLf
ps1 = ps1 & "$modeLabel.Location = New-Object System.Drawing.Point(95,108)" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($modeLabel)" & vbCrLf

ps1 = ps1 & "$mode = New-Object System.Windows.Forms.ComboBox" & vbCrLf
ps1 = ps1 & "$mode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList" & vbCrLf
ps1 = ps1 & "$mode.Font = New-Object System.Drawing.Font('Consolas',11)" & vbCrLf
ps1 = ps1 & "$mode.Location = New-Object System.Drawing.Point(220,104)" & vbCrLf
ps1 = ps1 & "$mode.Size = New-Object System.Drawing.Size(365,28)" & vbCrLf
ps1 = ps1 & "[void]$mode.Items.Add('Refract -> Hayha')" & vbCrLf
ps1 = ps1 & "[void]$mode.Items.Add('Hayha -> Refract')" & vbCrLf
ps1 = ps1 & "$mode.SelectedIndex = 0" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($mode)" & vbCrLf

ps1 = ps1 & "$global:ImportPath = ''" & vbCrLf
ps1 = ps1 & "$log = New-Object System.Windows.Forms.TextBox" & vbCrLf
ps1 = ps1 & "$log.Multiline = $true" & vbCrLf
ps1 = ps1 & "$log.ReadOnly = $true" & vbCrLf
ps1 = ps1 & "$log.ScrollBars = 'Vertical'" & vbCrLf
ps1 = ps1 & "$log.BackColor = [System.Drawing.Color]::Black" & vbCrLf
ps1 = ps1 & "$log.ForeColor = [System.Drawing.Color]::FromArgb(115,255,115)" & vbCrLf
ps1 = ps1 & "$log.Font = New-Object System.Drawing.Font('Consolas',10)" & vbCrLf
ps1 = ps1 & "$log.Location = New-Object System.Drawing.Point(35,225)" & vbCrLf
ps1 = ps1 & "$log.Size = New-Object System.Drawing.Size(600,175)" & vbCrLf
ps1 = ps1 & "$log.Text = 'READY > Select a conversion direction, then import a profile file.'" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($log)" & vbCrLf
ps1 = ps1 & "function Log([string]$t) { $log.AppendText([Environment]::NewLine + $t) }" & vbCrLf

ps1 = ps1 & "$btnImport = New-Object System.Windows.Forms.Button" & vbCrLf
ps1 = ps1 & "$btnImport.Text = 'IMPORT'" & vbCrLf
ps1 = ps1 & "$btnImport.Font = New-Object System.Drawing.Font('Consolas',18,[System.Drawing.FontStyle]::Bold)" & vbCrLf
ps1 = ps1 & "$btnImport.Location = New-Object System.Drawing.Point(95,153)" & vbCrLf
ps1 = ps1 & "$btnImport.Size = New-Object System.Drawing.Size(200,55)" & vbCrLf
ps1 = ps1 & "$btnImport.BackColor = [System.Drawing.Color]::FromArgb(8,42,14)" & vbCrLf
ps1 = ps1 & "$btnImport.ForeColor = [System.Drawing.Color]::FromArgb(130,255,130)" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($btnImport)" & vbCrLf

ps1 = ps1 & "$btnExport = New-Object System.Windows.Forms.Button" & vbCrLf
ps1 = ps1 & "$btnExport.Text = 'EXPORT'" & vbCrLf
ps1 = ps1 & "$btnExport.Font = New-Object System.Drawing.Font('Consolas',18,[System.Drawing.FontStyle]::Bold)" & vbCrLf
ps1 = ps1 & "$btnExport.Location = New-Object System.Drawing.Point(380,153)" & vbCrLf
ps1 = ps1 & "$btnExport.Size = New-Object System.Drawing.Size(200,55)" & vbCrLf
ps1 = ps1 & "$btnExport.BackColor = [System.Drawing.Color]::FromArgb(8,42,14)" & vbCrLf
ps1 = ps1 & "$btnExport.ForeColor = [System.Drawing.Color]::FromArgb(130,255,130)" & vbCrLf
ps1 = ps1 & "$form.Controls.Add($btnExport)" & vbCrLf

ps1 = ps1 & "$btnImport.Add_Click({" & vbCrLf
ps1 = ps1 & "  $ofd = New-Object System.Windows.Forms.OpenFileDialog" & vbCrLf
ps1 = ps1 & "  $ofd.Title = 'Select profile file to convert'" & vbCrLf
ps1 = ps1 & "  $ofd.Filter = 'Profile files (*.json;*.hayha)|*.json;*.hayha|All files (*.*)|*.*'" & vbCrLf
ps1 = ps1 & "  if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {" & vbCrLf
ps1 = ps1 & "    $global:ImportPath = $ofd.FileName" & vbCrLf
ps1 = ps1 & "    Log ('IMPORT > ' + $global:ImportPath)" & vbCrLf
ps1 = ps1 & "  }" & vbCrLf
ps1 = ps1 & "})" & vbCrLf

ps1 = ps1 & "$btnExport.Add_Click({" & vbCrLf
ps1 = ps1 & "  if ([string]::IsNullOrWhiteSpace($global:ImportPath) -or -not (Test-Path -LiteralPath $global:ImportPath)) {" & vbCrLf
ps1 = ps1 & "    [System.Windows.Forms.MessageBox]::Show('Import a profile file first.','Halo MLG Converter') | Out-Null" & vbCrLf
ps1 = ps1 & "    return" & vbCrLf
ps1 = ps1 & "  }" & vbCrLf
ps1 = ps1 & "  $direction = [string]$mode.SelectedItem" & vbCrLf
ps1 = ps1 & "  $sfd = New-Object System.Windows.Forms.SaveFileDialog" & vbCrLf
ps1 = ps1 & "  $sfd.Title = 'Save converted profile file'" & vbCrLf
ps1 = ps1 & "  if ($direction -eq 'Refract -> Hayha') {" & vbCrLf
ps1 = ps1 & "    $sfd.FileName = 'refract_converted.hayha'" & vbCrLf
ps1 = ps1 & "    $sfd.Filter = 'Hayha files (*.hayha)|*.hayha|JSON files (*.json)|*.json|All files (*.*)|*.*'" & vbCrLf
ps1 = ps1 & "  } else {" & vbCrLf
ps1 = ps1 & "    $sfd.FileName = 'hayha_converted.json'" & vbCrLf
ps1 = ps1 & "    $sfd.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'" & vbCrLf
ps1 = ps1 & "  }" & vbCrLf
ps1 = ps1 & "  if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }" & vbCrLf
ps1 = ps1 & "  try {" & vbCrLf
ps1 = ps1 & "    $raw = Get-Content -LiteralPath $global:ImportPath -Raw -Encoding UTF8" & vbCrLf
ps1 = ps1 & "    $profiles = @($raw | ConvertFrom-Json)" & vbCrLf
ps1 = ps1 & "    if ($profiles.Count -eq 0 -or $null -eq $profiles[0]) { throw 'No profiles found in the imported file.' }" & vbCrLf
ps1 = ps1 & "    $out = @()" & vbCrLf
ps1 = ps1 & "    if ($direction -eq 'Refract -> Hayha') {" & vbCrLf
ps1 = ps1 & "      $groupId = New-GuidText" & vbCrLf
ps1 = ps1 & "      foreach ($p in $profiles) {" & vbCrLf
ps1 = ps1 & "        $label = ('Refract profile ' + (TextValue $p.name))" & vbCrLf
ps1 = ps1 & "        $s = RequireProperty $p 'shipping' $label" & vbCrLf
ps1 = ps1 & "        $pay = RequireProperty $p 'payment' $label" & vbCrLf
ps1 = ps1 & "        $out += [ordered]@{" & vbCrLf
ps1 = ps1 & "          name = TextValue $p.name" & vbCrLf
ps1 = ps1 & "          shipping = [ordered]@{" & vbCrLf
ps1 = ps1 & "            firstName = TextValue $s.firstName" & vbCrLf
ps1 = ps1 & "            lastName = TextValue $s.lastName" & vbCrLf
ps1 = ps1 & "            email = TextValue $p.email" & vbCrLf
ps1 = ps1 & "            phone = TextValue $s.phone" & vbCrLf
ps1 = ps1 & "            address = TextValue $s.address1" & vbCrLf
ps1 = ps1 & "            address2 = TextValue $s.address2" & vbCrLf
ps1 = ps1 & "            country = TextValue $s.country" & vbCrLf
ps1 = ps1 & "            state = StateCode (TextValue $s.province)" & vbCrLf
ps1 = ps1 & "            city = TextValue $s.city" & vbCrLf
ps1 = ps1 & "            zipCode = TextValue $s.postalCode" & vbCrLf
ps1 = ps1 & "          }" & vbCrLf
ps1 = ps1 & "          sameAsBilling = $true" & vbCrLf
ps1 = ps1 & "          cardInfo = [ordered]@{" & vbCrLf
ps1 = ps1 & "            cardNumber = TextValue $pay.num" & vbCrLf
ps1 = ps1 & "            holder = TextValue $pay.name" & vbCrLf
ps1 = ps1 & "            expMonth = TextValue $pay.month" & vbCrLf
ps1 = ps1 & "            expYear = [int]$pay.year" & vbCrLf
ps1 = ps1 & "            cvv = TextValue $pay.cvv" & vbCrLf
ps1 = ps1 & "          }" & vbCrLf
ps1 = ps1 & "          id = New-GuidText" & vbCrLf
ps1 = ps1 & "          groupId = $groupId" & vbCrLf
ps1 = ps1 & "          encrypted = $false" & vbCrLf
ps1 = ps1 & "        }" & vbCrLf
ps1 = ps1 & "      }" & vbCrLf
ps1 = ps1 & "    } else {" & vbCrLf
ps1 = ps1 & "      foreach ($p in $profiles) {" & vbCrLf
ps1 = ps1 & "        $label = ('Hayha profile ' + (TextValue $p.name))" & vbCrLf
ps1 = ps1 & "        $s = RequireProperty $p 'shipping' $label" & vbCrLf
ps1 = ps1 & "        $card = RequireProperty $p 'cardInfo' $label" & vbCrLf
ps1 = ps1 & "        $out += [ordered]@{" & vbCrLf
ps1 = ps1 & "          name = TextValue $p.name" & vbCrLf
ps1 = ps1 & "          email = TextValue $s.email" & vbCrLf
ps1 = ps1 & "          shipping = [ordered]@{" & vbCrLf
ps1 = ps1 & "            firstName = TextValue $s.firstName" & vbCrLf
ps1 = ps1 & "            lastName = TextValue $s.lastName" & vbCrLf
ps1 = ps1 & "            address1 = TextValue $s.address" & vbCrLf
ps1 = ps1 & "            address2 = TextValue $s.address2" & vbCrLf
ps1 = ps1 & "            city = TextValue $s.city" & vbCrLf
ps1 = ps1 & "            province = TextValue $s.state" & vbCrLf
ps1 = ps1 & "            postalCode = TextValue $s.zipCode" & vbCrLf
ps1 = ps1 & "            country = TextValue $s.country" & vbCrLf
ps1 = ps1 & "            phone = TextValue $s.phone" & vbCrLf
ps1 = ps1 & "          }" & vbCrLf
ps1 = ps1 & "          billing = [ordered]@{" & vbCrLf
ps1 = ps1 & "            firstName = TextValue $s.firstName" & vbCrLf
ps1 = ps1 & "            lastName = TextValue $s.lastName" & vbCrLf
ps1 = ps1 & "            address1 = TextValue $s.address" & vbCrLf
ps1 = ps1 & "            address2 = TextValue $s.address2" & vbCrLf
ps1 = ps1 & "            city = TextValue $s.city" & vbCrLf
ps1 = ps1 & "            province = TextValue $s.state" & vbCrLf
ps1 = ps1 & "            postalCode = TextValue $s.zipCode" & vbCrLf
ps1 = ps1 & "            country = TextValue $s.country" & vbCrLf
ps1 = ps1 & "            phone = TextValue $s.phone" & vbCrLf
ps1 = ps1 & "          }" & vbCrLf
ps1 = ps1 & "          payment = [ordered]@{" & vbCrLf
ps1 = ps1 & "            num = TextValue $card.cardNumber" & vbCrLf
ps1 = ps1 & "            name = TextValue $card.holder" & vbCrLf
ps1 = ps1 & "            month = TextValue $card.expMonth" & vbCrLf
ps1 = ps1 & "            year = [int]$card.expYear" & vbCrLf
ps1 = ps1 & "            cvv = TextValue $card.cvv" & vbCrLf
ps1 = ps1 & "          }" & vbCrLf
ps1 = ps1 & "        }" & vbCrLf
ps1 = ps1 & "      }" & vbCrLf
ps1 = ps1 & "    }" & vbCrLf
ps1 = ps1 & "    $jsonOut = $out | ConvertTo-Json -Depth 20" & vbCrLf
ps1 = ps1 & "    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)" & vbCrLf
ps1 = ps1 & "    [System.IO.File]::WriteAllText($sfd.FileName, $jsonOut, $utf8NoBom)" & vbCrLf
ps1 = ps1 & "    Log ('MODE > ' + $direction)" & vbCrLf
ps1 = ps1 & "    Log ('EXPORT > ' + $sfd.FileName)" & vbCrLf
ps1 = ps1 & "    Log ('DONE > Converted ' + $out.Count + ' profile(s).')" & vbCrLf
ps1 = ps1 & "    [System.Windows.Forms.MessageBox]::Show(('Conversion complete! Converted ' + $out.Count + ' profile(s).'),'Halo MLG Converter') | Out-Null" & vbCrLf
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
