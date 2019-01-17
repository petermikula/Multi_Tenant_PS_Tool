<#
.SYNOPSIS
Functions-PSStoredCredentials - PowerShell functions to manage stored credentials for re-use

.DESCRIPTION 
This script adds two functions that can be used to manage stored credentials
on your admin workstation.

.EXAMPLE
. .\Functions-PSStoredCredentials.ps1

.LINK
https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

For more Office 365 tips, tricks and news
check out Practical 365.

* Website:	https://practical365.com
* Twitter:	https://twitter.com/practical365
#>
$KeyPath = "O365cred"


Function New-StoredCredential {

    <#
    .SYNOPSIS
    New-StoredCredential - Create a new stored credential

    .DESCRIPTION 
    This function will save a new stored credential to a .cred file.

    .EXAMPLE
    New-StoredCredential

    .LINK
    https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
    .NOTES
    Written by: Paul Cunningham

    Find me on:

    * My Blog:	http://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham
    * LinkedIn:	http://au.linkedin.com/in/cunninghamp/
    * Github:	https://github.com/cunninghamp

    For more Office 365 tips, tricks and news
    check out Practical 365.

    * Website:	https://practical365.com
    * Twitter:	https://twitter.com/practical365
    #>

    if (!(Test-Path Variable:\KeyPath)) {
        Write-Warning "The `$KeyPath variable has not been set. Consider adding `$KeyPath to your PowerShell profile to avoid this prompt."
        $path = Read-Host -Prompt "Enter a path for stored credentials"
        Set-Variable -Name KeyPath -Scope Global -Value $path

        if (!(Test-Path $KeyPath)) {
        
            try {
                New-Item -ItemType Directory -Path $KeyPath -ErrorAction STOP | Out-Null
            }
            catch {
                throw $_.Exception.Message
            }           
        }
    }

    $Credential = Get-Credential -Message "Enter a user name and password"

    $Credential.Password | ConvertFrom-SecureString | Out-File "$($KeyPath)\$($Credential.Username).cred" -Force

}



Function Get-StoredCredential {

    <#
    .SYNOPSIS
    Get-StoredCredential - Retrieve or list stored credentials

    .DESCRIPTION 
    This function can be used to list available credentials on
    the computer, or to retrieve a credential for use in a script
    or command.

    .PARAMETER UserName
    Get the stored credential for the username

    .PARAMETER List
    List the stored credentials on the computer

    .EXAMPLE
    Get-StoredCredential -List

    .EXAMPLE
    $credential = Get-StoredCredential -UserName admin@tenant.onmicrosoft.com

    .EXAMPLE
    Get-StoredCredential -List

    .LINK
    https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
    .NOTES
    Written by: Paul Cunningham

    Find me on:

    * My Blog:	http://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham
    * LinkedIn:	http://au.linkedin.com/in/cunninghamp/
    * Github:	https://github.com/cunninghamp

    For more Office 365 tips, tricks and news
    check out Practical 365.

    * Website:	https://practical365.com
    * Twitter:	https://twitter.com/practical365
    #>

    param(
        [Parameter(Mandatory=$false, ParameterSetName="Get")]
        [string]$UserName,
        [Parameter(Mandatory=$false, ParameterSetName="List")]
        [switch]$List
        )

    if (!(Test-Path Variable:\KeyPath)) {
        Write-Warning "The `$KeyPath variable has not been set. Consider adding `$KeyPath to your PowerShell profile to avoid this prompt."
        $path = Read-Host -Prompt "Enter a path for stored credentials"
        Set-Variable -Name KeyPath -Scope Global -Value $path
    }


    if ($List) {

        try {
        $CredentialList = @(Get-ChildItem -Path $keypath -Filter *.cred -ErrorAction STOP)

        foreach ($Cred in $CredentialList) {
            Write-Host "Username: $($Cred.BaseName)"
            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }

    }

    if ($UserName) {
        if (Test-Path "$($KeyPath)\$($Username).cred") {
        
            $PwdSecureString = Get-Content "$($KeyPath)\$($Username).cred" | ConvertTo-SecureString
            
            $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $PwdSecureString
        }
        else {
            throw "Unable to locate a credential for $($Username)"
        }

        return $Credential
    }
}


Function New-WarningRule ($credential)
<#
This function is adopted from https://gcits.com/knowledge-base/warn-users-external-email-arrives-display-name-someone-organisation/
#>
{
$ruleName = "External Senders with matching Display Names"
$ruleHtml = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width=`"100%`" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184; mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'>  <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'><td style='background:#910A19;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td><td width=`"100%`" style='width:100.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt; word-wrap:break-word' cellpadding=`"7px 5px 7px 15px`" color=`"#212121`"><div><p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:2.25pt; mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal: column;mso-height-rule:exactly'><span style='font-size:9.0pt;font-family: `"Segoe UI`",sans-serif;mso-fareast-font-family:`"Times New Roman`";color:#212121'>This message was sent from outside the company by someone with a display name matching a user in your organisation. Please do not click links or open attachments unless you recognise the source of this email and know the content is safe. <o:p></o:p></span></p></div></td></tr></table>"
 
 
Write-Host "Getting the Exchange Online cmdlets" -ForegroundColor Yellow
$Session = New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
    -ConfigurationName Microsoft.Exchange -Credential $credential `
    -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber -DisableNameChecking
 
$rule = Get-TransportRule | Where-Object {$_.Identity -contains $ruleName}
$displayNames = (Get-Mailbox -ResultSize Unlimited).DisplayName
 
if (!$rule) {
    Write-Host "Rule not found, creating rule" -ForegroundColor Green
    New-TransportRule -Name $ruleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" `
        -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $displayNames -ApplyHtmlDisclaimerText $ruleHtml
}
else {
    Write-Host "Rule found, updating rule" -ForegroundColor Green
    Set-TransportRule -Identity $ruleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" `
        -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $displayNames -ApplyHtmlDisclaimerText $ruleHtml
}
Remove-PSSession $Session
}
Function Show-Form1
# Main application window
{
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(700,400) 
$Form.text ="Available domains"
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Location = New-Object System.Drawing.Size(160,20) 
$groupBox.size = New-Object System.Drawing.Size(300,200) 
$groupBox.text = "Please check domains" 
$Form.Controls.Add($groupBox)
$Checkboxes = @()
$Checkboxes += New-Object System.Windows.Forms.CheckBox
# $Checkboxes.Location = New-Object System.Drawing.Size(10,20) 
$y = 20
$StoredCred = Get-ChildItem -Path $keypath -Filter *.cred
#$Label = New-Object System.Windows.Forms.textbox
#$Label.AutoSize = $True
Foreach ($cr in $StoredCred)
{
$Checkbox = New-Object System.Windows.Forms.CheckBox
$Checkbox.Text = $cr.BaseName
$Checkbox.Location = New-Object System.Drawing.Size(20,$y)
$Checkbox.AutoSize = $true
$y += 20
$groupBox.Controls.Add($Checkbox) 
$Checkboxes += $Checkbox
}
$groupBox.size = New-Object System.Drawing.Size(300,(50*$checkboxes.Count))

# Action button - Get Blocked list
$but1 = New-Object System.Windows.Forms.Button
$but1.Location = New-Object System.Drawing.Size(10,30)
$but1.Size = New-Object System.Drawing.Size(125,23)
$but1.Text = "Blocked Domains"
$but1.Enabled = $true
$Form.Controls.Add($but1)
$but1.Add_Click({Get-Checked ($Checkboxes)})

# Action button 2 - Set Blocked list
$but2 = New-Object System.Windows.Forms.Button
$but2.Location = New-Object System.Drawing.Size(10,60)
$but2.Size = New-Object System.Drawing.Size(125,23)
$but2.Text = "Add blocked"
$but2.Enabled = $true
$Form.Controls.Add($but2)
$but2.Add_Click({
$AddDomain = Get-UInput -Header "Add domain to block"
Get-Checked1 -Checkboxes $Checkboxes -AddDomain $AddDomain

})

# Action button 3 - New Credentials
$but3 = New-Object System.Windows.Forms.Button
$but3.Location = New-Object System.Drawing.Size(10,90)
$but3.Size = New-Object System.Drawing.Size(125,23)
$but3.Text = "New credentials"
$but3.Enabled = $true
$Form.Controls.Add($but3)
$but3.Add_Click({New-StoredCredential})

# Action button 4 - Warning Rule
$but4 = New-Object System.Windows.Forms.Button
$but4.Location = New-Object System.Drawing.Size(10,120)
$but4.Size = New-Object System.Drawing.Size(125,23)
$but4.Text = "Warning rule"
$but4.Enabled = $true
$Form.Controls.Add($but4)
$but4.Add_Click({Get-Checked2 ($Checkboxes)})

#OK button
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(600,300)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)
$form.ShowDialog()

}

Function Get-Checked ($Checkboxes)
{
Foreach ($ch in $Checkboxes)
{
if ($ch.CheckState -eq 1)
{
$dc=$ch.Text

Get-BlockedDomains -user $dc
}

}
}

Function Get-Checked1 ($Checkboxes,$AddDomain)
{
Foreach ($ch in $Checkboxes)
{
if ($ch.CheckState -eq 1)
{
$dc=$ch.Text

Set-BlockedDomains -user $dc -AddDomain $AddDomain

}

}
}

Function Get-Checked2 ($Checkboxes,$AddDomain)
{
Foreach ($ch in $Checkboxes)
{
if ($ch.CheckState -eq 1)
{
$dc=$ch.Text

$Credential = Get-StoredCredential -UserName $dc
New-WarningRule -credential $Credential
#Write-Host "Done!"
Show-Box -Header "Operation status" -FText "Operation completed."

}

}
}

Function Get-BlockedDomains ($user,$BlDomains)
# Connects to Office 365 and get Blocked domains lilst form spam filter
{
$Credentials = Get-StoredCredential -UserName $user
$Session = New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ -ConfigurationName Microsoft.Exchange -Credential $Credentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber -DisableNameChecking
$BlDomains = Get-HostedContentFilterPolicy Default | Select -ExpandProperty "BlockedSenderDomains"
Remove-PSSession $Session
Show-Box -FText $BlDomains -Header "Blocked Domains"
}

Function Set-BlockedDomains ($user,$BlDomains,$AddDomain)
# Connects to Office 365 and set Blocked domains lilst form spam filter
{
# Clear-Variable -Name BlDomains
$Credentials = Get-StoredCredential -UserName $user
$Session = New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ -ConfigurationName Microsoft.Exchange -Credential $Credentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber -DisableNameChecking
$BlDomains = Get-HostedContentFilterPolicy Default | Select -ExpandProperty "BlockedSenderDomains"
$BlDomains = $BlDomains + $AddDomain
Set-HostedContentFilterPolicy -Identity Default -BlockedSenderDomains $BlDomains
Remove-PSSession $Session
Show-Box -FText $BlDomains -Header "Blocked Domains"
}

Function Get-UInput ($Header)
{
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = $Header
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter the information in the space below:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $textBox.Text
    $x
}
}

Function Show-Box ($Header,$FText)
{
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object System.Windows.Forms.Form
$form.Text = $Header
$form.Size = New-Object System.Drawing.Size(400,300)
$form.StartPosition = 'CenterScreen'
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(125,220)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(200,220)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(350,200)
$label.Text = $FText
# $label.Multiline = $true
$form.Controls.Add($label)
$form.Topmost = $true
$result = $form.ShowDialog()
}
Show-Form1