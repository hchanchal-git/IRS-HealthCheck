# Created By Himanshu Kumar (Cloud Ops)
# Ensure script runs in Windows PowerShell ISE
if ($host.Name -notmatch "ISE") {
    Write-Host "Redirecting script to Windows PowerShell ISE..." -ForegroundColor Yellow
    Start-Process PowerShell_ISE -ArgumentList "-File `"$PSCommandPath`""
    exit
}

$ErrorActionPreference = "Stop"

# Ignore Ctrl + C completely by setting trap
$global:IgnoreCtrlC = $true
trap {
    if ($global:IgnoreCtrlC) {
        Write-Host "`nCtrl + C detected! PowerShell will NOT stop the script." -ForegroundColor Yellow
        continue
    }
}

# Prompt for Case ID
$caseID = Read-Host "Enter Case ID (Create ticket in CBU dept using \"Customer Email\" <Moriah.D.Cardona@irs.gov>)"

# Date Formatting
function Get-DateOrdinalSuffix { param ([int]$day)
    switch ($day) {
        { ($_ -eq 11) -or ($_ -eq 12) -or ($_ -eq 13) } { return "th" }
        { ($_ % 10 -eq 1) } { return "st" }
        { ($_ % 10 -eq 2) } { return "nd" }
        { ($_ % 10 -eq 3) } { return "rd" }
        default { return "th" }
    }
}
$day = [int](Get-Date -Format "dd")
$month = Get-Date -Format "MMMM"
$year = Get-Date -Format "yyyy"
$ordinalSuffix = Get-DateOrdinalSuffix -day $day
$dateFormatted = "$month $day$ordinalSuffix, $year"

# Email Details
$subject = "IRS Daily Health Check Report - $dateFormatted [#$caseID]"
$body = @"
Dear Customer,

Please find the attached Daily Health Check Report - $dateFormatted.

Regards,  
eGain Corp
"@

$emails = @"
To:
Vaishali.P.Narkhede@irs.gov
Moriah.D.Cardona@irs.gov
Donald.W.Russell@irs.gov
jb551t@att.com
Michael.A.Harrison@irs.gov
George.B.Lenoir@irs.gov
Erik.C.Schlenker@irs.gov
Darren.E.Jackson@irs.gov
Geoffrey.T.Dang@irs.gov
Venus.M.Hutson@irs.gov
Jeronima.G.Gomez@irs.gov
Wayne.M.Garrido@irs.gov
Dionecio.Headley@irs.gov
Melissa.Shuman@irs.gov
Kartez.I.Harris@irs.gov
Cc:
VPadmanaban@eGain.com
pgawande@eGain.com
PBoyle@egain.com
AGupta@eGain.com
achan@eGain.com
BUfoegbune@egain.com
EKozlowski@eGain.com
support@eGain.com
jmallory@egain.com
radhav@eGain.com
JKinderman@egain.com
eGainCloudNotifications@eGain.com
Bcc:
cbu-tsops@egain.com
"@

# Combine the email details into a single string for GUI display
$emailContent = @"
Subject: $subject

Body:
$body

Recipients:
$emails
"@

# GUI using Windows Forms to display the draft email content
Add-Type -AssemblyName "System.Windows.Forms"

$form = New-Object System.Windows.Forms.Form
$form.Text = "Draft Email - Copy the Content"
$form.Size = New-Object System.Drawing.Size(700, 500)
$form.StartPosition = "CenterScreen"

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$textBox.Dock = "Fill"
$textBox.Text = $emailContent
$textBox.ReadOnly = $true
$form.Controls.Add($textBox)
$form.ShowDialog()

# Wait for user confirmation before proceeding
Read-Host "`nPress Enter to continue with the health check URLs..."

# IRS URLs
$urls = @(
    "https://sa.www4.irs.gov/idp/startSSO.ping?PartnerSpId=IRS-eGain-IDP&TargetResource=https%3A%2F%2Fconnect.irs.gov%2Fsystem%2Ftemplates%2Fmessagecenter%2Firssecure%2Fen-US%2FIRS%3Fpoa%3Dyes%26lp%3Dhttps%3A%2F%2Fsa.www4.irs.gov%2Fsso%2Fprotected%2Flogout",
    "https://connect.irs.gov/system/templates/messagecenter/irscorp/en-US/LBIEXAM",
    "https://connect.irs.gov/system/templates/messagecenter/irsaca/en-US/LBI",
    "https://connect.irs.gov/acalogin/irs_teb",
    "https://connect.irs.gov/system/templates/chat/irs_us/index.html?entryPointId=1004&templateName=irs_us&ver=v11&locale=en-US&eglvrefname=VBD009&referer=",
    "https://connect.irs.gov/system/templates/chat/irs_us/index.html?entryPointId=1003&templateName=irs_us&ver=v11&locale=es-ES&eglvrefname=VBD009&referer=",
    "https://www.irs.gov/payments",
    "https://www.irs.gov/es/payments",
    "https://www.irs.gov/refunds",
    "https://www.irs.gov/es/refunds",
    "https://www.irs.gov/cp2000",
    "https://www.irs.gov/cp2501",
    "https://www.irs.gov/cp3219A",
    "https://www.irs.gov/individuals/understanding-your-566-t-letter",
    "https://www.jobs.irs.gov/careers",
    "https://www.taxpayeradvocate.irs.gov/",
    "https://es.taxpayeradvocate.irs.gov/",
    "https://sa.www4.irs.gov/ola/",
    "https://sa.www4.irs.gov/ola/es/",
    "https://connect.irs.gov/system/web/custom/vascripts/erc_launch_va.html",
    "https://connect.irs.gov/system/web/custom/vascripts/erc_travel_launch_va.html",
    "https://connect.irs.gov/system/web/custom/vascripts/esd_launch_va.html",
	"https://connect.irs.gov/system/templates/chat/sbse_nlp_va2/index.html?entryPointId=1001&locale=en-US&postChatAttributes=false&templateName=sbse_nlp_va2&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=sbseenprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://www.irs.gov&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fwww.irs.gov%2Fpayments&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/sbse_nlp_va2_spanish/index.html?entryPointId=1001&locale=es-ES&postChatAttributes=false&templateName=sbse_nlp_va2_spanish&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=sbseesprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://www.irs.gov&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fwww.irs.gov%2Fes%2Fpayments&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/wni_va_rel1_ie/index.html?entryPointId=1001&locale=en-US&postChatAttributes=false&templateName=wni_va_rel1_ie&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=wnienprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://www.irs.gov&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fwww.irs.gov%2Frefunds&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/wni_va_rel1_spanish_ie/index.html?entryPointId=1001&locale=es-ES&postChatAttributes=false&templateName=wni_va_rel1_spanish_ie&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=wniesprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://www.irs.gov&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fwww.irs.gov%2Fes%2Frefunds&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/aur_nlp_va_en/index.html?entryPointId=1001&locale=en-US&postChatAttributes=false&templateName=aur_nlp_va_en&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=wniesprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=https://va.connect.irs.gov/assistantIMG/Zoe/Emotions/neutral2.gif&providerId=186A7&wsname=https://www.irs.gov&egChatWindowState=false&VASessionId=6361c244-cbfe-43c0-9438-9227dcf03f80&VAActive=true&VAEscalated=null&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fwww.irs.gov%2Findividuals%2Funderstanding-your-cp3219a-notice%3Futm_source%3DOTC%26utm_medium%3DMail%26utm_term%3Dcp3219a%26utm_campaign%3DNotices&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/sbse_campus_exam/index.html?entryPointId=1001&locale=en-US&postChatAttributes=false&templateName=sbse_campus_exam&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=sbsecampusexamenprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://www.irs.gov&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fwww.irs.gov%2Findividuals%2Funderstanding-your-566-t-letter&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/irsjobsva/index.html?entryPointId=1001&locale=en-US&postChatAttributes=false&templateName=irsjobsva&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=jobsenprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://www.jobs.irs.gov&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fwww.jobs.irs.gov%2Fcareers&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/tas_va/index.html?entryPointId=1002&locale=en-US&postChatAttributes=false&templateName=tas_va&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=tasenprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://www.taxpayeradvocate.irs.gov&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fwww.taxpayeradvocate.irs.gov%2F&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/tas_va_spanish/index.html?entryPointId=1002&locale=es-ES&postChatAttributes=false&templateName=tas_va_spanish&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=tasesprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://es.taxpayeradvocate.irs.gov&EGAIN_AV_CHAT_STATE_DATA=null&parentLost=false&referer=https%3A%2F%2Fes.taxpayeradvocate.irs.gov%2F&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/ola_va/index.html?entryPointId=1001&locale=en-US&postChatAttributes=false&templateName=ola_va&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=olaenprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=https://va.connect.its.gov/assistantIMG/Zoe/Emotions/neutral2.gif&providerId=186A7&wsname=https://sa.www4.irs.gov&EGAINAVCHATSTATEDATA=nul1&parentLost=false&referer=https%3A$2F$2Fsa.www4.irs.gov%2Fola$2Fpaymentoptions&useCustomButton=false&storage=true&docked=true",
    "https://connect.irs.gov/system/templates/chat/ola_va_spanish/index.html?entryPointId=1001&locale=es-ES&postChatAttributes=false&templateName=ola_va_spanish&ver=v11&VAEnabled=true&vaChatEntryPointId=&vaChatServerURL=&VATenantAccId=TMPROD10067889&VATenantToken=TMPROD10067889&VAName=olaesprod&ShowPreChatOnEscalation=&serverURL=https://connect.irs.gov/system&vaLastAvatar=&providerId=186A7&wsname=https://sa.www4.irs.gov&EGAINAVCHATSTATEDATA=null&parentLost=false&referer=https%3A$2F$2Fsa.www4.irs.gov%2Fola%2Fes%2F&useCustomButton=false&storage=true&docked=true",
    "https://groundwork8.egain.cloud/status?link=best&hostName=connect.irs.gov"
)

# Browser Paths
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$firefoxPath = "C:\Program Files\Mozilla Firefox\firefox.exe"

function Open-AllURLsInNewWindow {
    param (
        [string]$browserPath,
        [string[]]$urls
    )
    try {
        $urlsArgument = $urls -join " "
        Start-Process $browserPath -ArgumentList "--new-window", $urlsArgument
    } catch {
        Write-Host "Failed to open URLs - Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

if (Test-Path $chromePath) {
    Open-AllURLsInNewWindow -browserPath $chromePath -urls $urls
} elseif (Test-Path $edgePath) {
    Open-AllURLsInNewWindow -browserPath $edgePath -urls $urls
} elseif (Test-Path $firefoxPath) {
    Open-AllURLsInNewWindow -browserPath $firefoxPath -urls $urls
} else {
    Write-Host "No supported browser found." -ForegroundColor Red
}

# Function to test URL with retry and timestamp
function Test-UrlStatus {
    param ([string]$url)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10
        if ($response.StatusCode -eq 200) {
            return "$timestamp - OK"
        }
    } catch {
        Start-Sleep -Seconds 2
        try {
            $retry = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10
            if ($retry.StatusCode -eq 200) {
                return "$timestamp - OK (Retry)"
            }
        } catch {
            return "$timestamp - NOT OK"
        }
    }
    return "$timestamp - NOT OK"
}

# Check all URLs
$urlResults = @()
foreach ($url in $urls) {
    Write-Host "Checking $url..."
    $status = Test-UrlStatus -url $url
    $urlResults += [PSCustomObject]@{
        URL = $url
        Status = $status
    }
}

# GUI: Display URL Status
$form = New-Object System.Windows.Forms.Form
$form.Text = "IRS URL Health Check Status"
$form.Size = New-Object System.Drawing.Size(1100, 600)
$form.StartPosition = "CenterScreen"

$listView = New-Object System.Windows.Forms.ListView
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Dock = 'Fill'
$listView.Columns.Add("URL", 750)
$listView.Columns.Add("Status with Timestamp", 300)

foreach ($result in $urlResults) {
    $item = New-Object System.Windows.Forms.ListViewItem($result.URL)
    $item.SubItems.Add($result.Status)
    if ($result.Status -like "*NOT OK*") {
        $item.ForeColor = "Red"
    } elseif ($result.Status -like "*Retry*") {
        $item.ForeColor = "DarkOrange"
    } else {
        $item.ForeColor = "Green"
    }
    $listView.Items.Add($item) | Out-Null
}

$form.Controls.Add($listView)
$form.ShowDialog()

# Loop to keep terminal open
while ($true) {
    $input = Read-Host "`nType 'exit' to close"
    if ($input -eq "exit") { break }
}
