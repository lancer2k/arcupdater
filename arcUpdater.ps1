Add-Type -AssemblyName PresentationCore,PresentationFramework

function  newVersionAtRepo($repoUrl,$file,$pattern,$datePattern,$localVersionDate){
    #Get Date from last version of the file in the repo
    $objSumaryLabel.Text += "`n"
    $objSumaryLabel.Text += $global:config.langResources.arcDPSCheckingNewVersionMessage
    
    try { 
        $r = Invoke-WebRequest -Uri $repoUrl
        $content = $r.ParsedHtml.getElementsByTagName("pre")[0].innerText
    
        $result = [Regex]::Matches($content,$pattern)
        $lastVersionDateString = $result.Groups[1].Value

        $lastVersionDate = [datetime]::ParseExact($lastVersionDateString, $datePattern , $systemGlob)
        $objSumaryLabel.Text += "OK"
        return $localVersionDate -lt $lastVersionDate;
       
    }     
    catch {
        $objSumaryLabel.Text += "Error"
        return $false
    }
    
}

function  newVersionAtGit($repoAPIURL, $localVersion){
    #Get Date from last version of the file in the repo
    $objSumaryLabel.Text += "`n"
    $objSumaryLabel.Text += $global:config.langResources.arcUdaterCheckingNewVersionMessage
    
    try { 
        $repoData = Invoke-WebRequest -Uri $repoAPIURL | ConvertFrom-Json  
        $objSumaryLabel.Text += "OK"
        return $localVersion -lt $repoData.tag_name
    }     
    catch {
        $objSumaryLabel.Text += "Error"
        return $false
    }
    
    
}

function loadGuildWars2(){
    $objSumaryLabel.Text += "`n"
    $objSumaryLabel.Text += $global:config.langResources.loadingGWMessage
    Start-Sleep -Seconds 1.5
    $objForm.Show()
    [System.Diagnostics.Process]::Start("$PSScriptRoot\..\Gw2-64.exe")
}

function Test-ControlKey
{
  # key code for Ctrl key:
  $key = 17    
    
  # this is the c# definition of a static Windows API method:
  $Signature = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
    public static extern short GetAsyncKeyState(int virtualKeyCode); 
'@

  Add-Type -MemberDefinition $Signature -Name Keyboard -Namespace PsOneApi
  [bool]([PsOneApi.Keyboard]::GetAsyncKeyState($key) -eq -32767)
}

function ArcDPSCanceling() {
    $loadArcDPS = $true
    Write-Host $global:config.langResources.arcDPSCancelingMessage -NoNewline
    $timeout = new-timespan -Seconds $global:config.ArcUpdater.AvoidLoadingArcDPSSeconds
    $sw = [diagnostics.stopwatch]::StartNew()

    while ($sw.elapsed -lt $timeout){
        Write-Host '.' -NoNewline
        $pressed = Test-ControlKey
        if ($pressed) {
            $loadArcDPS = $false
            break
        }        
        # write a dot and wait a second
        Start-Sleep -Seconds 1
    } 
    
    return $loadArcDPS
    
}

$global:config = get-content arcUpdater.ini | ConvertFrom-Json
$systemGlob = New-Object system.globalization.cultureinfo($global:config.ArcUpdater.culture)

$loadArcDPS = ArcDPSCanceling

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$objForm = New-Object System.Windows.Forms.Form
$objForm.StartPosition = "CenterScreen"
$objForm.ShowIcon = $False
$objForm.Width = 220
$objForm.Height = 150
$objForm.BackColor = "Black"
$objForm.Opacity = 0.85
$objForm.FormBorderStyle = 'None'


$objImage = [system.drawing.image]::FromFile(".\GW2Logo.png")
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Size(1,1)
$pictureBox.Image = $objImage
$pictureBox.Width = $objImage.Size.Width
$pictureBox.Height = $objImage.Size.Height
$objForm.Controls.Add($pictureBox)


$objTitleLabel = New-Object System.Windows.Forms.label
$objTitleLabel.Location = New-Object System.Drawing.Size(120,3)
$objTitleLabel.Size = New-Object System.Drawing.Size(130,15)
$objTitleLabel.BackColor = "Transparent"
$objTitleLabel.ForeColor = "White"
$objTitleLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$objTitleLabel.Text = "ArcUpdater"
$objForm.Controls.Add($objTitleLabel)

$objSumaryLabel = New-Object System.Windows.Forms.label
$objSumaryLabel.Location = New-Object System.Drawing.Size(7,25)
$objSumaryLabel.Size = New-Object System.Drawing.Size(200,150)
$objSumaryLabel.BackColor = "Transparent"
$objSumaryLabel.ForeColor = "White"
$objSumaryLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8)
$objForm.Controls.Add($objSumaryLabel)

$objForm.Show()

if (newVersionAtGit $global:config.ArcUpdater.repositoryAPIURL $global:config.ArcUpdater.version){
    $objSumaryLabel.Text += "`n"
    $objSumaryLabel.Text += $global:config.langResources.arcUdaterNewVersionMessage
    $updateAnswer = [System.Windows.MessageBox]::Show($global:config.langResources.arcUdaterNewVersionMsgboxQuestion,$global:config.langResources.arcUdaterNewVersionMsgboxTitle,'YesNoCancel')
    Switch ($updateAnswer) {
        'Yes' {
            $objSumaryLabel.Text += "`n"
            $objSumaryLabel.Text += $global:config.langResources.updatingMessage
            $repoData = Invoke-WebRequest -Uri $global:config.ArcUpdater.repositoryAPIURL | ConvertFrom-Json
            Invoke-WebRequest -Uri $repoData.zipball_url -OutFile $global:config.ArcUpdater.fileName
            Expand-Archive -Path $global:config.ArcUpdater.fileName -DestinationPath "." -Force
            Remove-Item $global:config.ArcUpdater.fileName
            Move-Item  -path "lancer2k-arcupdater-*" -destination "lancer2k-arcupdater" -Force
            Move-Item  -path "lancer2k-arcupdater\*" -destination "." -Force
            Remove-Item "lancer2k-arcupdater" -Recurse
            [System.Windows.MessageBox]::Show($global:config.langResources.arcUdaterUpdatedMsgbox)
        }
        'Cancel' {
            $objForm.Close()
            Exit
        }
    }
}


$repoURL = $global:config.ArcDPS.repositoryURL;
$fileName = $global:config.ArcDPS.fileName;
$pattern =  $global:config.ArcDPS.htmlPattern;
$datePattern =  $global:config.ArcDPS.datePattern;
$oldVersionFolder = $global:config.ArcUpdater.oldVersionPath;


$localArcDllFolderPath = "$PSScriptRoot\..\bin64\"
$localArcDllOldVersionFolderPath = "$PSScriptRoot\$oldVersionFolder\"

$localArcDllFilePath = "$localArcDllFolderPath$fileName"
$extraCharacter = "_"
$localArcDllAlternativeFilePath = "$localArcDllFolderPath$fileName$extraCharacter"

if ($loadArcDPS){

    if (Test-Path $localArcDllAlternativeFilePath){
        Move-Item -path $localArcDllAlternativeFilePath -destination $localArcDllFilePath -Force
    }

    $downloadNewVersion = $true
    if (Test-Path $localArcDllFilePath){
        $localVersionDate = (Get-Item $localArcDllFilePath).LastWriteTime 
        $downloadNewVersion = newVersionAtRepo $repoURL $fileName $pattern $datePattern $localVersionDate
    }



    if ( $downloadNewVersion ){
        $objSumaryLabel.Text += "`n"
        $objSumaryLabel.Text += $global:config.langResources.arcDPSNewVersionMessage
        $updateAnswer = [System.Windows.MessageBox]::Show($global:config.langResources.arcDPSNewVersionMsgboxQuestion,$global:config.langResources.arcDPSNewVersionMsgboxTitle,'YesNoCancel')
        Switch ($updateAnswer) {
            'Yes' {
                $objSumaryLabel.Text += "`n"
                $objSumaryLabel.Text += $global:config.langResources.updatingMessage
                $fileURL = "$repoURL$fileName"
                New-Item -ItemType Directory -Force -Path $localArcDllOldVersionFolderPath
                Move-Item  -path $localArcDllFilePath -destination $localArcDllOldVersionFolderPath -Force
                Invoke-WebRequest -Uri $fileURL -OutFile $localArcDllFilePath
                [System.Windows.MessageBox]::Show($global:config.langResources.arcDPSUpdatedMsgbox)

            }
            'Cancel' {
                $objForm.Close()
                Exit
            }
        }    
    }
}else{
    $objSumaryLabel.Text += "`n"
    $objSumaryLabel.Text += $global:config.langResources.avoidArcDPSLoadingMessage
    $objForm.Show()
    if (Test-Path $localArcDllFilePath){
        Move-Item  -path $localArcDllFilePath -destination $localArcDllAlternativeFilePath -Force
    }
}

loadGuildWars2

