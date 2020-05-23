Add-Type -AssemblyName PresentationCore,PresentationFramework

function  newVersionAtRepo($repoUrl,$file,$pattern,$datePattern,$localVersionDate){
    #Get Date from last version of the file in the repo
    $objSumaryLabel.Text += "`nChecking for ArcDPS updates ... "
    
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
    $objSumaryLabel.Text += "`nChecking for ArcUpdater updates ... "
    
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

$config = get-content arcUpdater.ini | ConvertFrom-Json
$systemGlob = New-Object system.globalization.cultureinfo($config.ArcUpdater.culture)

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

if (newVersionAtGit $config.ArcUpdater.repositoryAPIURL $config.ArcUpdater.version){
    $objSumaryLabel.Text += "`nNew ArcUpdate version found ... "
    $updateAnswer = [System.Windows.MessageBox]::Show($config.langResources.arcUdaterNewVersionMsgboxQuestion,$config.langResources.arcUdaterNewVersionMsgboxTitle,'YesNoCancel')
    Switch ($updateAnswer) {
        'Yes' {
            $objSumaryLabel.Text += "`nUPDATING"
            $repoData = Invoke-WebRequest -Uri $config.ArcUpdater.repositoryAPIURL | ConvertFrom-Json
            Invoke-WebRequest -Uri $repoData.zipball_url -OutFile $config.ArcUpdater.fileName
            Expand-Archive -Path $config.ArcUpdater.fileName -DestinationPath "." -Force
            Remove-Item $config.ArcUpdater.fileName
            Move-Item  -path "lancer2k-arcupdater-*" -destination "lancer2k-arcupdater" -Force
            Move-Item  -path "lancer2k-arcupdater\*" -destination "." -Force
            Remove-Item "lancer2k-arcupdater" -Recurse
            [System.Windows.MessageBox]::Show($config.langResources.arcUdaterUpdatedMsgbox)
        }
        'Cancel' {
            $objForm.Close()
            Exit
        }
    }
}



$repoURL = $config.ArcDPS.repositoryURL;
$fileName = $config.ArcDPS.fileName;
$pattern =  $config.ArcDPS.htmlPattern;
$datePattern =  $config.ArcDPS.datePattern;
$oldVersionFolder = $config.ArcUpdater.oldVersionPath;


$localArcDllFolderPath = "$PSScriptRoot\..\bin64\"
$localArcDllOldVersionFolderPath = "$PSScriptRoot\$oldVersionFolder\"
$localArcDllFilePath = "$localArcDllFolderPath$fileName"
$localVersionDate = (Get-Item $localArcDllFilePath).LastWriteTime 

if (newVersionAtRepo $repoURL $fileName $pattern $datePattern $localVersionDate){
    $objSumaryLabel.Text += "`nNew ArcDPS Version found ... "
    $updateAnswer = [System.Windows.MessageBox]::Show($config.langResources.arcDPSNewVersionMsgboxQuestion,$config.langResources.arcDPSNewVersionMsgboxTitle,'YesNoCancel')
    Switch ($updateAnswer) {
        'Yes' {
            $objSumaryLabel.Text += "`nUPDATING"
            $fileURL = "$repoURL$fileName"
            New-Item -ItemType Directory -Force -Path $localArcDllOldVersionFolderPath
            Move-Item  -path $localArcDllFilePath -destination $localArcDllOldVersionFolderPath -Force
            Invoke-WebRequest -Uri $fileURL -OutFile $localArcDllFilePath
            [System.Windows.MessageBox]::Show($config.langResources.arcDPSUpdatedMsgbox)

        }
        'Cancel' {
            $objForm.Close()
            Exit
        }
    }    
}
$objSumaryLabel.Text += "`nLoading Guild Wars 2 ... "
Start-Sleep -Seconds 1.5
$objForm.Show()
[System.Diagnostics.Process]::Start("$PSScriptRoot\..\Gw2-64.exe")
