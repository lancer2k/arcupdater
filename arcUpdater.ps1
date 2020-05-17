Add-Type -AssemblyName PresentationCore,PresentationFramework

function  newVersionAtRepo($repoUrl,$file,$pattern,$datePattern,$localVersionDate){
    #Get Date from last version of the file in the repo
    $r = Invoke-WebRequest -Uri $repoUrl
    $content = $r.ParsedHtml.getElementsByTagName("pre")[0].innerText
 
    $result = [Regex]::Matches($content,$pattern)
    $lastVersionDateString = $result.Groups[1].Value

    $lastVersionDate = [datetime]::ParseExact($lastVersionDateString, $datePattern , $systemGlob)

    return $localVersionDate -lt $lastVersionDate;
    
}

function  newVersionAtGit($repoAPIURL, $localVersion){
    #Get Date from last version of the file in the repo
    $repoData = Invoke-WebRequest -Uri $repoAPIURL | ConvertFrom-Json
    return $localVersion -lt $repoData.tag_name;
    
}

$config = get-content arcUpdater.ini | ConvertFrom-Json
$systemGlob = New-Object system.globalization.cultureinfo($config.ArcUpdater.culture)



if (newVersionAtGit $config.ArcUpdater.repositoryAPIURL $config.ArcUpdater.version){
    
    $updateAnswer = [System.Windows.MessageBox]::Show($config.langResources.arcUdaterNewVersionMsgboxQuestion,$config.langResources.arcUdaterNewVersionMsgboxTitle,'YesNoCancel')
    Switch ($updateAnswer) {
        'Yes' {
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
    
    $updateAnswer = [System.Windows.MessageBox]::Show($config.langResources.arcDPSNewVersionMsgboxQuestion,$config.langResources.arcDPSNewVersionMsgboxTitle,'YesNoCancel')
    Switch ($updateAnswer) {
        'Yes' {
            $fileURL = "$repoURL$fileName"
            New-Item -ItemType Directory -Force -Path $localArcDllOldVersionFolderPath
            Move-Item  -path $localArcDllFilePath -destination $localArcDllOldVersionFolderPath -Force
            Invoke-WebRequest -Uri $fileURL -OutFile $localArcDllFilePath
            [System.Windows.MessageBox]::Show($config.langResources.arcDPSUpdatedMsgbox)

        }
        'Cancel' {
            Exit
        }
    }    
}
[System.Diagnostics.Process]::Start("$PSScriptRoot\..\Gw2-64.exe")