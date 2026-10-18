# Powershell# 导入CSV文件
$userlist = Import-Csv "C:\Program Files (x86)\Ivanti\EPM Agent\sdmcache\soft\apps\Scripts\xml\miglist.csv"
$host1 = $env:COMPUTERNAME
$user1 = $userlist | Where-Object { $_.sourceVDI -eq $host1 } 
if ($user1) {
    $username = $user1.account
    $sourceVDI = $user1.sourceVDI
    $destVDI = $user1.destVDI
    if ($user1.status -eq "c" -and $sourceVDI -eq $host1) {
        $userName1 = "styfad\helpdesk2001"
        $passwdFile = "C:\Program Files (x86)\Ivanti\EPM Agent\sdmcache\soft\apps\Scripts\xml\Credential.txt"
        $keyFile = "C:\Program Files (x86)\Ivanti\EPM Agent\sdmcache\soft\apps\Scripts\xml\aes.key"
        $key = Get-Content $keyFile
        $Cred = New-Object -TypeName System.Management.Automation.PSCredential `
            -ArgumentList $userName1, (Get-Content $passwdFile | ConvertTo-SecureString -Key $key)
        $logfiles = "\\10.68.1.200\patch\Report\MigUserLogs"
        New-PSDrive -Credential $Cred -Name "R" -PSProvider "FileSystem" -Root $logfiles -ErrorAction Stop
        function Create-DirectoryIfNotExist {
            param ([string]$Path)
            if (-not (Test-Path -Path $Path)) {
                New-Item -Path $Path -ItemType Directory -Force | Out-Null}
        }
        $userLogDirectory = Join-Path -Path R:\ -ChildPath $username
        Create-DirectoryIfNotExist -Path $userLogDirectory
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $logFile = New-Item -Path "$userLogDirectory\MiguserLog_$timestamp.txt" 
        function Write-Log {
            param ([string]$Message)
            $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $logMessage = "$timestamp - $Message"
            Add-Content -Path $logFile -Value $logMessage
        }
        Write-Log "文件夹新建完成。" 
        #结束Office办公软件进程
$Office_Processes = @("EXCEL","WINWORD","OUTLOOK","POWERPNT","Lync","OneDrive","VISIO","WINPROJ","ONENOTEM")
foreach ($Office_Process in $Office_Processes) {
    $Processes = Get-Process -Name $Office_Process -ErrorAction SilentlyContinue
        if ($Processes) {
                Write-Log "找到进程并结束 $Office_Process"
                Stop-Process -Name $Office_Process -Force -ErrorAction SilentlyContinue}
        else{Write-Log "未找到进程 $Office_Process"
            }
        }
#删除输入法缓存
Remove-Item "C:\Users\$username\AppData\Roaming\Microsoft\InputMethod\Chs\*.tmp"  -Recurse -Force

# USMT 迁移工具准备
$amd64zip = "http://10.68.1.200/soft/apps/Scripts/xml/amd64.zip"
$migXml = "http://10.68.1.200/soft/apps/Scripts/xml/miguser.xml"
$outlook = "http://10.68.1.200/soft/apps/scripts/xml/outlook_config.xml"
$amd64zipLocalpath = "C:\tools\amd64.zip"
$migXmllocalpath = "C:\Tools\miguser.xml"
$outlook_config = "C:\Tools\outlook_config.xml"
$WebClient =New-Object System.Net.WebClient
$webclient.DownloadFile($amd64zip,$amd64zipLocalpath)
$webclient.DownloadFile($migXml,$migXmllocalpath)
$webclient.DownloadFile($outlook,$outlook_config)
Write-Log "下载USMT迁移工具" 

#解压zip文件
Expand-Archive "C:\Tools\amd64.zip" -DestinationPath "C:\Tools\" -Force -ea SilentlyContinue
if(Test-Path "C:\Tools\amd64"){
    Write-Log "文件解压完成"}else{
    Write-Log "未找到zip文件"}
$Session = New-PSSession -ComputerName $destvdi -Credential $Cred
Write-Log $Session.ComputerName,$Session.computertype,$Session.state

#运行scanstate命令写入日志
$Localpath_log = "No error in migration process is recorded."
$destuserdate1 = "\\$destVDI\c$\tools"
Invoke-Command -Session $Session -ScriptBlock {
    param($username,$destuserdate1)
    New-Item -Name $username -Path $destuserdate1 -ItemType Directory -Force
} -ArgumentList $username,$destuserdate1
$destuserdate2 = $destuserdate1 + "\$username" 
Set-Location "C:\Tools\amd64"
New-PSDrive -Name "S" -PSProvider FileSystem -Root $destuserdate2 -Credential $Cred -Persist
$scanstate = .\scanstate.exe S:\  /i:C:\Tools\miguser.xml /i:C:\tools\amd64\MigApp.xml /i:C:\Tools\outlook_config.xml /ui:$username /ue:* /o /v:13 /l:C:\tools\$username.log
Copy-Item C:\tools\$username.log -Destination $userLogDirectory
Copy-Item C:\Tools\MigLog.xml -Destination $userLogDirectory
$Miguser_Log = Select-String -Path $userLogDirectory\$username.log -Pattern  "No error in migration process is recorded." | Select-Object -ExpandProperty Line
if ($Miguser_Log -match "No\s+(.+)"){
$result = "No " + $matches[1]
Write-Log "没有发现错误_$result"
    if($Localpath_log -eq $result){
    Write-Log "IMG_UserData文件迁移完成"
    }else{ Write-Log "IMG_UserData文件迁移失败"}
}else { Write-Log "运行scanstate_data命令失败" }

Set-Location C:\
remove-item C:\tools\amd64 -Recurse -Force -ea SilentlyContinue
Remove-Item C:\tools\amd64.zip -Force -ea SilentlyContinue
Remove-Item C:\Tools\miguser.xml -Force -ea SilentlyContinue
Remove-Item C:\tools\$username.log,C:\Tools\MigLog.xml -Force -ea SilentlyContinue

#远程下载文件
$destvdiamd64zip = "http://10.68.1.200/soft/apps/Scripts/xml/amd64.zip"
$destvdimigXml = "http://10.68.1.200/soft/apps/Scripts/xml/miguser.xml"
$destvdimigXmlapps = "http://10.68.1.200/soft/apps/Scripts/xml/MigApp.xml"
$destoutlook = "http://10.68.1.200/soft/apps/scripts/xml/outlook_config.xml"
$WebClient =New-Object System.Net.WebClient
$webclient.DownloadFile($destvdiamd64zip,"S:\amd64.zip")
$webclient.DownloadFile($destvdimigXml,"S:\miguser.xml")
$webclient.DownloadFile($destvdimigXmlapps,"S:\MigApp.xml")
$webclient.DownloadFile($destoutlook,"S:\outlook_config.xml")
Expand-Archive -Path S:\amd64.zip -DestinationPath S:\ -Force
if(Test-Path "S:\amd64"){
    Write-Log "文件解压完成"}else{
    Write-Log "未找到zip文件"}
Write-Log "远程解压完成"

#loadstate
$destUserPath = "C:\tools\$username"
$loadstatePath = "$destUserPath\amd64\loadstate.exe"
$destvdimigXml_1 = "/i:C:\tools\$username\MigUser.xml"
$destvdimigXml_2 = "/i:C:\tools\$username\amd64\MigApp.xml"
$destvdimigXml_3 = "/i:C:\tools\$username\outlook_config.xml"
$uiusername = "/ui:$username"
$logPath = "$destUserPath\Loadstate.log"

Invoke-Command -Session $Session -ScriptBlock {
param ($destUserPath, $loadstatePath, $destvdimigXml_1,$destvdimigXml_2,$destvdimigXml_3,$uiusername, $logPath)
$process = & $loadstatePath $destUserPath $destvdimigXml_1 $destvdimigXml_2 $destvdimigXml_3 $uiusername "/v:13" "/l:$logPath" 
        } -ArgumentList $destUserPath, $loadstatePath, $destvdimigXml_1,$destvdimigXml_2,$destvdimigXml_3,$uiusername, $logPath -Verbose
Write-Log $Processes
#复制文件到远程目录
Copy-Item S:\Loadstate.log -Destination $userLogDirectory\Loadstate_$username.log
Copy-Item S:\MigLog.xml -Destination  $userLogDirectory\LoadstateMiglog_$username.xml
$MiguserLoadstate_Log = Select-String -Path $userLogDirectory\Loadstate_$username.log -Pattern  "No error in migration process is recorded." | Select-Object -ExpandProperty Line
if ($MiguserLoadstate_Log -match "No\s+(.+)"){
$result1 = "No " + $matches[1]
Write-Log "没有发现错误_$result"
    if($Localpath_log -eq $result1){
    Write-Log "IMG_UserData文件释放完成"
    }else{ Write-Log "IMG_UserData文件释放失败"}
}else { Write-Log "运行Loadstatestate_data命令失败" }


Remove-PSDrive -Name "S","R"
Invoke-Command -Session $Session -ScriptBlock {
param($username)
Remove-Item C:\Tools\$username -Recurse -Force -ea SilentlyContinue
} -ArgumentList $username
Remove-PSSession -Session $Session
}
}
