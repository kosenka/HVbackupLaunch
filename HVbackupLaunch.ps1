param($o,$l)

# simple function for logging to screen/logfile
function LogWrite($msg) {
	Add-Content $log ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + " " + $msg)
	Write-Output ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + " " + $msg)
}

#echo "# STEP 1. Read parameters from command line"
if (!$o -or !$l) {
	echo "HVbackup Launcher`n`t
	Usage:`n`t
	HVbackupLaunch -o <path/to/backup/folder> -l <list,of,virtuals,machines>`n`n"
	exit(0)
}
if (!(Test-Path $o)) {
	echo("Backup folder {0} not exist.`r`n" -f $p)
	exit(0)
}

#echo "# STEP 1.2. check for 7zip"
if (-not (test-path "$env:ProgramFiles\7-Zip\7z.exe"))
{
	throw "$env:ProgramFiles\7-Zip\7z.exe needed"
	exit(0)
}


#echo "# STEP 2. Read parameters from ini-file and set default values"
$inifile = Join-Path $PSScriptRoot ( $MyInvocation.MyCommand.Name.Replace("ps1", "ini") )
if (!(Test-Path $inifile)) {
	echo("INI-file not found ({0}).`r`nRTFM`r`n" -f $inifile)
	exit(0)
}

$ini = ConvertFrom-StringData((Get-Content $inifile) -join "`n")
if (!($ini.ContainsKey("LOGPATH"))) {
	$ini.Add("LOGPATH", $PSScriptRoot) # whereis this script
}

if (!($ini.ContainsKey("RPTPATH"))) {
	$ini.Add("RPTPATH", $PSScriptRoot) # whereis this script
}

if (!($ini.ContainsKey("LOGSTORETIME"))) {
	echo("Set LOGSTORETIME parameter in {0}.`r`nRecomended (days):`r`n" -f $inifile)
	echo("`t LOGSTORETIME=30")
	exit(0)
}

if (!($ini.ContainsKey("KEEPBACKUPS"))) {
	echo("Set KEEPBACKUPS parameter in {0}.`r`nRecomended minimum:`r`n" -f $inifile)
	echo("`t KEEPBACKUPS=1")
	exit(0)
}

#echo "# STEP 3. Remove old logfiles (YYYYMMDD.log)"
Get-ChildItem $ini["LOGPATH"] | Where-Object {$_.Name -match "^\d{8}.log$"}`
	| ? {$_.PSIsContainer -eq $false -and $_.lastwritetime -lt (get-date).adddays(-$ini["LOGSTORETIME"])}`
	| Remove-Item -Force

#echo "# STEP 4. Create logfile"
$log = Join-Path $ini["LOGPATH"] ( (Get-Date).ToString('yyyyMMdd') + ".log" )
if (!(Test-Path $log)) {
	New-Item -type file $log -force
}

# and temporary report file
$report = Join-Path $ini["RPTPATH"] "report.txt"
New-Item -type file $report -force

LogWrite("`tStart")
LogWrite("`tBackup {0} virtual machine(s) ({1}) to {2}"`
	-f $l.count, ($l -join ', '), $p)

#echo "# STEP 5. Remove very old backups"
foreach ($item in $l) {
	$files = Get-ChildItem -Path (Join-Path $o ($item + "_*.*")) | sort desc
	for ($j=[int]$ini["KEEPBACKUPS"]; $j -lt $files.Count; $j++)
	{
		LogWrite("`tDelete {0}" -f $files[$j].FullName)
		Remove-Item $files[$j].FullName
	}
}


#echo "# STEP 6. Backup for every virtual machine"
$totalSuccess = 0
$totalSize = 0
$msgSummary = ''
foreach ($item in $l) {

	LogWrite("`tStopping virtual machine: {0}" -f $item.ToUpper())
	Stop-VM $item # Stop virtual machine
	LogWrite("`tStopped virtual machine: {0}" -f $item.ToUpper())

	LogWrite("`tGetting Virtual machine: {0}" -f $item.ToUpper())
	$Vms = Get-VM -Name $item # Get virtual machine by name

	$vmVhd=$Vms.HardDrives.Path # Get virtual machine hard drive path
	LogWrite("`tGetting Virtual machine (VHD): {0}" -f $vmVhd)
	
	$7z = Join-Path $o ($item + '_' + (get-date -format yyyyMMdd) + '.7z')

	$cmd =  $ini["HVBACKUPEXE"] + ' ' + $7z + ' "' + $vmVhd + '" 2>&1'
	LogWrite("`tBackup {0}. Run: {1}`n" -f $item.ToUpper(), $cmd)
	$cmdResult = invoke-expression $cmd
	LogWrite($cmdResult | out-string)
	LogWrite("`tBackup complete`n")

	# get summary
	if (Test-Path $7z)
	{
		$totalSuccess += 1
		$totalSize += ((Get-Item $7z).length/1GB)
		$msgSummary += $7z + "`t" + ("{0:N1}" -f ((Get-Item $7z).length/1GB)) + "Gb`t`n"
	}

	LogWrite("`tStarting virtual machine: {0}" -f $item.ToUpper())
	Start-VM –Name $item # Start virtual machine
	LogWrite("`tStarted virtual machine: {0}" -f $item.ToUpper())
}

$msgSubject = ('HVbackup {0}. Success: {1}/{2}, Size: {3:N1} Gb'`
	-f (Get-Item env:\Computername).Value, $totalSuccess, $l.Count, $totalSize)
Add-Content $report ($msgSubject + "`n`n" + $msgSummary)

LogWrite("`tSuccessful Stop")

#echo ("# STEP 6. (optional) Send summary report")

if ($ini.ContainsKey("MAILADDRESS") -and $ini.ContainsKey("MAILSERVER"))  {
	$msg = New-Object Net.Mail.MailMessage
	$msg.from = $ini["MAILUSER"]
	$msg.to.add($ini["MAILADDRESS"])
	$msg.Subject = $msgSubject
	$msg.Body = ((Get-Content $report) -join "`n")`
		 + "`n----------------------- Detailed Log -------------------------------`n`n"`
		 + ((Get-Content $log) -join "`n")

	$ini["MAILSERVERPORT"] = "25"
	if ($ini["MAILSERVER"].Contains(":")) {
		$mailserver = $ini["MAILSERVER"].Split(":")
		$ini["MAILSERVER"] = $mailserver[0]
		$ini["MAILSERVERPORT"] = $mailserver[1]
	}

	$smtp = New-Object Net.Mail.SmtpClient($ini["MAILSERVER"],$ini["MAILSERVERPORT"])
	$smtp.EnableSSL = $true
	if ($ini.ContainsKey("MAILUSER") -and $ini.ContainsKey("MAILPASSWORD"))  {
		$smtp.Credentials = New-Object System.Net.NetworkCredential($ini["MAILUSER"], $ini["MAILPASSWORD"]);
	}
	$smtp.Send($msg)
	LogWrite("`t... sent summary")
}

