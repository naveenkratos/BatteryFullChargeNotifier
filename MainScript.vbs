set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
for each oResult in oResults
   iFull = oResult.FullChargedCapacity
next

RepeatTimeInMin = 0.02 ' For flashing notification

while (1)
  set oResults = oServices.ExecQuery("select * from batterystatus")
  for each oResult in oResults
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next
  iPercent = ((iRemaining / iFull) * 100) mod 100
  if bCharging and (iPercent >= 95) Then ' Run PowerShell Script to Show Battery fully charged remove charger notification banner 
	Set shell = CreateObject("WScript.Shell")

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	' Get the folder where this VBScript is located
	scriptFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

	' Build the full path to the PowerShell script
	ps1File = scriptFolder & "\Notify.ps1"
	
    ' Build PowerShell command to run notify.ps1 with parameters
    psCommand = "powershell -NoProfile -ExecutionPolicy Bypass -File """ & ps1File & """ " & RepeatTimeInMin & " " & iPercent
	
    ' Run PowerShell silently (0 = hidden window, False = no wait)
    shell.Run psCommand, 0, False
  end if
  wscript.sleep RepeatTimeInMin*60000 ' milliseconds

wend
