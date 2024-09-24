$ScheduledTaskContent = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2024-09-24T15:11:08.3290381</Date>
    <Author>Automox</Author>
    <URI>\Perform Regular Maintainence Reboot</URI>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>2024-09-24T03:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Sunday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\Windows\System32\shutdown.exe</Command>
      <Arguments>/r /t 5 /f</Arguments>
      <WorkingDirectory>C:\Windows\System32</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
'@

[System.IO.FileInfo]$ScheduledTaskExportPath = "$($Env:Temp.TrimEnd('\'))\Automox\ScheduledTasks\WeeklyReboot.xml"

Switch ($ScheduledTaskExportPath.Directory.Exists)
    {
        {($_ -eq $False)}
            {
                $Null = [System.IO.Directory]::CreateDirectory($ScheduledTaskExportPath.Directory.FullName)
            }
    }

$ScheduledTaskEncoding = [System.Text.Encoding]::Default

$Null = [System.IO.File]::WriteAllText($ScheduledTaskExportPath.FullName, $ScheduledTaskContent, $ScheduledTaskEncoding)

$ScheduledTaskName = "Perform Regular Maintainence Reboot"

$ImportScheduledTaskResult = Start-Process -FilePath 'schtasks.exe' -ArgumentList "/Create /XML `"$($ScheduledTaskExportPath.FullName)`" /tn `"$($ScheduledTaskName)`"" -PassThru -WindowStyle Hidden -Wait

$ImportScheduledTaskResult.ExitCode
