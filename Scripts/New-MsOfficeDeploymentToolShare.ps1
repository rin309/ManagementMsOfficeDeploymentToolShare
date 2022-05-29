#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
 .Synopsis
   Assists in ODT placement on the server

 .Description
   The task schedule automatically deletes the old version.

 .Parameter ConfigPath
   Set ODT config path
   Create from 'Microsoft 365 Apps admin center'

   https://config.office.com/

 .Parameter LocalOfficeDeploymentToolPath
   Set to officedeploymenttool_*.exe path
   Available from 'Official Microsoft Download Center'

   https://go.microsoft.com/fwlink/p/?LinkID=626065

 .Parameter DestinationDirectory
   Root folder to place shared folders

 .Parameter ShareIsHidden
   Add '$' to the end of the created shared folder name flag

 .Parameter DirectoryName
   Deployment directory name and shared folder name

 .Parameter NoRegisterTask
   Not register task schedule flag

 .Parameter TaskName
   Task name

 .Parameter TaskTrigger
   Task trigger

 .Parameter StartTask
   Start registered task

 .Parameter AcceptMsOdtEula
   Agree to EULA (MICROSOFT SOFTWARE LICENSE TERMS, MICROSOFT OFFICE DEPLOYMENT TOOL) and do not display prompt
   If you want to check the contents of EULA, please execute ODT directly.

 .Parameter Force
   Ignore errors that can be continued

 .Example
   # Simply command sample
   New-MsOfficeDeploymentToolShare -ConfigPath $env:UserProfile\Downloads\Configuration.xml -LocalOfficeDeploymentToolPath $env:UserProfile\Downloads\officedeploymenttool_15128-20224.exe

 .Example
   # Select additional option
   New-MsOfficeDeploymentToolShare -ConfigPath $env:UserProfile\Downloads\Configuration.xml -LocalOfficeDeploymentToolPath $env:UserProfile\Downloads\officedeploymenttool_15128-20224.exe -DestinationDirectory D:\Shares -DirectoryName PerpetualVL2021 -Force

 .Example
   # Install to running command sample by command prompt
   PowerShell -ExecutionPolicy ByPass -Command "Import-Module ManagementMsOfficeDeploymentToolShare; New-MsOfficeDeploymentToolShare -ConfigPath $env:UserProfile\Downloads\Configuration.xml -LocalOfficeDeploymentToolPath $env:UserProfile\Downloads\officedeploymenttool_15128-20224.exe-ConfigPath \""$env:UserProfile\Downloads\Configuration.xml\"" -LocalOfficeDeploymentToolPath \""$env:UserProfile\Downloads\officedeploymenttool_15128-20224.exe\"" -Force"

#>
Function New-MsOfficeDeploymentToolShare{
    Param(
        [Parameter(Mandatory)][String]$ConfigPath,
        [Parameter(Mandatory)][String]$LocalOfficeDeploymentToolPath,
        [String]$DestinationDirectory = "C:\Shares",
        [Bool]$ShareIsHidden = $True,
        [String]$DirectoryName = "Office",
        [Switch]$NoRegisterTask = $False,
        [String]$TaskName = "Invoke-MsOfficeDeploymentToolAndCompress",
        $TaskTrigger,
        [Switch]$StartTask = $False,
        [Switch]$AcceptMsOdtEula = $False,
        [Switch]$AcceptEula = $False,
        [Switch]$Force
    )

    Function Invoke-Application($Path, $Argument){
        Try{
            $Process = Start-Process -FilePath $Path -ArgumentList $Argument -PassThru -Wait
            If ($Process.ExitCode -ne 0){
                Write-Warning "$($NewMsOfficeDeploymentToolShareMessageTable.WarnExitCodeByInvokeApplication) [$Path $Argument] : $($Process.ExitCode)" -Verbose
            }
        }
        Catch{
            Write-Warning "$($NewMsOfficeDeploymentToolShareMessageTable.WarnByInvokeApplication) [$Path $Argument] : $($_.Exception.Message)"
        }
    }
    Function Invoke-DownloadToTemporaryDirectory($Url){
        Try{
            # support Windows Server 2016
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

            $Path = $((New-TemporaryFile).FullName)
            Invoke-WebRequest $Url -OutFile $Path
            Return $Path
        }
        Catch{
            Write-Warning "$($NewMsOfficeDeploymentToolShareMessageTable.WarnByInvokeDownloadToTemporaryDirectory) [$Url] : $($_.Exception.Message)"
        }
    }
    Function Invoke-WriteError($Exception, $Force){
        If ($Force){
            Write-Warning $Exception
        }
        Else{
            Write-Error $Exception -ErrorAction Stop
        }
    }
    Try{
        # MS-EULA
        If (-not $AcceptMsOdtEula){
            $ChoiceDescription = [System.Management.Automation.Host.ChoiceDescription]
            $MessageOptions = @(
                New-Object $ChoiceDescription ($NewMsOfficeDeploymentToolShareMessageTable.AgreeLable,"")
                New-Object $ChoiceDescription ($NewMsOfficeDeploymentToolShareMessageTable.ExitLable, "")
            )
            $MessageResult = $Host.Ui.PromptForChoice($NewMsOfficeDeploymentToolShareMessageTable.MsOdtEulaLabel, $NewMsOfficeDeploymentToolShareMessageTable.MsOdtEulaMessage, $MessageOptions, 0)
            Switch ($MessageResult){
                1{
                    Return
                }
            }
        }

        # check param
        If (-Not (Test-Path $ConfigPath -PathType Leaf)){
            Write-Error ([System.IO.FileNotFoundException]::new("$($NewMsOfficeDeploymentToolShareMessageTable.NotFoundConfigPath): [$ConfigPath]")) -ErrorAction Stop
        }
        If (-Not (Test-Path $DestinationDirectory -PathType Container)){
            Invoke-WriteError -Exception ([System.IO.DirectoryNotFoundException]::new("$($NewMsOfficeDeploymentToolShareMessageTable.NotFoundDestinationDirectory) [$DestinationDirectory]")) -Force $Force
            If ($Force){
                New-Item -Path $DestinationDirectory -ItemType Directory -Force | Out-Null
            }
        }
        If (-not [String]::IsNullOrEmpty($LocalOfficeDeploymentToolPath) -and -Not (Test-Path $LocalOfficeDeploymentToolPath -PathType Leaf)){
            Write-Error ([System.IO.FileNotFoundException]::new("$($NewMsOfficeDeploymentToolShareMessageTable.NotFoundLocalOfficeDeploymentToolPath) [$LocalOfficeDeploymentToolPath]")) -ErrorAction Stop
        }
        If ([String]::IsNullOrEmpty($DirectoryName)){
            $Configuration = ([Xml](Get-Content $ConfigPath -Encoding UTF8)).Configuration
            $DirectoryName = "$($Configuration.Add.Channel).$($Configuration.Add.OfficeClientEdition)"
        }

        If ($ShareIsHidden){
            $DirectoryName += "$"
        }

        If ((Get-SmbShare | Where-Object Name -eq $DirectoryName).Count -ne 0){
            Invoke-WriteError -Exception "$($NewMsOfficeDeploymentToolShareMessageTable.ExistsFromShares): [$DirectoryName]" -Force $Force
        }
        $DestinationDirectory = (Join-Path $DestinationDirectory $DirectoryName)

        If (-not $NoRegisterTask){
            $ScheduleService = New-Object -ComObject Schedule.Service
            $ScheduleService.Connect()
            $ScheduleTask = $ScheduleService.GetFolder("\").GetTasks(0) | Where-Object Name -eq $TaskName
            If ($ScheduleTask -ne $Null){
                Invoke-WriteError -Exception "$($NewMsOfficeDeploymentToolShareMessageTable.ExistsFromTaskSchedule): [$TaskName]" -Force $Force
                If ($Force){
                    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$False
                }
            }
        }

        # download ODT installer
        If ([String]::IsNullOrEmpty($LocalOfficeDeploymentToolPath)){
            $Path = Invoke-DownloadToTemporaryDirectory -Url "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_15128-20224.exe"
            If (-not (Test-Path $Path -PathType Leaf)){
                Write-Error ([System.IO.FileNotFoundException]::new("$($NewMsOfficeDeploymentToolShareMessageTable.NotFoundLocalOfficeDeploymentToolPath): [$LocalOfficeDeploymentToolPath]")) -ErrorAction Stop
            }
            $LocalOfficeDeploymentToolPath = "$Path.exe"
            Move-Item $Path -Destination $LocalOfficeDeploymentToolPath
        }

        # test ODT's digital signature (If it is digitally signed, it is judged that the file is not damaged.)
        If ((Get-AuthenticodeSignature $LocalOfficeDeploymentToolPath).Status -ne "Valid"){
            Invoke-WriteError -Exception "$($NewMsOfficeDeploymentToolShareMessageTable.InvalidDigitalSignature): [$LocalOfficeDeploymentToolPath]" -Force $Force
        }

        # extract ODT installer
        $TemporaryDirectoryPath = (New-TemporaryFile).FullName
        Remove-Item $TemporaryDirectoryPath -Force

        Invoke-Application -Path $LocalOfficeDeploymentToolPath -Argument "/quiet /extract:$TemporaryDirectoryPath"
        $TemporaryMsOdtPath = (Join-Path $TemporaryDirectoryPath "setup.exe")

        If (-not (Test-Path $TemporaryMsOdtPath)){
            Write-Error ([System.IO.FileNotFoundException]::new("$($NewMsOfficeDeploymentToolShareMessageTable.NotFoundTemporaryMsOdtPath): [$TemporaryMsOdtPath]")) -ErrorAction Stop
        }

        # Copy to Directory
        New-Item $DestinationDirectory -ItemType Container -Force | Out-Null

        $DestinationConfigPath = (Join-Path $DestinationDirectory (Split-Path $ConfigPath -Leaf))
        Copy-Item $ConfigPath $DestinationConfigPath
        $DestinationOdtPath = (Join-Path $DestinationDirectory "setup.exe")
        Copy-Item $TemporaryMsOdtPath (Join-Path $DestinationDirectory "setup.exe")

        # Set share
        Get-SmbShare | Where-Object Name -eq $DirectoryName | Remove-SmbShare -Force
        New-SmbShare -Name $DirectoryName -Path $DestinationDirectory -ReadAccess "Everyone" -LeasingMode Full -Description $NewMsOfficeDeploymentToolShareMessageTable.SharedFolderDescription | Out-Null

        # Register task
        If (-not $NoRegisterTask){
            $Actions = (New-ScheduledTaskAction -Execute "%WinDir%\system32\WindowsPowerShell\v1.0\powershell.exe" -Argument ("-ExecutionPolicy ByPass -Command ""Import-Module ManagementMsOfficeDeploymentToolShare; Invoke-MsOfficeDeploymentToolAndCompress -WorkingDirectory " + "'" + "$DestinationDirectory" + "'" + " -ConfigFileName " + "'" + "$(Split-Path $ConfigPath -Leaf)" + "'" + """"))
            If ($TaskTrigger -eq $Null){
                $TaskTrigger = New-ScheduledTaskTrigger -Daily -At ([DateTime]"2:00").AddMinutes((Get-Random -Maximum (60 * 5))).ToString("H:mm") # 2:00 AM ~ 5:00 AM, Everyday
            }
            $Principal = New-ScheduledTaskPrincipal -UserId "System" -RunLevel Highest
            $Settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -WakeToRun
            $Task = New-ScheduledTask -Action $Actions -Principal $Principal -Trigger $TaskTrigger -Settings $Settings -Description $NewMsOfficeDeploymentToolShareMessageTable.TaskDescription
            $Task.Author = "ManagementMsOfficeDeploymentToolShare"
            Register-ScheduledTask $TaskName -InputObject $Task | Out-Null
            If ($StartTask){
                Start-ScheduledTask -TaskName $TaskName
            }
        }
    }
    Catch{
        Write-Error "$($_.Exception.Message)"
    }
    Try{
        Remove-Item $TemporaryDirectoryPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    Catch{}
}
Export-ModuleMember -Function New-MsOfficeDeploymentToolShare