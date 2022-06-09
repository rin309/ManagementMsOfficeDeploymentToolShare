#Requires -Version 5.0

<#
 .Synopsis
   Assists in ODT placement on the server

 .Description
   The task schedule automatically deletes the old version.

 .Parameter WorkingDirectory
   Root folder to place shared folders

 .Parameter ConfigFileName
   Set ODT config file name
   Create from 'Microsoft 365 Apps admin center'

   https://config.office.com/

 .Parameter UsingOfficeReleases
   Use Office release information for cleanup
   
   https://clients.config.office.net/releases/v1.0/OfficeReleases

 .Parameter NoEventLogging
   No event logging

 .Example
   # Simply command sample
   Invoke-MsOfficeDeploymentToolAndCompress -WorkingDirectory "C:\Office" -ConfigFileName "Configuration.xml"

 .Example
   # Multiple configuration command sample
   Invoke-MsOfficeDeploymentToolAndCompress -WorkingDirectory "C:\Office" -ConfigFileName @("configuration-Office365-x64.xml","configuration-Office2021Enterprise.xml")

#>
Function Invoke-MsOfficeDeploymentToolAndCompress{
    Param(
        [Parameter(Mandatory)][String]$WorkingDirectory,
        [Parameter(Mandatory)][String[]]$ConfigFileName=@(),
        [Switch]$UsingOfficeReleases = $False,
        [Switch]$NoEventLogging = $False
    )

    Function Invoke-Application($Path, $Argument, $WorkingDirectory, $NoEventLogging, $EventLogSourceName){
        $VersionInfo = (Get-ItemProperty $Path).VersionInfo
        $Message = "$Path`n$($VersionInfo.ProductVersion)"

        If (-not ($EventLogSourceName -in ((Get-WmiObject -Class Win32_NTEventlogFile -Filter "FileName='Application'").Sources))){
            New-EventLog -LogName Application -Source $EventLogSourceName
            Limit-EventLog -LogName $EventLogSourceName -RetentionDays 180
        }

        Try{
            $Process = Start-Process -FilePath $Path -ArgumentList $Argument -WorkingDirectory $WorkingDirectory -WindowStyle Hidden -PassThru -Wait
            If ($Process.ExitCode -ne 0){
                Write-Warning "$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.WarnExitCodeByInvokeApplication) [$Path $Argument] : $($Process.ExitCode)" -Verbose
                If (-not $NoEventLogging){
                    Write-EventLog -LogName Application -Source $EventLogSourceName -EventID ([Math]::Abs($Process.ExitCode)) -Message $Message -Category 2 -EntryType Error
                }
                Return
            }
            If (-not $NoEventLogging){
                Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 0 -Message $Message -Category 2
            }
        }
        Catch{
            Write-Warning "$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.WarnByInvokeApplication) [$Path $Argument] : $($_.Exception.Message)"
            If (-not $NoEventLogging){
                Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message "$($_.Exception.Message)`n`n$Message" -Category 2 -EntryType Error
            }
        }
    }
    Function Remove-ClickToRunOldVersionsFromDirectory{
        $FirstItem = $True
        $OfficeDataDirectoryPath = ".\Office\Data"
        $Log = "$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.LatestLabel):"
        Get-ChildItem -Path $OfficeDataDirectoryPath -Directory | Where-Object {($_.Name -as [System.Version]) -ne $Null} | Sort-Object {[System.Version]$_.Name} -Descending | ForEach-Object {
            If (-not $FirstItem){
                Remove-Item -Path $_.FullName -Force -Recurse
                Get-ChildItem -Path $OfficeDataDirectoryPath -File -Filter "v64_$($_.Name).cab" | Remove-Item -Force
                Get-ChildItem -Path $OfficeDataDirectoryPath -File -Filter "v86_$($_.Name).cab" | Remove-Item -Force
                $Log = "$Log`n$($_.Name)"
            }
            Else{
                $Log = "$Log`n$($_.Name)`n`n$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.RemovedLabel):"
            }
            $FirstItem = $False
        }
        Return $Log
    }
    Function Get-OfficeChannelDisplayName($Channel){
        Switch($Channel){
            "Beta" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.BetaChannelName}
            "InsiderFast" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.InsiderFastChannelName}
            "CurrentPreview" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.CurrentPreviewChannelName}
            "Insiders" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.InsidersChannelName}
            "Current" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.CurrentChannelName}
            "Monthly" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.MonthlyChannelName}
            "MonthlyEnterprise" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.MonthlyEnterpriseChannelName}
            "SemiAnnualEnterprisePreview" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.SemiAnnualEnterprisePreviewChannelName}
            "Targeted" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.TargetedChannelName}
            "SemiAnnual" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.SemiAnnualChannelName}
            "Broad" {$InvokeMsOfficeDeploymentToolAndCompressMessageTable.BroadChannelName}
            default {Return "$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.UnknownChannelName) [$Channel]"}
        }
    }
    Function Remove-ClickToRunOldVersionsFromDirectoryUsingOfficeReleases{
        $OfficeDataDirectoryPath = ".\Office\Data"

        Try{
            # support Windows Server 2016
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

            $OfficeReleases = Invoke-WebRequest "https://clients.config.office.net/releases/v1.0/OfficeReleases"
            If ($OfficeReleases.StatusCode -eq "200"){
                $IgnoreVersions = @()
                @($ConfigFileName) | ForEach-Object {
                    $Xml = @(([xml](Get-Content $_ -Encoding UTF8)).Configuration.Add)[0]
                    $Log += "$(Get-OfficeChannelDisplayName $Xml.Channel) $($InvokeMsOfficeDeploymentToolAndCompressMessageTable.LatestLabel): "
                    If ($Xml.Version -eq $Null){
                        $LatestVersion = (($OfficeReleases.Content | ConvertFrom-Json).SyncRoot | Where-Object channelId -like ($Xml.Channel)).latestVersion
                        $Log += "$LatestVersion`n"
                        $IgnoreVersions += $LatestVersion
                    }
                    Else{
                        $Log += "$($Xml.Version)`n"
                        $IgnoreVersions += $Xml.Version
                    }
                }
            }
            Else{
                Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message $InvokeMsOfficeDeploymentToolAndCompressMessageTable.NotFoundInvokeDownloadOfficeReleases -Category 3 -EntryType Warning
                Return
            }
        }
        Catch{
            Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message "$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.NotFoundInvokeDownloadOfficeReleases): $($_.Exception.Message)" -Category 3 -EntryType Error
            Return
        }

        If ($IgnoreVersions -eq $Null){
            Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message $InvokeMsOfficeDeploymentToolAndCompressMessageTable.NotFoundLatestVersionFromOfficeReleases -Category 3 -EntryType Warning
            Return
        }
        Else{
            $Log += "`n`n$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.RemovedLabel):`n"
            Get-ChildItem -Path $OfficeDataDirectoryPath -Directory | Where-Object {($_.Name -as [System.Version]) -ne $Null} | Where-Object Name -notin $IgnoreVersions | ForEach-Object {
                Remove-Item -Path $_.FullName -Force -Recurse
                Get-ChildItem -Path $OfficeDataDirectoryPath -File -Filter "v64_$($_.Name).cab" | Remove-Item -Force
                Get-ChildItem -Path $OfficeDataDirectoryPath -File -Filter "v86_$($_.Name).cab" | Remove-Item -Force

                $Log += "`n$($_.Name)"
            }
        }
        Return $Log
    }
    $EventLogSourceName = "ManagementMsOfficeDeploymentToolShare"
    Try{
        If (-not (Test-Path $WorkingDirectory)){
            If (-not $NoEventLogging){
                Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message "$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.NotFoundWorkingDirectory): [$WorkingDirectory]" -Category 1 -EntryType Error
            }
        }
        If (-not (Test-Path (Join-Path $WorkingDirectory "setup.exe"))){
            If (-not $NoEventLogging){
                Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message "$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.NotFoundSetupExe): [$WorkingDirectory]" -Category 1 -EntryType Error
            }
        }
        @($ConfigFileName) | ForEach-Object {
            If (-not (Test-Path (Join-Path $WorkingDirectory $_))){
                If (-not $NoEventLogging){
                    Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message "$($InvokeMsOfficeDeploymentToolAndCompressMessageTable.NotFoundConfigFile): [$_]" -Category 1 -EntryType Warning
                }
            }
        }
        If ($ConfigPath.Count -gt 1){
            $UsingOfficeReleases = $True
            If (-not $NoEventLogging){
                Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message $InvokeMsOfficeDeploymentToolAndCompressMessageTable.UsingOfficeReleasesIsEnabled -Category 1 -EntryType Information
            }
        }

        Set-Location -Path $WorkingDirectory
        @($ConfigFileName) | ForEach-Object {
            Invoke-Application -Path "setup.exe" -Argument "/Download $_" -WorkingDirectory $WorkingDirectory -NoEventLogging $NoEventLogging -EventLogSourceName $EventLogSourceName
        }

        $Message = ""
        If ($UsingOfficeReleases){
            $Message += Remove-ClickToRunOldVersionsFromDirectoryUsingOfficeReleases
        }
        If (-not $UsingOfficeReleases){
            $Message += Remove-ClickToRunOldVersionsFromDirectory
        }

        If (-not $NoEventLogging){
            Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 0 -Message $Message -Category 1
        }
    }
    Catch{
        If (-not $NoEventLogging){
            Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message $_.Exception.Message -Category 1 -EntryType Error
        }
    }
    
}
Export-ModuleMember -Function Invoke-MsOfficeDeploymentToolAndCompress
