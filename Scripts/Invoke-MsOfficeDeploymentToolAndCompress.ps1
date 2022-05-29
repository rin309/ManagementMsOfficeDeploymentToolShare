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

 .Example
   # 
   Invoke-MsOfficeDeploymentToolAndCompress -WorkingDirectory "C:\Office" -ConfigFileName "Configuration.xml"
 
#>
Function Invoke-MsOfficeDeploymentToolAndCompress{
    Param(
        [Parameter(Mandatory)][String]$WorkingDirectory,
        [Parameter(Mandatory)][String]$ConfigFileName,
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
                    Write-EventLog -LogName Application -Source $EventLogSourceName -EventID $Process.ExitCode -Message $Message -Category 2 -EntryType Error
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
    $EventLogSourceName = "ManagementMsOfficeDeploymentToolShare"
    Try{
    
        Set-Location -Path $WorkingDirectory
        Invoke-Application -Path "setup.exe" -Argument "/Download $ConfigFileName" -WorkingDirectory $WorkingDirectory -NoEventLogging $NoEventLogging -EventLogSourceName $EventLogSourceName
        $Message = Remove-ClickToRunOldVersionsFromDirectory
        If (-not $NoEventLogging){
            Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 0 -Message $Message -Category 1
        }
    }
    Catch{
        If (-not $NoEventLogging){
            Write-EventLog -LogName Application -Source $EventLogSourceName -EventID 1 -Message "$($_.Exception.Message)" -Category 1 -EntryType Error
        }
    }
    
}
Export-ModuleMember -Function Invoke-MsOfficeDeploymentToolAndCompress
