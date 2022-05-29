Get-ChildItem -Path (Join-Path $PSScriptRoot "Scripts\*.ps1") | ForEach-Object { . $_}
Import-LocalizedData -BindingVariable NewMsOfficeDeploymentToolShareMessageTable -BaseDirectory (Join-Path $PSScriptRoot "Resources") -FileName "New-MsOfficeDeploymentToolShare.psd1"
Import-LocalizedData -BindingVariable InvokeMsOfficeDeploymentToolAndCompressMessageTable -BaseDirectory (Join-Path $PSScriptRoot "Resources") -FileName "Invoke-MsOfficeDeploymentToolAndCompress.psd1"
