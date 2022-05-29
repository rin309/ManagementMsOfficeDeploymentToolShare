$NewMsOfficeDeploymentToolShareMessageTable = Data {
    ConvertFrom-StringData @'
    WarnExitCodeByInvokeApplication = The return value of the executed application was non-zero. Execution may have failed.
    WarnByInvokeApplication = Failed to run the application.
    WarnByInvokeDownloadToTemporaryDirectory = Failed to download the file.

    NotFoundConfigPath = Could not find ConfigPath
    NotFoundDestinationDirectory = Could not find DestinationDirectory
    NotFoundLocalOfficeDeploymentToolPath = Could not find LocalOfficeDeploymentToolPath
    ExistsFromShares = File shares already exists
    ExistsFromTaskSchedule = Task schedule already exists
    InvalidDigitalSignature = Execution was canceled because the digital signature is invalid. Please download again or make sure your PC's digital signature is updated.
    NotFoundTemporaryMsOdtPath = Could not find TemporaryMsOdtPath

    AgreeLable = &Agree
    ExitLable = E&xit
    MsOdtEulaLabel = MICROSOFT SOFTWARE LICENSE TERMS, MICROSOFT OFFICE DEPLOYMENT TOOL
    MsOdtEulaMessage = Check the contents by executing officedeploymenttool_*.exe

    SharedFolderDescription = Using upates for Office application
    TaskDescription = Microsoft Office application updates and removal of older versions of the cache
'@
}