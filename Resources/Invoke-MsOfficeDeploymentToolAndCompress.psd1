$InvokeMsOfficeDeploymentToolAndCompressMessageTable = Data {
    ConvertFrom-StringData @'
    NotFoundWorkingDirectory = Could not find WorkingDirectory
    NotFoundSetupExe = Could not find Setup.exe. Place the Office Deployment Tool in WorkingDirectory.
    NotFoundConfigFile = Could not find ConfigFileName

    WarnExitCodeByInvokeApplication = The return value of the executed application was non-zero. Execution may have failed.
    WarnByInvokeApplication = Failed to run the application.
    NotFoundInvokeDownloadOfficeReleases = Could not get Office release information
    NotFoundLatestVersionFromOfficeReleases = Could not get the latest version from Office release information
    UsingOfficeReleasesIsEnabled = UsingOfficeReleases was enabled because multiple files were specified

    LatestLabel = Latest
    RemovedLabel = Removed

    BetaChannelName = Beta Channel
    InsiderFastChannelName = Beta Channel (Insider Fast channel)
    CurrentPreviewChannelName = Current Channel [Preview]
    InsidersChannelName = Current Channel [Preview] (Insiders channel)
    CurrentChannelName = Current channel
    MonthlyChannelName = Current channel (Monthly Channel)
    MonthlyEnterpriseChannelName = Monthly Enterprise channel
    SemiAnnualEnterprisePreviewChannelName = Semi-Annual Enterprise Channel [Preview]
    TargetedChannelName = Semi-Annual Enterprise Channel [Preview] (Target Channel)
    SemiAnnualChannelName = Semi-Annual Enterprise Channel
    BroadChannelName = Semi-Annual Enterprise Channel (Broad channel)
    PerpetualVL2019 = Microsoft Offie 2019 [Volume License]
    PerpetualVL2021 = Microsoft Offie 2021 [Volume License]

    UnknownChannelName = Unknown Channel

    NotFoundTargetVersionInChannel = No valid {Version} was found for {ChannelName}

'@
}
