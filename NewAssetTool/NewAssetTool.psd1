@{
    RootModule        = 'NewAssetTool.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '00000000-0000-0000-0000-000000000001'
    Author            = 'NewAssetTool Team'
    CompanyName       = 'Contoso'
    CompatiblePSEditions = @('Desktop','Core')
    PowerShellVersion = '5.1'
    Description       = 'Core UI and helpers for the New Asset Tool.'
    FileList          = @('NewAssetTool.psm1','NewAssetTool.NativeMethods.dll')
    FunctionsToExport = @('Start-NewAssetTool')
    AliasesToExport   = @()
    CmdletsToExport   = @()
    VariablesToExport = '*'
}
