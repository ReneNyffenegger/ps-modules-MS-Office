@{
   RootModule        = 'MS-Office.psm1'
   ModuleVersion     = '0.3'
   RequiredModules   = @(
      'COM'
   )
   FunctionsToExport = @(
      'get-msOfficeVersion',
      'enable-msOfficeDeveloperTab',
      'get-msOfficeComObject',
      'get-msOfficeInstallationRoot'
   )
   AliasesToExport   = @(
   )
}
