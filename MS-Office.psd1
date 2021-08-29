@{
   RootModule        = 'MS-Office.psm1'
   ModuleVersion     = '0.4'
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
