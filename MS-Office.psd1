@{
   RootModule        = 'MS-Office.psm1'
   ModuleVersion     = '0.5'
   RequiredModules   = @(
      'COM'
   )
   FunctionsToExport = @(
      'get-msOfficeVersion',
      'enable-msOfficeDeveloperTab',
      'grant-msOfficeVBAaccess',
      'get-msOfficeComObject',
      'get-msOfficeInstallationRoot'
   )
   AliasesToExport   = @(
   )
}
