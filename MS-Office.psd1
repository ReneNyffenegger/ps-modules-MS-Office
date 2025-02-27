@{
   RootModule        = 'MS-Office.psm1'
   ModuleVersion     = '0.8'
   RequiredModules   = @(
      'COM'
   )
   FunctionsToExport = @(
      'get-msOfficeVersion',
      'enable-msOfficeDeveloperTab',
      'grant-msOfficeVBAaccess',
      'get-msOfficeComObject',
      'get-msOfficeRegRoot',
      'get-msOfficeInstallationRoot'
   )
   AliasesToExport   = @(
   )
}
