@{
   RootModule        = 'MS-Office.psm1'
   ModuleVersion     = '0.2'
   RequiredModules   = @(
      'COM'
   )
   FunctionsToExport = @(
      'get-msOfficeVersion',
      'enable-msOfficeDeveloperTab',
      'get-msOfficeComObject'
   )
   AliasesToExport   = @(
   )
}
