set-strictMode -version latest

function get-msOfficeVersion {
   return ( (get-item hklm:\Software\Classes\excel.application\curVer).getValue('')  -replace '.*\.(\d+)', '$1' )
}

function enable-msOfficeDeveloperTab {

   $regKeyOfficeRootV = "hkcu:\Software\Microsoft\Office\$(get-msOfficeVersion).0"

   foreach ($app in ( (get-childItem $regKeyOfficeRootV).psChildName )) { # | select-object psChildName  ) ) {
      if ('Excel', 'Outlook', 'PowerPoint', 'Word' -contains $app) {
         set-itemProperty "$regKeyOfficeRootV\$app\Options" -name DeveloperTools -type dWord -value 1
      }
   }
}
