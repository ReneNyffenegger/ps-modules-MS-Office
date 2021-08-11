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

function get-msOfficeComObject {

   param (
      [string] $app
   )

   $progId = "$app.application"

   $officeObj = get-activeObject $progId
   if ($officeObj -eq $null) {
      write-debug "no obj found for $progId"
      $officeObj = new-object -com $progId
      if ($app -eq 'outlook') {
         #
         # Outlook does not have a .visible property on the application object!
         #
         # 6 = olFolderInbox
         #
         $officeObj.GetNamespace('MAPI').GetDefaultFolder(6).display()
      }
      else {
         $officeObj.visible = $true
      }
   }
   return $officeObj
}
