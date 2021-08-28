#
#  V0.3
#
set-strictMode -version latest

function get-msOfficeVersion {
   return ( (get-item hklm:\Software\Classes\excel.application\curVer).getValue('')  -replace '.*\.(\d+)', '$1' )
}


function get-msOfficeProducts {

   return (new-object psObject -property @{ name = 'Excel'     ; exe = 'excel.exe'    }),
          (new-object psObject -property @{ name = 'Word'      ; exe = 'winWord.exe'  }),
          (new-object psObject -property @{ name = 'PowerPoint'; exe = 'powerpnt.exe' }),
          (new-object psObject -property @{ name = 'Outlook'   ; exe = 'outlook.exe'  })

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

function get-msOfficeInstallationRoot {

    foreach ($prod in (get-msOfficeProducts)) {
       $ret = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$($prod.exe)" path -errorAction ignore
       if ($ret -ne $null) {
           return $ret
       }
    }

    write-textInConsoleWarningColor "No installation root found for MS Office"
    return $null
}
