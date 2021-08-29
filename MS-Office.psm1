#
#  V0.4
#
set-strictMode -version latest

function get-msOfficeVersion {
   return ( (get-item hklm:\Software\Classes\excel.application\curVer).getValue('')  -replace '.*\.(\d+)', '$1' )
}


function get-msOfficeProducts {

   return (new-object psObject -property ( [ordered] @{ name = 'Excel'     ; exe = 'excel.exe'   ; devTools = $true  })),
          (new-object psObject -property ( [ordered] @{ name = 'Access'    ; exe = 'msAccess.exe'; devTools = $false })),
          (new-object psObject -property ( [ordered] @{ name = 'Word'      ; exe = 'winWord.exe' ; devTools = $true  })),
          (new-object psObject -property ( [ordered] @{ name = 'Visio'     ; exe = 'visio.exe'   ; devTools = $false })),
          (new-object psObject -property ( [ordered] @{ name = 'Outlook'   ; exe = 'outlook.exe' ; devTools = $true  })),
          (new-object psObject -property ( [ordered] @{ name = 'PowerPoint'; exe = 'powerpnt.exe'; devTools = $true  }))
}

function enable-msOfficeDeveloperTab {

   param (
      [switch] $off
   )

   $regKeyOfficeRootV = "hkcu:\Software\Microsoft\Office\$(get-msOfficeVersion).0"

   foreach ($prod in get-msOfficeProducts | where-object devTools) {

      $regKeyOfficeApp ="$regKeyOfficeRootV/$($prod.name)"
      if (test-path $regKeyOfficeApp) {
         $value = 1
         if ($off) { $value = 0 }
         set-itemProperty "$regKeyOfficeApp\Options" -name DeveloperTools -type dWord -value $value
      }
      else {
          write-host "Not found for $regKeyOfficeApp"
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
