param([string]$OutputDirectory, [switch]$UseGdiMeasurement)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
if(!$OutputDirectory){$OutputDirectory=Join-Path $env:TEMP ('access-modern-ui-'+[guid]::NewGuid())}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$database=Join-Path $OutputDirectory 'kit-test.accdb'
if(Test-Path $database){throw 'Use an empty output directory; test database already exists.'}
. (Join-Path $repo 'tools/VbaLineNumbers.ps1')   # test the numbered build copy, as the apps run it
. (Join-Path $repo 'tools/AccessWatch.ps1')      # a dialog fails the run instead of stalling it
$access=New-Object -ComObject Access.Application
$access.Visible=$false
$watch=Start-AccessWatch $access -TimeoutSec 600
$dialogs=@()
try {
 $access.NewCurrentDatabase($database)
 $modules=@(Get-ChildItem (Join-Path $repo 'kit/modules') -Filter *.bas)+@(Get-ChildItem (Join-Path $repo 'examples/palette') -Filter *.bas)+@(Get-Item (Join-Path $repo 'examples/modules/modErrors.bas'))+@(Get-ChildItem $PSScriptRoot -Filter *.bas)
 if($UseGdiMeasurement){$modules += Get-Item (Join-Path $repo 'kit/adapters/modTextMeasure.bas')}
 foreach($module in $modules){
  $temporary=Join-Path $OutputDirectory $module.Name
  $text=[IO.File]::ReadAllText($module.FullName) -replace "`r?`n","`r`n"
  if($UseGdiMeasurement -and $module.BaseName -eq 'modTestHooks'){
   $text=[regex]::Replace($text,'(?s)Public Function MeasureText\(.*?End Function\r\n','')
  }
  $text=Add-VbaLineNumbers $text
  [IO.File]::WriteAllText($temporary,$text,[Text.Encoding]::ASCII)
  $access.LoadFromText(5,$module.BaseName,$temporary)
 }
 $access.DoCmd.RunCommand(126)
 if($UseGdiMeasurement){
  $wide='WWW';$narrow='iii';$fontSize=[long]9;$weight=[long]400
  $wideWidth=$access.Run('MeasureText',[ref]$wide,[ref]$fontSize,[ref]$weight)
  $narrowWidth=$access.Run('MeasureText',[ref]$narrow,[ref]$fontSize,[ref]$weight)
  if($wideWidth -le $narrowWidth){throw 'GDI measurement did not distinguish proportional character widths'}
 }
 $formDir=Join-Path $OutputDirectory 'kit/forms'
 New-Item -ItemType Directory -Force -Path $formDir | Out-Null
 Copy-Item (Join-Path $repo 'kit/forms/subGrid.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgMessage.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgToast.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgBusy.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgInput.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgMenu.vba') $formDir
Copy-Item (Join-Path $repo 'kit/forms/dlgPalette.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgHelp.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgPopover.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgPersonnelSetup.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgSignIn.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgPersonnel.vba') $formDir
 Copy-Item (Join-Path $repo 'kit/forms/dlgPerson.vba') $formDir
 Copy-Item (Join-Path $PSScriptRoot 'testHost.vba') $formDir
 Copy-Item (Join-Path $PSScriptRoot 'testRefuse.vba') $formDir
 Copy-Item (Join-Path $PSScriptRoot 'testConflict.vba') $formDir
 Copy-Item (Join-Path $PSScriptRoot 'testStamp.vba') $formDir
 Copy-Item (Join-Path $PSScriptRoot 'testKeys.vba') $formDir
Copy-Item (Join-Path $PSScriptRoot 'testShell.vba') $formDir
Copy-Item (Join-Path $PSScriptRoot 'testLeave.vba') $formDir
Copy-Item (Join-Path $PSScriptRoot 'testShellBare.vba') $formDir
 Add-VbaLineNumbersToFolder $formDir
 $arg=[string]$OutputDirectory
 $result=$access.Run('TestKit',[ref]$arg)
 $result|Set-Content (Join-Path $OutputDirectory 'result.txt')
 Write-Output $result
}catch{
 $_.Exception.Message|Set-Content (Join-Path $OutputDirectory 'result.txt')
 throw
}finally{
 if(-not $watch.Killed){
  try{$access.CloseCurrentDatabase()}catch{}
  try{$access.Quit(2)}catch{}
 }
 try{[Runtime.InteropServices.Marshal]::ReleaseComObject($access)|Out-Null}catch{}
 $dialogs=@(Stop-AccessWatch $watch)
 if($dialogs.Count){('Access showed a dialog: '+($dialogs -join ' / '))|Add-Content (Join-Path $OutputDirectory 'result.txt')}
}
# Any dialog fails the run, even when every check passed: unattended, it would have hung.
if($dialogs.Count){throw ('Access showed a dialog: '+($dialogs -join ' / '))}
