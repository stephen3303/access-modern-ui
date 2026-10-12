param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
if(!$OutputDirectory){$OutputDirectory=Join-Path $env:TEMP ('access-modern-ui-'+[guid]::NewGuid())}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$database=Join-Path $OutputDirectory 'kit-test.accdb'
if(Test-Path $database){throw 'Use an empty output directory; test database already exists.'}
$access=New-Object -ComObject Access.Application
try {
 $access.Visible=$false
 $access.NewCurrentDatabase($database)
 $modules=@(Get-ChildItem (Join-Path $repo 'kit/modules') -Filter *.bas)+@(Get-ChildItem (Join-Path $repo 'examples/palette') -Filter *.bas)+@(Get-Item (Join-Path $repo 'examples/modules/modErrors.bas'))+@(Get-ChildItem $PSScriptRoot -Filter *.bas)
 foreach($module in $modules){
  $temporary=Join-Path $OutputDirectory $module.Name
  $text=[IO.File]::ReadAllText($module.FullName) -replace "`r?`n","`r`n"
  [IO.File]::WriteAllText($temporary,$text,[Text.Encoding]::ASCII)
  $access.LoadFromText(5,$module.BaseName,$temporary)
 }
 $access.DoCmd.RunCommand(126)
 $formDir=Join-Path $OutputDirectory 'kit/forms'
 New-Item -ItemType Directory -Force -Path $formDir | Out-Null
 Copy-Item (Join-Path $repo 'kit/forms/subGrid.vba') $formDir
 Copy-Item (Join-Path $PSScriptRoot 'testHost.vba') $formDir
 $arg=[string]$OutputDirectory
 $result=$access.Run('TestKit',[ref]$arg)
 $result|Set-Content (Join-Path $OutputDirectory 'result.txt')
 Write-Output $result
}catch{
 $_.Exception.Message|Set-Content (Join-Path $OutputDirectory 'result.txt')
 throw
}finally{
 try{$access.CloseCurrentDatabase()}catch{}
 $access.Quit(2)
 [Runtime.InteropServices.Marshal]::ReleaseComObject($access)|Out-Null
}
