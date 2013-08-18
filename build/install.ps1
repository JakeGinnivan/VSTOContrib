param($installPath, $toolsPath, $package, $project)

$frameworkVersionString = $project.Properties.Item('TargetFrameworkMoniker').Value
$frameworkVersionString -match 'Version=(v\d\.\d)' > null
$frameworkVersion = $matches[1]

if ($frameworkVersion -ge 'v4.0')
{
    $project.Object.References | Where-Object { $_.EmbedInteropTypes -eq $true } | ForEach-Object { $_.EmbedInteropTypes = $false }
}

Write-Host "Adding VSTO Contrib References"

$interopAssembly = $project.Object.References | Where-Object { $_.Name.StartsWith('Microsoft.Office.Interop.{{Application}}') }
$frameworkVersionString = $project.Properties.Item('TargetFrameworkMoniker').Value
$frameworkVersionString -match 'Version=(v\d\.\d)' > null
$frameworkVersion = $matches[1]
if ($interopAssembly.MajorVersion -eq '14') { $officeVersion = '2010' }
if ($interopAssembly.MajorVersion -eq '15') { $officeVersion = '2013' }
else { $officeVersion = '2007' }

$netVersionFolder = "bin-" + $netVersionFolder

$vstoContribCoreDll = Join-Path (Join-Path (Join-Path $toolsPath "bin") $officeVersion) "VSTOContrib.Core.dll"
$vstoContribApplicationDll = Join-Path (Join-Path (Join-Path $toolsPath "bin") $officeVersion) "VSTOContrib.{{Application}}.dll"

Write-Host "Adding references to $vstoContribCoreDll and $vstoContribApplicationDll"
$project.Object.References.Add($vstoContribCoreDll)
$project.Object.References.Add($vstoContribApplicationDll)

$project.Save()