param($installPath, $toolsPath, $package, $project)

$interopAssembly = $project.Object.References | Where-Object { $_.Name.StartsWith('Microsoft.Office.Interop.{{Application}}') }
$frameworkVersionString = $project.Properties.Item('TargetFrameworkMoniker').Value
$frameworkVersionString -match 'Version=(v\d\.\d)' > null
$frameworkVersion = $matches[1]
if ($interopAssembly.MajorVersion -eq '14') { $officeVersion = '2010' }
else { $officeVersion = '2007' }

if ($frameworkVersion -eq 'v4.0')
{
	write-host "Installing for .net 4.0"
    $toolsAssembly = $project.Object.References | Where-Object { $_.Name.StartsWith('Microsoft.Office.Tools.{{Application}}') }
    if ($toolsAssembly.MajorVersion -eq 9)
    {
        throw '.NET 4.0 projects should not be referencing v9 of the VSTO tools. Upgrade references to correct versions'
    }

    #Disable embedded interop types
    $project.Object.References | Where-Object { $_.EmbedInteropTypes -eq $true } | ForEach-Object { $_.EmbedInteropTypes = $false }
    $netVersionFolder = 'net40'
} elseif ($frameworkVersion -eq 'v3.5')
{
	write-host "Installing for .net 3.5"
    $netVersionFolder = 'net35'
} else 
{
	write-host "Unknown .net framework version $frameworkVersion"
}

$vstoContribCoreDll = Join-Path (Join-Path (Join-Path $toolsPath $netVersionFolder) $officeVersion) "VSTOContrib.Core.dll"
$vstoContribApplicationDll = Join-Path (Join-Path (Join-Path $toolsPath $netVersionFolder) $officeVersion) "VSTOContrib.{{Application}}.dll"

write-host "Adding references to $vstoContribCoreDll and $vstoContribApplicationDll"
$project.Object.References.Add($vstoContribCoreDll)
$project.Object.References.Add($vstoContribApplicationDll)

$project.Save()