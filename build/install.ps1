param($installPath, $toolsPath, $package, $project)

$frameworkVersionString = $project.Properties.Item('TargetFrameworkMoniker').Value
$frameworkVersionString -match 'Version=(v\d\.\d)' > null
$frameworkVersion = $matches[1]

if ($frameworkVersion -ge 'v4.0')
{
    $project.Object.References | Where-Object { $_.EmbedInteropTypes -eq $true } | ForEach-Object { $_.EmbedInteropTypes = $false }
}

$project.Save()