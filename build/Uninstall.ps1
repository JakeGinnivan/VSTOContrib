param($installPath, $toolsPath, $package, $project)

$project.Object.References | Where-Object { $_.Name.StartsWith('VSTOContrib') } | { $_.Remove() }
