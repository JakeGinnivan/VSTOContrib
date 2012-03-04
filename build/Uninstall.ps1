param($installPath, $toolsPath, $package, $project)

$project.Object.References | Where-Object -FilterScript { $_.Name.StartsWith('VSTOContrib') } | ForEach-Object { $_.Remove() }
$project.Save()