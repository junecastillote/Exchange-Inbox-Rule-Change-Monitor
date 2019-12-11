$Path = [System.IO.Path]::Combine($PSScriptRoot, 'FUNCTIONS')
Get-Childitem $Path -Filter *.ps1 -Recurse | Foreach-Object {
    . $_.Fullname
}

[array]$a = (Import-Clixml .\a.xml)
[array]$b = (Import-Clixml .\b.xml)
$finalResult = @()
foreach ($i in $b) {
    $ref = ($a | Where-Object { $_.Identity -eq ($i.Identity) })
    if ($ref) {$finalResult += (Compare-InboxRules $ref $i)}
}
$finalResult