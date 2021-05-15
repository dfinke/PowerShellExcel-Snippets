param(
    [Switch]$ToClipboard
)

$snippets = Get-Content .\snippets\powershellExcel.code-snippets | ConvertFrom-Json

$names = $snippets[0].psobject.properties.name

$result = foreach ($name in $names) {
    [PSCustomObject][Ordered]@{
        Trigger     = $snippets.$name.prefix
        Description = $snippets.$name.description
    }
}
$md = $result | ConvertTo-MarkdownTable 

if ($ToClipboard) {
    $md | clip
    return 
}

$md

