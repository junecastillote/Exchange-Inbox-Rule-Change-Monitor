Function Write-InboxRuleChangeHTMLReport {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [psobject]$ReportObject,

        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    # Validate filename format
    $Filename = (Split-Path $Path -Leaf)
    $regex = "([a-zA-Z0-9\s_\\.\-\(\):])+(.html|.html)$"
    if ($Filename -notmatch $regex) {
        Write-Output "The specified filename - $Filename - is not valid. Only valid characters are accepted and the extension must be .HTML or .HTM"
        Break
    }

    # Create directory if it does not exist
    $Directory = (Split-Path $Path -Parent)
    if (!(Test-Path $Directory)) {
        $null = New-Item -ItemType Directory -Path $Directory -Force
    }

    # Group the report object using the RuleName
    $ReportObject = ($ReportObject | Group-Object RuleName)
    $ModuleInfo = Get-Module ExchangeInboxRuleChangeMonitor

    # Build the HTML report
    $css = Get-Content (($ModuleInfo.ModuleBase.ToString()) + '\source\resource\style.css') -Raw
    $title = "Inbox Rule Change Detected on Mailbox - $($Mailbox)"

    $html += '<html><head><title>' + $title + '</title>'
    $html += '<style type="text/css">'
    $html += $css
    $html += '</style></head>'
    $html += '<body>'
    $html += '<table id="rules">'
    foreach ($Report in $ReportObject) {
        [int]$x = ($Report.Count + 1)
        $html += '<tr>'
        $html += '<th rowspan="'+$x+'">' + ($Report.Name) +'</th>'
        $html += '<th>Property Name</th>'
        $html += '<th>Old Value</th>'
        $html += '<th>New Value</th>'

        foreach ($item in ($Report.Group)) {
            $html += '<tr>'
            $html += '<td>' + ($item.PropertyName) + '</td>'
            $html += '<td>' + ($item.OldValue) + '</td>'
            $html += '<td>' + ($item.NewValue) + '</td>'
            $html += '</tr>'
        }
        $html += '</tr>'
    }
    $html += '</table>'
    $html += '<p style="font-family:Tahoma;"><br><br><br><br>'
    $html += '<a href="'+ $ModuleInfo.ProjectUri +'">'+($ModuleInfo.Name)+'</a>'
    $html += '</p>'
    $html += '</body></html>'
    $html | Out-File $Path -Encoding utf8 -Force
}