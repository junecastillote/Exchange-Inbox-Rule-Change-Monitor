Function Write-InboxRuleListHTMLReport {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [psobject]$ReportObject,

        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $ReportObject = ($ReportObject | Sort-Object RuleState)

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

    $ModuleInfo = Get-Module ExchangeInboxRuleChangeMonitor

    # Build the HTML report
    $css = Get-Content (($ModuleInfo.ModuleBase.ToString()) + '\source\resource\style.css') -Raw
    $title = "Inbox Rules List - $($Mailbox)"

    $html += '<html><head><title>' + $title + '</title>'
    $html += '<style type="text/css">'
    $html += $css
    $html += '</style></head>'
    $html += '<body>'
    $html += '<table id="rules">'
    $html += '<tr>'
    $html += '<th>Rule Name</th>'
    $html += '<th>Rule ID</th>'
    $html += '<th>Status</th>'
    foreach ($Report in $ReportObject) {
        $html += '<tr>'
        $html += '<td>' + ($Report.Name) + '</td>'
        $html += '<td>' + ($Report.RuleIdentity) + '</td>'
        $html += '<td>' + ($Report.RuleState) + '</td>'
        $html += '</tr>'
        $html += '</tr>'
    }
    $html += '</table>'
    $html += '<p style="font-family:Tahoma;"><br><br><br><br>'
    $html += '<a href="' + $ModuleInfo.ProjectUri + '">' + ($ModuleInfo.Name) + ' v' + ($ModuleInfo.Version.ToString()) + '</a>'
    $html += '</p>'
    $html += '</body></html>'
    $html | Out-File $Path -Encoding utf8 -Force
}