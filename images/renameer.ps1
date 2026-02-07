$DO_RENAME = $true   # set to $true when ready

$shell  = New-Object -ComObject Shell.Application
$folder = $shell.Namespace((Get-Location).Path)

# Find the "Date taken" column index dynamically
$DATE_TAKEN_IDX = (0..400 | Where-Object { $folder.GetDetailsOf($null, $_) -eq 'Date taken' } | Select-Object -First 1)
if ($null -eq $DATE_TAKEN_IDX) { throw "'Date taken' column not found." }

function Clean-DateText([string]$s) {
    if (-not $s) { return "" }
    $s = [regex]::Replace($s, "[\p{Cf}\u00A0]", "")
    $s.Trim()
}

function Parse-DateTaken([string]$dtText) {
    $t = Clean-DateText $dtText
    if (-not $t) { return $null }

    $formats = @(
        "M/d/yyyy h:mm tt",
        "M/d/yyyy hh:mm tt",
        "MM/dd/yyyy h:mm tt",
        "MM/dd/yyyy hh:mm tt",
        "M/d/yyyy h:mm:ss tt",
        "MM/dd/yyyy h:mm:ss tt"
    )

    $cultures = @(
        [System.Globalization.CultureInfo]::CurrentCulture,
        [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
    )

    foreach ($c in $cultures) {
        foreach ($fmt in $formats) {
            $d = [datetime]::MinValue
            if ([datetime]::TryParseExact(
                $t, $fmt, $c,
                [System.Globalization.DateTimeStyles]::AllowWhiteSpaces,
                [ref]$d
            )) {
                return $d
            }
        }
    }
    return $null
}

$files = Get-ChildItem -File |
    Where-Object { $_.Extension -ne '.ps1' } |
    ForEach-Object {

        $item   = $folder.ParseName($_.Name)
        $dtText = $folder.GetDetailsOf($item, $DATE_TAKEN_IDX)
        $taken  = Parse-DateTaken $dtText

        if ($taken) {
            [pscustomobject]@{
                File = $_
                Date = $taken
                Est  = $false
            }
        } else {
            [pscustomobject]@{
                File = $_
                Date = $_.CreationTime
                Est  = $true
            }
        }
    } |
    Sort-Object Date, @{Expression={$_.File.Name}; Ascending=$true}

$i = 0
foreach ($x in $files) {
    $f = $x.File
    $suffix = if ($x.Est) { " est" } else { "" }

    $newName = ('{0:D3} - {1:yyyy-MM-dd, HHmm}{2}{3}' -f `
        $i, $x.Date, $suffix, $f.Extension)

    if ($DO_RENAME) {
        Rename-Item -LiteralPath $f.FullName -NewName $newName
    } else {
        "{0}  [{1:yyyy-MM-dd HH:mm}{2}]  ->  {3}" -f `
            $f.Name, $x.Date, $suffix, $newName
    }

    $i++
}
