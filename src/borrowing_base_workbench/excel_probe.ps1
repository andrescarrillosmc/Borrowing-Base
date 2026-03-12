param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"

function Emit-Json($Object) {
    $Object | ConvertTo-Json -Depth 6 -Compress
}

function New-StagedWorkbookCopy {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $scratchDir = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) "scratch"
    New-Item -ItemType Directory -Force -Path $scratchDir | Out-Null
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $extension = [System.IO.Path]::GetExtension($SourcePath)
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
    $stagedPath = Join-Path $scratchDir ("excel_probe_" + $fileName + "_" + $timestamp + $extension)
    Copy-Item -LiteralPath $SourcePath -Destination $stagedPath -Force
    return $stagedPath
}

$excel = $null
$workbook = $null
$openedPath = $WorkbookPath

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    try {
        $openedPath = New-StagedWorkbookCopy -SourcePath $WorkbookPath
        $workbook = $excel.Workbooks.Open($openedPath)
    }
    catch {
        Emit-Json @{
            status = "error"
            workbook_path = $WorkbookPath
            opened_workbook_path = $openedPath
            message = $_.Exception.Message
            remediation = @(
                "Close any open handle to the workbook and try Read Current Model again.",
                "Verify the product workbook still exists at the configured path.",
                "If OneDrive is delaying the file, wait for sync completion and retry."
            )
        }
        exit 0
    }

    $excel.CalculateFullRebuild()
    $availability = $workbook.Worksheets("Availability")
    $dealTeam = $workbook.Worksheets("Deal Team Input")
    $cl = $workbook.Worksheets("Concentration Limits")

    $limitRows = @(
        @{ DtiRow = 74; ClRow = 49 },
        @{ DtiRow = 75; ClRow = 50 },
        @{ DtiRow = 76; ClRow = 51 },
        @{ DtiRow = 78; ClRow = 52 },
        @{ DtiRow = 79; ClRow = 57 },
        @{ DtiRow = 80; ClRow = 58 },
        @{ DtiRow = 82; ClRow = 59 },
        @{ DtiRow = 83; ClRow = 61 },
        @{ DtiRow = 84; ClRow = 64 },
        @{ DtiRow = 85; ClRow = 67 },
        @{ DtiRow = 87; ClRow = 68 },
        @{ DtiRow = 88; ClRow = 69 },
        @{ DtiRow = 89; ClRow = 70 }
    )

    $limits = @()
    foreach ($row in $limitRows) {
        $limits += @{
            limit_type = $dealTeam.Range("A$($row.DtiRow)").Text
            limit_percent = $dealTeam.Range("B$($row.DtiRow)").Value2
            applicable_limit = $dealTeam.Range("C$($row.DtiRow)").Value2
            prior_actual = $cl.Range("L$($row.ClRow)").Value2
            prior_excess = $cl.Range("K$($row.ClRow)").Value2
        }
    }

    Emit-Json @{
        status = "ok"
        workbook_path = $WorkbookPath
        opened_workbook_path = $openedPath
        metrics = @{
            availability = $availability.Range("L26").Value2
            total_portfolio_par = $availability.Range("L31").Value2
            aggregate_adjusted_bv = $availability.Range("L33").Value2
            excess_concentration = $availability.Range("L34").Value2
            net_adjusted_bv = $availability.Range("L35").Value2
            credit_enhancement_test = $availability.Range("L42").Text
            weighted_avg_advance_rate = $availability.Range("L48").Value2
            current_advances = $availability.Range("L51").Value2
        }
        concentration_limits = $limits
    }
}
catch {
    Emit-Json @{
        status = "error"
        workbook_path = $WorkbookPath
        opened_workbook_path = $openedPath
        message = $_.Exception.Message
    }
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
