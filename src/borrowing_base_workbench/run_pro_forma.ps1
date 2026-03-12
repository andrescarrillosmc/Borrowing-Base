param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [Parameter(Mandatory = $true)]
    [string]$ScenarioPath
)

$ErrorActionPreference = "Stop"

function Emit-Json($Object) {
    $Object | ConvertTo-Json -Depth 8 -Compress
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
    $stagedPath = Join-Path $scratchDir ("pro_forma_" + $fileName + "_" + $timestamp + $extension)
    Copy-Item -LiteralPath $SourcePath -Destination $stagedPath -Force
    return $stagedPath
}

function Get-FirstBlankRow {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][string]$Column,
        [Parameter(Mandatory = $true)][int]$StartRow,
        [Parameter(Mandatory = $true)][int]$EndRow
    )

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        if ([string]::IsNullOrWhiteSpace($Worksheet.Range("$Column$row").Text)) {
            return $row
        }
    }
    throw ("No blank row available in {0}!{1}{2}:{1}{3}" -f $Worksheet.Name, $Column, $StartRow, $EndRow)
}

function Capture-ConcentrationSnapshot {
    param(
        [Parameter(Mandatory = $true)]$DealTeam,
        [Parameter(Mandatory = $true)]$Concentration
    )

    $rows = @(
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
    foreach ($row in $rows) {
        $limits += @{
            limit_type = [string]$DealTeam.Range("A$($row.DtiRow)").Text
            limit_percent = $DealTeam.Range("B$($row.DtiRow)").Value2
            applicable_limit = $DealTeam.Range("C$($row.DtiRow)").Value2
            actual = $Concentration.Range("L$($row.ClRow)").Value2
            excess = $Concentration.Range("K$($row.ClRow)").Value2
        }
    }
    return $limits
}

function Set-OptionalNumericCell {
    param(
        [Parameter(Mandatory = $true)]$Range,
        [Parameter(Mandatory = $false)]$Value
    )

    if ([string]::IsNullOrWhiteSpace([string]$Value)) {
        $Range.ClearContents() | Out-Null
    }
    else {
        $Range.Value2 = [double]$Value
    }
}

$excel = $null
$workbook = $null
$openedPath = $WorkbookPath

try {
    $scenario = Get-Content -LiteralPath $ScenarioPath -Raw | ConvertFrom-Json

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $openedPath = New-StagedWorkbookCopy -SourcePath $WorkbookPath
    $workbook = $excel.Workbooks.Open($openedPath)

    $availability = $workbook.Worksheets("Availability")
    $dealTeam = $workbook.Worksheets("Deal Team Input")
    $smSupport = $workbook.Worksheets("SM Support")
    $loanTape = $workbook.Worksheets("Loan Tape - Settled")
    $portfolio = $workbook.Worksheets("Portfolio")
    $concentration = $workbook.Worksheets("Concentration Limits")

    $beforeAvailability = $availability.Range("L26").Value2
    $beforeAggregateAdjustedBv = $availability.Range("L33").Value2
    $beforeExcessConcentration = $availability.Range("L34").Value2
    $beforeNetAdjustedBv = $availability.Range("L35").Value2
    $beforeCreditEnhancement = $availability.Range("L42").Text
    $beforeWeightedAvgAdvanceRate = $availability.Range("L48").Value2
    $beforeCurrentAdvances = $availability.Range("L51").Value2
    $beforeConcentration = Capture-ConcentrationSnapshot -DealTeam $dealTeam -Concentration $concentration

    $smRow = Get-FirstBlankRow -Worksheet $smSupport -Column "B" -StartRow 3 -EndRow 63
    $loanRow = Get-FirstBlankRow -Worksheet $loanTape -Column "G" -StartRow 5 -EndRow 55
    $portfolioRow = Get-FirstBlankRow -Worksheet $portfolio -Column "A" -StartRow 9 -EndRow 59

    $smSupport.Range("A$($smRow - 1):O$($smRow - 1)").Copy($smSupport.Range("A$smRow:O$smRow"))
    $loanTape.Range("A$($loanRow - 1):BJ$($loanRow - 1)").Copy($loanTape.Range("A$loanRow:BJ$loanRow"))
    $portfolio.Range("A$($portfolioRow - 1):HN$($portfolioRow - 1)").Copy($portfolio.Range("A$portfolioRow:HN$portfolioRow"))

    $dealTeam.Range("B7").Value2 = $scenario.company_name
    $dealTeam.Range("B8").Value2 = $scenario.security_type
    $dealTeam.Range("B14").Value2 = [double]$scenario.ltm_revenue
    $dealTeam.Range("B15").Value2 = [double]$scenario.ltm_adj_ebitda
    $dealTeam.Range("B16").Value2 = [double]$scenario.drawn_revolver
    $dealTeam.Range("B17").Value2 = [double]$scenario.first_out_balance
    $dealTeam.Range("B18").Value2 = [double]$scenario.pari_passu
    $dealTeam.Range("B19").Value2 = [double]$scenario.bdc_balance
    $dealTeam.Range("B20").Value2 = [double]$scenario.total_sm_balance
    $dealTeam.Range("B21").Value2 = [double]$scenario.cash_balance
    Set-OptionalNumericCell -Range $dealTeam.Range("B22") -Value $scenario.interest_coverage
    $dealTeam.Range("B28").Value2 = $scenario.loan_denomination
    $dealTeam.Range("B29").Value2 = [double]$scenario.purchase_price
    $dealTeam.Range("B30").Value2 = $scenario.country
    $dealTeam.Range("B31").Value2 = $scenario.investment_date
    $dealTeam.Range("B32").Value2 = $scenario.maturity_date
    $dealTeam.Range("B33").Value2 = $scenario.rate_type
    $dealTeam.Range("B34").Value2 = $scenario.industry_classification
    $dealTeam.Range("B35").Value2 = $scenario.payment_frequency
    $dealTeam.Range("B36").Value2 = [double]$scenario.pik_pct
    Set-OptionalNumericCell -Range $dealTeam.Range("B37") -Value $scenario.spread
    Set-OptionalNumericCell -Range $dealTeam.Range("B38") -Value $scenario.sofr_floor
    Set-OptionalNumericCell -Range $dealTeam.Range("B39") -Value $scenario.attach_point

    $smSupport.Range("B$smRow").Value2 = $scenario.company_name
    $smSupport.Range("C$smRow").Value2 = $scenario.security_type
    $smSupport.Range("D$smRow").Value2 = [double]$scenario.bdc_balance
    $smSupport.Range("E$smRow").Value2 = [double]$scenario.bdc_balance
    $smSupport.Range("F$smRow").Value2 = [double]$scenario.ltm_revenue
    $smSupport.Range("G$smRow").Value2 = [double]$scenario.ltm_adj_ebitda
    $smSupport.Range("H$smRow").Value2 = [double]$scenario.drawn_revolver
    $smSupport.Range("I$smRow").Value2 = [double]$scenario.first_out_balance
    $smSupport.Range("J$smRow").Value2 = [double]$scenario.pari_passu
    $smSupport.Range("K$smRow").Value2 = [double]$scenario.total_sm_balance
    $smSupport.Range("L$smRow").Value2 = [double]$scenario.cash_balance
    Set-OptionalNumericCell -Range $smSupport.Range("O$smRow") -Value $scenario.interest_coverage

    $securityDisplay = if ($scenario.security_type -eq "First Lien") { "1st Lien" } else { $scenario.security_type }

    $loanTape.Range("G$loanRow").Value2 = $scenario.company_name
    $loanTape.Range("H$loanRow").Value2 = $securityDisplay
    $loanTape.Range("J$loanRow").Value2 = $scenario.security_type
    $loanTape.Range("K$loanRow").Value2 = "Term Loan"
    $loanTape.Range("L$loanRow").Value2 = $scenario.loan_denomination
    $loanTape.Range("Q$loanRow").Value2 = [double]$scenario.purchase_price
    $loanTape.Range("T$loanRow").Value2 = $scenario.country
    $loanTape.Range("U$loanRow").Value2 = $scenario.investment_date
    $loanTape.Range("V$loanRow").Value2 = $scenario.maturity_date
    $loanTape.Range("W$loanRow").Value2 = $scenario.rate_type
    $loanTape.Range("X$loanRow").Value2 = $scenario.industry_classification
    $loanTape.Range("Y$loanRow").Value2 = $scenario.payment_frequency
    $loanTape.Range("AD$loanRow").Value2 = [double]$scenario.pik_pct
    Set-OptionalNumericCell -Range $loanTape.Range("AE$loanRow") -Value $scenario.spread
    Set-OptionalNumericCell -Range $loanTape.Range("AF$loanRow") -Value $scenario.sofr_floor
    $loanTape.Range("AO$loanRow").Value2 = $scenario.investment_date

    $excel.CalculateFullRebuild()

    $afterAvailability = $availability.Range("L26").Value2
    $afterConcentration = Capture-ConcentrationSnapshot -DealTeam $dealTeam -Concentration $concentration

    $eligibility = $portfolio.Range("P$portfolioRow").Text
    $failedTests = @()
    if ($eligibility -ne "Yes") {
        for ($col = 5; $col -le 14; $col++) {
            $header = [string]$portfolio.Cells(8, $col).Text
            $value = [string]$portfolio.Cells($portfolioRow, $col).Text
            if ($value -ne "TRUE" -and $value -ne "True") {
                $failedTests += $header
            }
        }
    }

    Emit-Json @{
        status = "ok"
        workbook_path = $WorkbookPath
        opened_workbook_path = $openedPath
        before = @{
            availability = $beforeAvailability
            aggregate_adjusted_bv = $beforeAggregateAdjustedBv
            excess_concentration = $beforeExcessConcentration
            net_adjusted_bv = $beforeNetAdjustedBv
            credit_enhancement_test = $beforeCreditEnhancement
            weighted_avg_advance_rate = $beforeWeightedAvgAdvanceRate
            current_advances = $beforeCurrentAdvances
            concentration_limits = $beforeConcentration
        }
        after = @{
            availability = $afterAvailability
            aggregate_adjusted_bv = $availability.Range("L33").Value2
            excess_concentration = $availability.Range("L34").Value2
            net_adjusted_bv = $availability.Range("L35").Value2
            credit_enhancement_test = $availability.Range("L42").Text
            weighted_avg_advance_rate = $availability.Range("L48").Value2
            current_advances = $availability.Range("L51").Value2
            concentration_limits = $afterConcentration
        }
        scenario = @{
            sm_support_row = $smRow
            loan_tape_row = $loanRow
            portfolio_row = $portfolioRow
        }
        eligibility = @{
            status = [string]$eligibility
            failed_tests = $failedTests
        }
    }
}
catch {
    Emit-Json @{
        status = "error"
        workbook_path = $WorkbookPath
        opened_workbook_path = $openedPath
        message = $_.Exception.Message
        position = $_.InvocationInfo.PositionMessage
        stack = $_.ScriptStackTrace
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
