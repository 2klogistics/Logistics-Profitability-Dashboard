$ErrorActionPreference = "Stop"
$csvPath = "C:\Users\ADMIN\Desktop\Data sum Daily EXPRESS 4 months\merged_all_months.csv"
$outputPath = "C:\Users\ADMIN\Desktop\Data sum Daily EXPRESS 4 months\dashboard\benchmark_data.js"

Write-Host "Reading CSV file..."
$rawData = Get-Content $csvPath | Select-Object -Skip 2 | ConvertFrom-Csv

Write-Host "Grouping data by Route+Month -> Driver..."
$routes = @{}
$mths = @('January','February','March','April','May','June','July','August','September','October','November','December')

foreach ($row in $rawData) {
    if ($row.customer -match "FLASH") {
        $row.customer = "FASH"
    }

    $pay = 0; $oil = 0; $pct = 0; $recv = 0; $margin = 0
    if ([double]::TryParse($row.price_pay_thb, [ref]$pay)) {}
    if ([double]::TryParse($row.oil_advance_thb, [ref]$oil)) {}
    if ([double]::TryParse($row.profit_pct, [ref]$pct)) {}
    if ([double]::TryParse($row.price_receive_thb, [ref]$recv)) {}
    if ([double]::TryParse($row.margin_thb, [ref]$margin)) {}

    # Skip records missing essential pricing data to prevent skewing benchmark averages
    if ($pay -eq 0 -or $recv -eq 0) {
        continue
    }

    # 1. Get Month Key first
    $m = $row.month
    if ([string]::IsNullOrWhiteSpace($m)) {
        $parts = $row.date.split('-')
        $mIdx = if ($parts.Length -ge 2) { [int]$parts[1] - 1 } else { 0 }
        if ($mIdx -ge 0 -and $mIdx -lt 12) { $m = $mths[$mIdx] } else { $m = "Unknown" }
    }
    if ([string]::IsNullOrWhiteSpace($m)) { $m = "Unknown" }

    # 2. Route+Month Key
    $rk = "$($row.customer)|$($row.vehicle_type)|$($row.route_name)|$m"
    if (-not $routes.ContainsKey($rk)) {
        $routes[$rk] = @{
            customer = $row.customer
            vtype = $row.vehicle_type
            route = $row.route_name
            routeDesc = $row.route_description
            month = $m
            drivers = @{}
            totalTrips = 0
            payCount = 0
            sumPay = 0
            sumOil = 0
            sumMargin = 0
        }
    }
    
    # 3. Driver Key
    $dk = "$($row.driver_name)|$($row.plate_number)"
    if (-not $routes[$rk].drivers.ContainsKey($dk)) {
        $routes[$rk].drivers[$dk] = @{
            driver = $row.driver_name
            plate = $row.plate_number
            totalTrips = 0
            payCount = 0
            sumPay = 0
            sumOil = 0
            sumMargin = 0
            maxPay = -1
            minPay = 999999999
            records = @()
        }
    }
    
    # Variables already parsed at the start of the loop
    
    $dr = $routes[$rk].drivers[$dk]
    $dr.totalTrips++
    $dr.sumPay += $pay
    $dr.sumOil += $oil
    $dr.sumMargin += $margin
    
    if ($pay -gt 0) {
        $dr.payCount++
        $routes[$rk].payCount++
    }
    
    if ($pay -gt $dr.maxPay) { $dr.maxPay = $pay }
    if ($pay -gt 0 -and $pay -lt $dr.minPay) { $dr.minPay = $pay }
    
    $dr.records += @{
        date = $row.date
        payee = $row.contractor
        recv = $recv
        pay = $pay
        oil = $oil
        margin = $margin
    }

    $routes[$rk].totalTrips++
    $routes[$rk].sumPay += $pay
    $routes[$rk].sumOil += $oil
    $routes[$rk].sumMargin += $margin
}

# Convert and Calculate Averages
$outData = @()
foreach ($rk in $routes.Keys) {
    $r = $routes[$rk]
    if ($r.totalTrips -gt 0) {
        $r.avgPay = if ($r.payCount -gt 0) { $r.sumPay / $r.payCount } else { 0 }
        $r.avgOil = $r.sumOil / $r.totalTrips
        $r.avgMargin = $r.sumMargin / $r.totalTrips
    } else {
        $r.avgPay = 0; $r.avgOil = 0; $r.avgMargin = 0
    }

    $driverArr = @()
    foreach ($dk in $r.drivers.Keys) {
        $d = $r.drivers[$dk]
        if ($d.totalTrips -gt 0) {
            $d.avgPay = if ($d.payCount -gt 0) { $d.sumPay / $d.payCount } else { 0 }
            $d.avgOil = $d.sumOil / $d.totalTrips
            $d.avgMargin = $d.sumMargin / $d.totalTrips
        } else {
            $d.avgPay = 0; $d.avgOil = 0; $d.avgMargin = 0
        }
        $driverArr += $d
    }
    $r.drivers = @($driverArr | Sort-Object -Property totalTrips -Descending)
    $outData += $r
}

$outData = @($outData | Sort-Object { 
    $mIdx = $mths.IndexOf($_.month)
    if ($mIdx -lt 0) { 99 } else { $mIdx }
}, @{Expression="totalTrips"; Descending=$true})

Write-Host "Found $($outData.Count) unique route-month combinations."
$json = $outData | ConvertTo-Json -Depth 10 -Compress
$jsContent = "const BENCHMARK_DATA = $json;"
[System.IO.File]::WriteAllText($outputPath, $jsContent, [System.Text.Encoding]::UTF8)
Write-Host "Done! Saved to $outputPath"
