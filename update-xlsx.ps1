# Copyright 2020 Philipp Serr (episource)
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
#     http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# Enable common parameters
[CmdletBinding()] 
Param(
)

$xlsx = $PSScriptRoot + "\swe-corona.xlsx"
$imgFile = $PSScriptRoot + "\swe-corona.png"
$imgWidth = $None

$sheetSwedishAgency = "Folkhälsomyndigheten"
$sheetEcdcSweden = "ECDC Schweden"
$sheetEcdcGermany = "ECDC Deutschland"
$sheetChart = "Schaubild"
$chartName = "Neuinfektionen"

$colMapSwedishAgency = @{ "Totalt_antal_fall" = 2; "Norrbotten" = 4; "Halland" = 6; "Västra_Götaland" = 8; "Stockholm" = 10 }

$queryUrlSwedishAgency = "https://services5.arcgis.com/fsYDFeRKu1hELJJs/arcgis/rest/services/FOHM_Covid_19_FME_1/FeatureServer/1/query?f=json&where=1%3D1&returnGeometry=false&spatialRel=esriSpatialRelIntersects&outFields=*&orderByFields=Statistikdatum%20desc&outSR=102100&resultOffset=0&resultRecordCount=32000&resultType=standard&cacheHint=true"
$queryUrlEcdc = "https://opendata.ecdc.europa.eu/covid19/casedistribution/csv"


function Update-SwedishAgencySheet($sheet, $lastUpdate) {
    $r = Invoke-WebRequest -UseBasicParsing $queryUrlSwedishAgency
    $j = ConvertFrom-Json $r.Content
    
    $curRow = 3
    $j.features | %{
        $curRow++
        $_.attributes.PSObject.Properties | % {
            if ($_.Name -eq "Statistikdatum") {
                $date = (Get-Date 01.01.1970) + ([System.TimeSpan]::FromMilliseconds($_.Value))
                $sheet.Cells.Item($curRow, 1) = $date.ToOADate()
            } elseif ($colMapSwedishAgency.Contains($_.Name)) {
                $sheet.Cells.Item($curRow, $colMapSwedishAgency[$_.Name]) = $_.Value
            }
        }
    }
    
    $sheet.Cells.Item(1, 1) = $lastUpdate
}

function Update-EcdcSheet($sheet, $data, $country, $lastUpdate) {
    $curRow = 3
    $data | ?{
        $_.country -eq $country 
    } | %{
        $curRow++
        $sheet.Cells.Item($curRow, 1) = $_.date.ToOADate()
        $sheet.Cells.Item($curRow, 2) = $_.cases
        $sheet.Cells.Item($curRow, 4) = $_.population
    }
    
    $sheet.Cells.Item(1, 1) = $lastUpdate
}

function Get-DataFromEcdc() {
    $r = Invoke-WebRequest -UseBasicParsing $queryUrlEcdc
    
    return [String]::new($r.Content) | %{
        $_ -split "[\r\n]+" 
    } | Select-Object -Skip 1 | ?{
        $_.Length -gt 0
    } | %{
        $row = $_.Split(",")
        [PSCustomObject]@{ 
            "date"=[DateTime]::ParseExact($row[0], "dd/MM/yyyy", $null)
            "cases"=$row[4]
            "country"=$row[6]
            "population"=$row[9]
        }
    } 
}

$now = Get-Date -format "dddd yyyy-MM-dd HH:mm"
$lastUpdate = "Datenabruf: $now"

$excel = New-Object -ComObject Excel.Application
try {
    $excel.Visible = $true
    $excel.ScreenUpdating = $False 
    $excelWb = $excel.Workbooks.Open($xlsx)
    $excelSheetSwedishAgency = $excelWb.Sheets($sheetSwedishAgency)
    $excelSheetEcdcSweden = $excelWb.Sheets($sheetEcdcSweden)
    $excelSheetEcdcGermany = $excelWb.Sheets($sheetEcdcGermany)
    $excelSheetChart = $excelWb.Sheets($sheetChart)
    
    
    Update-SwedishAgencySheet $excelSheetSwedishAgency $lastUpdate
    
    $ecdcData = Get-DataFromEcdc
    Update-EcdcSheet $excelSheetEcdcSweden $ecdcData "sweden" $lastUpdate
    Update-EcdcSheet $excelSheetEcdcGermany $ecdcData "germany" $lastUpdate
    
    
    $excelChart = $excelSheetChart.ChartObjects($chartName)
    $excelChartXValues = $excelChart.Chart.SeriesCollection(1).XValues
    
    $wc = [Math]::floor($excelChartXValues.Length / 7)
    $lastOaDate = $excelChartXValues.Get(1)
    $lastDate = [System.DateTime]::FromOaDate($lastOaDate)
    $daysToNextMonday = (8 - $lastDate.DayOfWeek) % 7
    
    $excelChart.Chart.Axes(1).MinimumScale = $lastOaDate + $daysToNextMonday - $wc * 7
    $excelChart.Chart.Axes(1).MaximumScale = $lastOaDate + $daysToNextMonday
    $excelChart.Chart.ChartTitle.Text = "Neuinfektionen/100k (7 Tage) - $lastUpdate"
    
    $excel.ScreenUpdating = $True
    
    try {
        $excelChart.CopyPicture([Microsoft.Office.Interop.Excel.XlPictureAppearance]::xlScreen, [Microsoft.Office.Interop.Excel.XlCopyPictureFormat]::xlBitmap)
        
        $img = Get-Clipboard -Format Image
        if (-not $img) {
            throw "clipboard empty"
        }

        if ($imgWidth -eq $None) {
            $imgWidth = $img.Width
        }
        $imgHeight = [int]($img.Height / $img.Width * $imgWidth)
        $outBitmap = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $imgWidth, $imgHeight
        
        $outG = [System.Drawing.Graphics]::FromImage($outBitmap)
        $outG.SmoothingMode = "HighQuality"
        $outG.InterpolationMode = "HighQualityBicubic"
        $outG.PixelOffsetMode = "HighQuality"
        $outGRectangle = 
        $outG.DrawImage($img, [System.Drawing.Rectangle]::new(0, 0, $imgWidth, $imgHeight))
        
        $outBitmap.Save($imgFile)
    } catch {
        Write-Warning "Failed to save chart image: $_"
        throw
    }
    
    $excelWb.Save()
} finally {
    $excel.ScreenUpdating = $True
    $excel.Quit()
}

