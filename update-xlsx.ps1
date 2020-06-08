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
$colMap = @{ "Totalt_antal_fall" = 2; "Norrbotten" = 4; "Halland" = 6; "Västra_Götaland" = 8; "Stockholm" = 10 }

$queryUrl = "https://services5.arcgis.com/fsYDFeRKu1hELJJs/arcgis/rest/services/FOHM_Covid_19_FME_1/FeatureServer/1/query?f=json&where=1%3D1&returnGeometry=false&spatialRel=esriSpatialRelIntersects&outFields=*&orderByFields=Statistikdatum%20desc&outSR=102100&resultOffset=0&resultRecordCount=32000&resultType=standard&cacheHint=true"


$r = Invoke-WebRequest -UseBasicParsing $queryUrl
$j = ConvertFrom-Json $r.Content

$excel = New-Object -ComObject Excel.Application
$curRow = 3
try {
    $excel.Visible = $true
    $excel.ScreenUpdating = $False 
    $excelWb = $excel.Workbooks.Open($xlsx)
    $excelSheet = $excelWb.ActiveSheet
    
    $j.features | %{
        $curRow++
        $_.attributes.PSObject.Properties | % {
            if ($_.Name -eq "Statistikdatum") {
                $date = (Get-Date 01.01.1970) + ([System.TimeSpan]::FromMilliseconds($_.Value))
                $excelSheet.Cells.Item($curRow, 1) = $date.ToOADate()
            } elseif ($colMap.Contains($_.Name)) {
                $excelSheet.Cells.Item($curRow, $colMap[$_.Name]) = $_.Value
            }
        }
    }
    
    $excelChart = $excelSheet.ChartObjects("Neuinfektionen").Chart
    $excelChartXValues = $excelChart.SeriesCollection(1).XValues
    
    $wc = [Math]::floor($excelChartXValues.Length / 7)
    $lastOaDate = $excelChartXValues.Get(1)
    $lastDate = [System.DateTime]::FromOaDate($lastOaDate)
    $daysToNextMonday = (8 - $lastDate.DayOfWeek) % 7
    
    $excelChart.Axes(1).MinimumScale = $lastOaDate + $daysToNextMonday - $wc * 7
    $excelChart.Axes(1).MaximumScale = $lastOaDate + $daysToNextMonday
    
    $excelSheet.Cells.Item(1, 1) = "Last Update: $(Get-Date)"
    $excelWb.Save()
} finally {
    $excel.ScreenUpdating = $True
    $excel.Quit()
}