# . .\trs_sample_01.ps1

Set-StrictMode -Version 3.0
#このスクリプトは別環境での動作検証用

# サンプルなのでべた書き
$applicant_company_name = "ES-2023"
$applicant_name_katakana = "ゼヒモナシ"
$applicant_name = "是非もなし"
$applicant_num = "00-000001"

$excel = New-Object -ComObject Excel.Application
#$excel.Visible = $False
#$excel.DisplayAlerts = $False

$temp_path = ".\data\from_T\xlsx\ed_cd.xlsx"
#$xlsx_path = (Get-ChildItem $temp_path).FullName
$xlsx_path = Get-ChildItem $temp_path
$book = $excel.Workbooks.Open($xlsx_path, 0, $true)


$sheet = $book.Sheets("c(TS)")
$sheet.Cells.Item(13, 1) = $applicant_company_name
$sheet.Cells.Item(13, 5) = $applicant_name_katakana
$sheet.Cells.Item(14, 5) = $applicant_name
$sheet.Cells.Item(13, 8) = $applicant_num

#関数へ切り分けたいね
$From = 1
$To = 1
$Copies = 1
$book.PrintOut.Invoke(@($From, $To, $Copies))

# bookの変更内容を保存しない
$book.Close($False)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
