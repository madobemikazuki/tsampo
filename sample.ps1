Set-StrictMode -Version 3.0
#このスクリプトは別環境での動作検証用

# 所属,氏名（カタカナ）,氏名,中登
# 13:1,13:5,14:5,13:8

# サンプルなのでべた書き
$applicant_company_name = "ES-2023"
$applicant_name_kakakana = "ゼヒモナシ"
$applicant_name = "是非もなし"
$applicant_num = "00-000000"

#$address_table = Import-Csv $csv_path -Encoding utf8


#New-Object -ComObject でCOMオブジェクトを使用。
$excel = New-Object -ComObject Excel.Application

#.Visible = $false でExcelを表示しないで処理を実行できる。
$excel.Visible = $False

# 上書き保存時に表示されるアラートなどを非表示にする
$excel.DisplayAlerts = $False

$temp_path = ".\data\from_T\excel\c_ed.xlsx"
$xlsx_path = (Get-ChildItem $temp_path).FullName
$book = $excel.Workbooks.Open($xlsx_path, 0, $true)
$sheet = $book.Sheets("c(TS)")

$sheet.Cells.Item(13, 1) = $applicant_company_name
$sheet.Cells.Item(13, 5) = $applicant_name_kakakana
$sheet.Cells.Item(14, 5) = $applicant_name
$sheet.Cells.Item(13, 8) = $applicant_num

#関数へ切り分ける
$From = 1
$To = 1
$Copies = 1
$book.PrintOut.Invoke(@($From, $To, $Copies))



$book.Close()

$excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
