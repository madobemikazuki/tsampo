Set-StrictMode -Version latest

Set-Variable -name DELIMITER -value ',' -option Constant 

$header = Get-Content -Path .\header.txt -Encoding UTF8
$values = Get-Content -Path .\values.txt -Encoding UTF8

$csv_object = $values | ConvertFrom-Csv -Header $header

$csv_object | Export-Csv -NoTypeInformation -Path .\exported.csv -Delimiter $DELIMITER -Encoding UTF8

