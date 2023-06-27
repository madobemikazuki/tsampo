Set-StrictMode -Version latest

Set-Variable -name DELIMITER -value ',' -option Constant 


try {
  $header = Get-Content -Path .\header.txt -Encoding UTF8 -ErrorAction Stop
  $values = Get-Content -Path .\values.txt -Encoding UTF8 -ErrorAction Stop

  $csv_object = $values | ConvertFrom-Csv -Header $header

  $csv_object | Export-Csv -NoTypeInformation -Path .\exported.csv -Delimiter $DELIMITER -Encoding UTF8
}
catch {
  Write-Host "エラーが発生しました：$($_.Exception.Message)"
}

