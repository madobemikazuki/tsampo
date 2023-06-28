Set-StrictMode -Version Latest


Set-Variable -Name DELIMITER -Value ',' -Option Constant
try {
  $header = Get-Content .\src\gZen_header.txt

  # *なまえ*.txt はSHIFT-JISでエンコードされている。
  $source_values = Get-childItem -Path . -Name -File -Include "*なまえ*.txt"
  Get-Content -Encoding Default -Path ".\$source_values" | Set-Content -Encoding UTF8 -Path ".\samples\a.txt"
  $values = Get-Content ".\samples\a.txt"

  $csv_object = $values | ConvertFrom-Csv -Header $header.Split($DELIMITER)
  # $csv_object.gettype()
  # $csv_object | ft
  $csv_object | Export-Csv -NoTypeInformation -Path ".\exported.csv" -Delimiter $DELIMITER -Encoding UTF8
}
catch {
  Write-Host "エラー発生 :: $($_.Exception.Message)"
}