Set-StrictMode -Version Latest

# 必要な情報は設定ファイルとして読み込む init()が必要
Set-Variable -Name DLMT_comma -Value "," -Option Constant
Set-Variable -Name TARGET_NAME -Value "*事前申請書*.txt" -Option Constant
Set-Variable -Name GZEN_HEADER -Value ".\src\tmp\gZEN_header.txt" -Option Constant
Set-Variable -Name OUTPUT_DIC -Value "$HOME\downloads" -Option Constant
Set-Variable -Name OUTPUT_TEMP -Value "$OUTPUT_DIC\temp.csv" -Option Constant
Set-Variable -Name OUTPUT_FILE_NAME -Value "\export.csv" -Option Constant
Set-Variable -Name WEB_APP_PATH -Value "$HOME\apps\cpy\cpy.html" -Option Constant
Set-Variable -Name C_NAME -Value "gZEN" -Option Constant
Set-Variable -Name NOTIFY_MESSAGE -Value "🐈.,💩. 💩,,.  💩,,   🌲🏡" -Option Constant 

function notifycation {
  param(
    [String]$title,
    [String]$message
  )
  . .\src\notify.ps1 $title $message
}

function create_csv {
  Param(
    [Parameter(Mandatory = $true)]
    [String[]]$header,
    [Parameter(Mandatory = $true)]
    [String[]]$values
  )
}

function ascii_to_utf8 {
  Param(
    [Parameter(Mandatory = $true)]
    [String]$_path
  )
  try {
    $output = "$HOME\downloads\temp.txt"
    Get-Content -Encoding Default -Path $_path | Set-Content -Encoding UTF8 -Path $output -Force
    $text_utf8 = Get-Content -Encoding UTF8 -Path $output
    Remove-Item -Path $output
    return $text_utf8
  }
  catch {
    Write-Host "エラー発生 :: $($_.Exception.Message)"
    notifycation "エラー発生 ::" "$($_.Exception.Message)"
  }
}

function export_utf8_csv {
  Param(
    [Parameter(Mandatory = $true)]
    [Object[]]$csv_obj,
    [Parameter(Mandatory = $true)]
    [String]$_path,
    [String]$_delimiter = ','
  )
  $csv_obj | Export-Csv -NotypeInformation -Delimiter $_delimiter -Encoding UTF8 -Path $_path -Force
}

function post_code {
  Param(
    [Parameter(Mandatory = $true)]
    [String]$_num
  )
  return Write-Output("{0:000-0000}" -f [Int]$_num)
}

function combined_name {
  Param(
    [Parameter(Mandatory = $true)]
    [String]$_first_name,
    [Parameter(Mandatory = $true)]
    [String]$_last_name,
    [String]$_delimiter = "　"
  )
  return $_first_name + $_delimiter + $_last_name
}

try {
  $header = Get-Content -Path $GZEN_HEADER -Encoding UTF8

  # *事前申請書*.txt はSHIFT-JISでエンコードされている。
  $values_filename = Get-childItem -Path . -Name -File -Include $TARGET_NAME
  $csv_obj = ascii_to_utf8 ".\$values_filename" | ConvertFrom-Csv -Header $header.Split($DLMT_comma)
  export_utf8_csv $csv_obj "$OUTPUT_DIC/$OUTPUT_FILE_NAME" $DLMT_comma

  #clipboardに出力する
  $a = (Get-Content -Path "$OUTPUT_DIC/$OUTPUT_FILE_NAME" -Encoding UTF8).Replace('"', '')
  $a.Replace($DLMT_comma, "`t") | Set-Clipboard

  # $values＿file_name の削除
  #Remove-Item -Path $values_filename

  # EDGEブラウザ起動し、cpyを起動する。
  Start-Process msedge $WEB_APP_PATH

  # 通知を表示
  notifycation $C_NAME $NOTIFY_MESSAGE
}
catch {
  Write-Host "エラー発生 :: $($_.Exception.Message)"
  notifycation "エラー発生 ::" "$($_.Exception.Message)"
}



