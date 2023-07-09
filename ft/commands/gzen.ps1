Set-StrictMode -Version Latest

# 必要な情報は設定ファイルとして読み込む init()が必要
Set-Variable -Name DLMT_comma -Value "," -Option Constant
Set-Variable -Name TARGET_NAME -Value "*事前申請書*.txt" -Option Constant
Set-Variable -Name GZEN_HEADER -Value ".\data\header\gZEN_header_ANSI.txt" -Option Constant
Set-Variable -Name SELECTED_LIST -Value ".\data\header\gZen_select_items_ANSI.txt" -Option Constant
Set-Variable -Name OUTPUT_DIC -Value "$HOME\downloads" -Option Constant
Set-Variable -Name OUTPUT_FILE_NAME -Value "\export.csv" -Option Constant
Set-Variable -Name WEB_APP_PATH -Value "$HOME\apps\cpy\cpy.html" -Option Constant
Set-Variable -Name C_NAME -Value "gZEN" -Option Constant

function notifycation {
  Param(
    [String]$title,
    [String]$message
  )
  . .\commands\utils\notify.ps1 
  notify_balloon $title $message
}

function shape_values {
  Param(
    [PSObject[]]$arg,
    [String[]]$list
  )
  . .\commands\ft_core\combined_name.ps1
  . .\commands\utils\util_format.ps1
  #プロパティの追加
  $new_obj = $arg | Select-Object *, @{
    Name       = "カタカナ氏名";
    Expression = { combined_name $_.'カナ氏名（姓）' $_.'カナ氏名（名）' }
  }
  $new_obj = $new_obj | Select-Object *, @{
    Name       = '漢字氏名';
    Expression = { combined_name $_.'漢字氏名（姓）' $_.'漢字氏名（名）' }
  }

  foreach ($_ in $new_obj) {
    $_.'現住所（住民票）郵便番号' = post_code $_.'現住所（住民票）郵便番号'
    $_.'現住所（現在住んでいる）郵便番号' = post_code $_.'現住所（現在住んでいる）郵便番号'
    $_.psobject.properties.remove('カナ氏名（姓）')
    $_.psobject.properties.remove('カナ氏名（名）')
  }
  # 出力プロパティを任意の文字列[]で取得する。
  $selected_obj = $new_obj | Select-Object -Property $list
  return $selected_obj
}

try {
  . .\commands\utils\util_txt.ps1
  $header = read_to_array $GZEN_HEADER
  # *事前申請書*.txt はSHIFT-JISでエンコードされている。
  $values_filename = find_file_name $OUTPUT_DIC $TARGET_NAME
  $values = read_to_array "$OUTPUT_DIC\$values_filename"

  #テキストファイルを読み込み、ヘッダーをつけてCSVファイルを生成。
  . .\commands\utils\util_csv.ps1
  $csv_obj = bind_as_csv $header $values

  # 必要な項目の情報を抽出する
  $list = read_to_array $SELECTED_LIST
  $selected_csv_obj = $csv_obj | Select-Object -Property $list

  $output_list = read_to_array ".\data\header\gZen_output_items_ANSI.txt"
  $new_csv_obj = shape_values $selected_csv_obj $output_list
  #CSVファイルを出力
  export_csv $new_csv_obj "$OUTPUT_DIC/$OUTPUT_FILE_NAME" $DLMT_comma

  #clipboardに出力する
  $plain_text = (Get-Content -Path "$OUTPUT_DIC/$OUTPUT_FILE_NAME" -Encoding UTF8)
  $formatted_text = $plain_text.Replace('"', '')
  $formatted_text.Replace($DLMT_comma, "`t") | Set-Clipboard

  # $values＿file_name の削除
  #Remove-Item -Path $values_filename

  # EDGEブラウザ起動し、cpyを起動する。
  Start-Process msedge $WEB_APP_PATH


  # 通知を表示
  notifycation $C_NAME "🐈.,💩💩,,.  💩,  🌲🏡"
}
catch {
  Write-Host "エラー発生 :: $($_.Exception.Message)"
  notifycation "エラー発生 ::" "$($_.Exception.Message)"
}
