Set-StrictMode -Version Latest


function find_file_name {
  Param(
    [Parameter(Mandatory = $true)]
    [String]$dic,
    [Parameter(Mandatory = $true)]
    [String]$target
  )
  # ファイルが存在しないときのエラー処理がわからない。
  try {
    $file_name = Get-childItem -Path $dic -Name -File -Include $target
    return $file_name
  }
  catch {
    Write-Host "エラー発生 :: $($_.Exception.Message)"
  }
}

function read_to_array {
  Param(
    [Parameter(Mandatory = $true)][String]$txt_path,
    [String]$encode = "Default"
  )
  
  return Get-Content -Encoding $encode $txt_path
}
