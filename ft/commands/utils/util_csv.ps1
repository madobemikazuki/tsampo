Set-StrictMode -Version Latest

function bind_as_csv {
  Param(
    [Object[]]$header,
    [Object[]]$values,
    [String]$delimiter = ','
  )
  [PSCustomObject[]]$csv_object =  $values | ConvertFrom-Csv -Header $header.Split($delimiter)
  return $csv_object
}

function export_csv {
  Param(
    [Parameter(Mandatory = $true)][Object[]]$_csv_obj,
    [Parameter(Mandatory = $true)][String]$_path,
    [String]$_delimiter = ',',
    [String]$_encode = "UTF8"#Default ではブラウザで参照すると文字化けする。
  )
  $_csv_obj | Export-Csv -NotypeInformation -Path $_path -Delimiter $_delimiter -Encoding $_encode -Force
}

function remove_quauts{
  
}