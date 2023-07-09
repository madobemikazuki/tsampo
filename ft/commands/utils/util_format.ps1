Set-StrictMode -Version Latest

function post_code {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^\d{7}")][String]$arg
  )
  return Write-Output("{0:000-0000}" -f [Int]$arg)
}
