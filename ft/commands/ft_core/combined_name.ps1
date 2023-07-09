Set-StrictMode -Version 3.0


set-Variable Z_BLANC '　' -option Constant
set-Variable UNDER_SCORE '_' -option Constant 
function combined_name {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)][String]$first_name,
        [Parameter(Mandatory)][String]$last_name,
        #デフォルト引数
        [String]$blanc = $Z_BLANC
    )
    $sb = New-Object System.Text.StringBuilder

    #副作用処理  StringBuilderならちょっと速いらしい。要素数が少ないから意味ないかも。
    @($first_name, $blanc ,$last_name) |
    ForEach-Object{[void] $sb.Append($_)}

    $sb.ToString()
    #初期の実装コード
    #[system.String]::Concat($first_name, $blanc ,$last_name)

}

function one_liner {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)][PSCustomObject[]]$name_list
    )
    return $name_list -join $UNDER_SCORE 
}

exit 0