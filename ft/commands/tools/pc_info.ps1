Set-StrictMode -Version 3.0

function pc_info{
  [CmdletBinding()]

  $message = "--- このPCの情報 ---"
  Write-Output $message
  os_info
  cpu_info
  pmem_GiB
}

function os_info {
  $os_infos = Get-WmiObject Win32_OperatingSystem
  [String]"OS エディション: " + $os_infos.Caption
  [String]"OS バージョン: " +$os_infos.Version
}

function cpu_info {
  $cpu_infos = Get-WmiObject Win32_Processor
  [String]"CPU名: " + $cpu_infos.Caption
  [String]"CPU名: "+$cpu_infos.Name
  [String]"コア数: "+$cpu_infos.NumberOfCores
  [String]"スレッド数: "+$cpu_infos.NumberOfLogicalProcessors
  [String]"最大クロック数: " + $cpu_infos.MaxClockSpeed + "Mhz"
}

#PhysicalMemory をGiB表示で返す
function pmem_GiB{
  $mem = Get-WmiObject Win32_PhysicalMemory | Select-Object -Property Capacity
  [String]$v = ($mem[0].Capacity / 1024 / 1024 /1024)
  return "物理メモリ: " + $v + " GiB"
}

exit 0