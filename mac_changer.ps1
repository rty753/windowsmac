chcp 65001 | Out-Null
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Host.UI.RawUI.WindowTitle = "MAC 地址修改器"

# ============================================
#  注册表根路径
# ============================================
$regBase = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}"

# ============================================
#  格式化 MAC 地址: AABBCCDDEEFF -> AA-BB-CC-DD-EE-FF
# ============================================
function Format-MAC($mac) {
    if (-not $mac -or $mac.Length -ne 12) { return $mac }
    return ($mac -replace '(.{2})', '$1-').TrimEnd('-')
}

# ============================================
#  生成随机 MAC (单播 + 本地管理位)
# ============================================
function New-RandomMAC {
    $first = Get-Random -Minimum 0 -Maximum 256
    $first = ($first -band 0xFE) -bor 0x02   # bit0=0 单播, bit1=1 本地管理
    $bytes = @($first)
    for ($i = 0; $i -lt 5; $i++) {
        $bytes += Get-Random -Minimum 0 -Maximum 256
    }
    return ($bytes | ForEach-Object { '{0:X2}' -f $_ }) -join ''
}

# ============================================
#  获取物理网络适配器列表
# ============================================
function Get-PhysicalAdapters {
    $result = @()

    # 获取物理适配器 (排除虚拟/隧道等)
    $netAdapters = Get-NetAdapter | Where-Object {
        $_.HardwareInterface -eq $true -and
        $_.InterfaceDescription -notmatch 'Virtual|Miniport|Filter|WAN|Bluetooth|Debug|Kernel|Teredo|ISATAP|6to4|Hosted|Direct'
    }

    if (-not $netAdapters -or $netAdapters.Count -eq 0) {
        # 兜底：取所有非虚拟的
        $netAdapters = Get-NetAdapter | Where-Object { $_.HardwareInterface -eq $true }
    }

    foreach ($na in $netAdapters) {
        # 在注册表中找到对应子键
        $regPath = $null
        $subKeys = Get-ChildItem $regBase -ErrorAction SilentlyContinue
        foreach ($sk in $subKeys) {
            try {
                $props = Get-ItemProperty $sk.PSPath -ErrorAction SilentlyContinue
                if ($props.NetCfgInstanceId -eq $na.InterfaceGuid) {
                    $regPath = $sk.PSPath
                    break
                }
            } catch {}
        }

        if (-not $regPath) { continue }

        # 检查是否已有自定义MAC
        $customMAC = $null
        try {
            $customMAC = (Get-ItemProperty $regPath -Name 'NetworkAddress' -ErrorAction SilentlyContinue).NetworkAddress
        } catch {}

        # 获取原始硬件MAC (通过WMI)
        $originalMAC = ""
        try {
            $wmi = Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.GUID -eq $na.InterfaceGuid } | Select-Object -First 1
            if ($wmi -and $wmi.MACAddress) {
                $originalMAC = ($wmi.MACAddress -replace '[:-]', '').ToUpper()
            }
        } catch {}

        $currentMAC = ($na.MacAddress -replace '-', '').ToUpper()
        if (-not $originalMAC) { $originalMAC = $currentMAC }

        $result += [PSCustomObject]@{
            Name        = $na.Name
            Description = $na.InterfaceDescription
            Status      = $na.Status
            CurrentMAC  = $currentMAC
            OriginalMAC = $originalMAC
            RegPath     = $regPath
            HasCustom   = ($null -ne $customMAC -and $customMAC -ne '')
        }
    }

    return $result
}

# ============================================
#  显示适配器列表
# ============================================
function Show-Adapters($adapters) {
    Write-Host ""
    Write-Host "  当前网络适配器:" -ForegroundColor Yellow
    Write-Host "  ============================================"
    $i = 1
    foreach ($a in $adapters) {
        if ($a.Status -eq 'Up') {
            $statusText = "[已连接]"
            $statusColor = "Green"
        } else {
            $statusText = "[未连接]"
            $statusColor = "DarkGray"
        }
        Write-Host "  $i. " -NoNewline -ForegroundColor White
        Write-Host "$($a.Name) " -NoNewline -ForegroundColor Cyan
        Write-Host $statusText -ForegroundColor $statusColor
        Write-Host "     当前 MAC: $(Format-MAC $a.CurrentMAC)" -ForegroundColor White
        Write-Host "     原始 MAC: $(Format-MAC $a.OriginalMAC)" -ForegroundColor DarkGray
        if ($a.HasCustom) {
            Write-Host "     (已自定义修改过)" -ForegroundColor DarkYellow
        }
        $i++
    }
    Write-Host "  ============================================"
}

# ============================================
#  修改MAC并重启适配器
# ============================================
function Set-NewMAC($adapter, $newMAC) {
    $oldMAC = $adapter.CurrentMAC

    Write-Host ""
    Write-Host "  适配器 : $($adapter.Name)" -ForegroundColor Cyan
    Write-Host "  旧 MAC : $(Format-MAC $oldMAC)" -ForegroundColor Red
    Write-Host "  新 MAC : $(Format-MAC $newMAC)" -ForegroundColor Green
    Write-Host ""

    # 写入注册表
    Write-Host "  [1/3] 写入注册表..." -ForegroundColor Yellow
    try {
        Set-ItemProperty -Path $adapter.RegPath -Name 'NetworkAddress' -Value $newMAC -ErrorAction Stop
        Write-Host "  [1/3] 注册表写入成功" -ForegroundColor Green
    } catch {
        Write-Host "  [1/3] 注册表写入失败: $_" -ForegroundColor Red
        return $false
    }

    # 禁用适配器
    Write-Host "  [2/3] 禁用适配器（会短暂断网）..." -ForegroundColor Yellow
    try {
        Disable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
        Start-Sleep 2
        Write-Host "  [2/3] 适配器已禁用" -ForegroundColor Green
    } catch {
        Write-Host "  [2/3] 禁用失败，尝试 netsh..." -ForegroundColor Yellow
        netsh interface set interface $adapter.Name disabled 2>$null
        Start-Sleep 2
    }

    # 启用适配器
    Write-Host "  [3/3] 启用适配器..." -ForegroundColor Yellow
    try {
        Enable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
        Start-Sleep 3
        Write-Host "  [3/3] 适配器已启用" -ForegroundColor Green
    } catch {
        Write-Host "  [3/3] 启用失败，尝试 netsh..." -ForegroundColor Yellow
        netsh interface set interface $adapter.Name enabled 2>$null
        Start-Sleep 3
    }

    # 验证
    Write-Host ""
    try {
        $verify = (Get-NetAdapter -Name $adapter.Name -ErrorAction Stop).MacAddress -replace '-', ''
        if ($verify -eq $newMAC) {
            Write-Host "  ✔ 修改成功！当前 MAC: $(Format-MAC $verify)" -ForegroundColor Green
        } else {
            Write-Host "  ✔ 注册表已写入，当前 MAC: $(Format-MAC $verify)" -ForegroundColor Yellow
            Write-Host "    部分网卡需要重启电脑才能完全生效" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  ✔ 注册表已写入，请检查网络连接" -ForegroundColor Yellow
    }
    Write-Host "  ✔ 此 MAC 地址重启电脑后依然保持！" -ForegroundColor Green

    # 写日志
    $logFile = Join-Path $PSScriptRoot "mac_change_log.txt"
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logFile -Value "$ts  适配器=$($adapter.Name)  旧MAC=$(Format-MAC $oldMAC)  新MAC=$(Format-MAC $newMAC)" -Encoding UTF8

    return $true
}

# ============================================
#  恢复出厂MAC
# ============================================
function Restore-OriginalMAC($adapter) {
    $oldMAC = $adapter.CurrentMAC

    Write-Host ""
    Write-Host "  适配器 : $($adapter.Name)" -ForegroundColor Cyan
    Write-Host "  当前MAC: $(Format-MAC $oldMAC)" -ForegroundColor Red
    Write-Host "  原始MAC: $(Format-MAC $adapter.OriginalMAC)" -ForegroundColor Green
    Write-Host ""

    # 删除注册表中的自定义MAC
    Write-Host "  [1/3] 删除注册表自定义MAC..." -ForegroundColor Yellow
    try {
        Remove-ItemProperty -Path $adapter.RegPath -Name 'NetworkAddress' -ErrorAction SilentlyContinue
        Write-Host "  [1/3] 注册表已清除" -ForegroundColor Green
    } catch {
        Write-Host "  [1/3] 清除失败: $_" -ForegroundColor Red
    }

    # 重启适配器
    Write-Host "  [2/3] 禁用适配器..." -ForegroundColor Yellow
    try {
        Disable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
    } catch {
        netsh interface set interface $adapter.Name disabled 2>$null
    }
    Start-Sleep 2

    Write-Host "  [3/3] 启用适配器..." -ForegroundColor Yellow
    try {
        Enable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
    } catch {
        netsh interface set interface $adapter.Name enabled 2>$null
    }
    Start-Sleep 3

    # 验证
    Write-Host ""
    try {
        $verify = (Get-NetAdapter -Name $adapter.Name -ErrorAction Stop).MacAddress -replace '-', ''
        Write-Host "  ✔ 已恢复！当前 MAC: $(Format-MAC $verify)" -ForegroundColor Green
    } catch {
        Write-Host "  ✔ 注册表已清除，请检查网络连接" -ForegroundColor Yellow
    }

    # 写日志
    $logFile = Join-Path $PSScriptRoot "mac_change_log.txt"
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logFile -Value "$ts  恢复出厂  适配器=$($adapter.Name)  旧MAC=$(Format-MAC $oldMAC)" -Encoding UTF8
}

# ============================================
#  选择适配器
# ============================================
function Select-Adapter($adapters) {
    if ($adapters.Count -eq 1) {
        return $adapters[0]
    }
    $sel = Read-Host "  选择适配器序号 (1-$($adapters.Count), 默认1)"
    if ($sel -eq '') { $sel = '1' }
    $idx = [int]$sel - 1
    if ($idx -lt 0 -or $idx -ge $adapters.Count) {
        Write-Host "  无效选择！" -ForegroundColor Red
        return $null
    }
    return $adapters[$idx]
}

# ============================================
#  主程序
# ============================================
Clear-Host
Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "         Windows MAC 地址修改器              " -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan

# 加载适配器
Write-Host ""
Write-Host "  正在扫描网络适配器..." -ForegroundColor Yellow
$adapters = @(Get-PhysicalAdapters)

if ($adapters.Count -eq 0) {
    Write-Host ""
    Write-Host "  未找到物理网络适配器！" -ForegroundColor Red
    Read-Host "  按回车退出"
    exit
}

# 主循环
while ($true) {
    Show-Adapters $adapters

    Write-Host ""
    Write-Host "  操作选项:" -ForegroundColor Yellow
    Write-Host "  [回车] 随机更换MAC地址（最常用）" -ForegroundColor White
    Write-Host "  [2]   手动输入MAC地址" -ForegroundColor White
    Write-Host "  [3]   恢复出厂MAC地址" -ForegroundColor White
    Write-Host "  [4]   刷新列表" -ForegroundColor White
    Write-Host "  [0]   退出" -ForegroundColor White
    Write-Host ""
    $choice = Read-Host "  请选择"
    if ($choice -eq '') { $choice = '1' }

    switch ($choice) {
        '0' { exit }
        '1' {
            $adapter = Select-Adapter $adapters
            if ($adapter) {
                $newMAC = New-RandomMAC
                Set-NewMAC $adapter $newMAC | Out-Null
                Write-Host ""
                # 刷新列表
                $adapters = @(Get-PhysicalAdapters)
            }
        }
        '2' {
            $adapter = Select-Adapter $adapters
            if ($adapter) {
                $input = Read-Host "  请输入新MAC地址 (如 AA-BB-CC-DD-EE-FF 或 AABBCCDDEEFF)"
                $newMAC = ($input -replace '[^0-9a-fA-F]', '').ToUpper()
                if ($newMAC.Length -ne 12) {
                    Write-Host "  MAC地址格式错误，需要12位十六进制！" -ForegroundColor Red
                } else {
                    Set-NewMAC $adapter $newMAC | Out-Null
                    $adapters = @(Get-PhysicalAdapters)
                }
            }
        }
        '3' {
            $adapter = Select-Adapter $adapters
            if ($adapter) {
                Restore-OriginalMAC $adapter
                $adapters = @(Get-PhysicalAdapters)
            }
        }
        '4' {
            Write-Host "  正在刷新..." -ForegroundColor Yellow
            $adapters = @(Get-PhysicalAdapters)
        }
        default {
            Write-Host "  无效选择！" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "  -------------------------------------------" -ForegroundColor DarkGray
}
