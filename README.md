# Windows MAC 地址管理工具

基于 Python + tkinter 的 Windows MAC 地址管理工具，支持手动修改、随机生成、定时自动更换和恢复出厂 MAC 地址。

## 功能

- 列出所有物理网络适配器及 MAC 地址，标注活跃状态
- 手动输入或随机生成合法 MAC 地址并应用
- 通过注册表实现永久修改（重启不恢复）
- 定时自动随机更换 MAC（0.5 / 1 / 2 / 6 / 12 / 24 小时间隔）
- 系统托盘常驻，关闭窗口最小化到托盘
- 一键恢复网卡原始硬件 MAC
- 操作日志记录

## 安装

```bash
pip install -r requirements.txt
```

## 运行

**必须以管理员权限运行**，程序启动时会自动检测并请求 UAC 提权。

```bash
python mac_manager.py
```

## 打包为 EXE

```bash
pyinstaller --onefile --windowed --uac-admin --icon=NONE --name="MAC地址管理器" mac_manager.py
```

参数说明：
- `--onefile`：打包为单个 .exe 文件
- `--windowed`：不显示控制台窗口
- `--uac-admin`：自动请求管理员权限（嵌入 manifest）
- 生成的 exe 位于 `dist/` 目录

## 技术原理

### 注册表修改方式

MAC 地址修改通过 Windows 注册表实现：

```
HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\00xx
```

- 每个 `00xx` 子键对应一个网络适配器
- 写入 `NetworkAddress`（REG_SZ）值为 12 位大写十六进制 MAC 字符串（无分隔符）
- 写入后需禁用再启用适配器才能生效（通过 `netsh interface set interface` 实现）
- 删除 `NetworkAddress` 键即恢复出厂 MAC

### 原始 MAC 获取

通过 WMI 查询 `Win32_NetworkAdapter` 的 `MACAddress` 属性获取硬件原始 MAC。

### 随机 MAC 生成规则

第一字节遵循 IEEE 802 规范：
- **bit 0 = 0**：单播地址（非多播）
- **bit 1 = 1**：本地管理地址（LAA，非全局唯一）

即第一字节低 4 位为 `x2`、`x6`、`xA`、`xE`。

### 生效机制

修改注册表后，通过 `netsh` 禁用并重新启用适配器：

```
netsh interface set interface "以太网" disabled
netsh interface set interface "以太网" enabled
```

这会导致短暂断网（约 3-5 秒）。

## 注意事项

1. **管理员权限**：修改注册表和操作网络适配器必须有管理员权限
2. **驱动兼容性**：部分老旧网卡驱动不支持 `NetworkAddress` 注册表键，此类网卡无法通过本工具修改 MAC
3. **虚拟适配器**：工具已过滤 VPN、虚拟网卡等非物理适配器
4. **断网风险**：每次应用修改会短暂断网，定时自动更换时请注意
5. **企业网络**：某些企业网络使用 MAC 白名单进行准入控制，随意修改 MAC 可能导致无法联网
6. **Windows 版本**：支持 Windows 10 / 11，更早版本未测试
7. **杀毒软件**：部分杀毒软件可能对注册表修改行为报警，属正常现象

## 文件说明

| 文件 | 说明 |
|------|------|
| `mac_manager.py` | 主程序源码 |
| `requirements.txt` | Python 依赖 |
| `mac_change_log.txt` | 运行时自动生成的操作日志 |
| `mac_manager_config.json` | 运行时自动生成的配置文件（保存自动更换状态） |

## 依赖

- Python 3.10+
- pystray（系统托盘图标）
- Pillow（托盘图标绘制）
- 内置模块：tkinter、winreg、subprocess、ctypes
