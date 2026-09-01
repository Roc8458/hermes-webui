# 🇨🇳 中文远程访问指南

> **原作者**: [nesquena/hermes-webui](https://github.com/nesquena/hermes-webui) (14.8K⭐)
>
> **本 Fork 维护者**: [Roc8458](https://github.com/Roc8458)
>
> **适用场景**: Windows 电脑 + iPhone/Android 手机，通过 Tailscale 远程访问 Hermes WebUI

---

## 📋 目录

- [原版功能](#原版功能)
- [我的实用改动](#我的实用改动)
- [快速开始](#快速开始)
- [开机自启配置](#开机自启配置)
- [防火墙配置](#防火墙配置)
- [常见问题](#常见问题)

---

## 原版功能

本 Fork **未修改任何源码**，所有功能与原版完全一致。原版特性包括：

- 支持 iPhone PWA（添加到主屏幕当 App 用）
- 三栏布局：左侧会话列表 + 中间聊天 + 右侧工作区文件浏览
- 零构建步骤，纯 Python + vanilla JS，启动即用
- 1:1 复刻 CLI 的所有功能
- 支持多语言、多主题

**原版仓库**: https://github.com/nesquena/hermes-webui

---

## 我的实用改动

以下是将 WebUI 从「本机工具」改造成「手机也能用的远程服务」的配置优化：

### 1. 绑定地址（关键）

**原版默认**: `127.0.0.1`（仅本机可访问）

**改为**: `0.0.0.0`（所有网卡可访问）

```bash
# 启动时指定
HERMES_WEBUI_HOST=0.0.0.0 python server.py
```

**效果**: 手机通过 Tailscale 网络能访问

### 2. 开机自启（Windows）

**原版**: 无自启，需手动启动

**新增**: Watchdog 脚本，开机自动运行，掉线自动拉起

文件结构：
```
C:\Users\<你的用户名>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\
└── Hermes_WebUI.vbs          # 开机自启入口

C:\Users\<你的用户名>\AppData\Local\hermes\webui-service\
└── Hermes_WebUI.vbs          # Watchdog（每60秒检查）
```

### 3. 防火墙规则

**原版**: 无防火墙配置

**新增**: 放行 8787 端口

```powershell
# 需要管理员权限运行
netsh advfirewall firewall add rule name="Hermes WebUI" dir=in action=allow protocol=TCP localport=8787
```

### 4. Tailscale 组网

**原版**: 无远程访问方案

**新增**: 通过 Tailscale 实现电脑 ↔ 手机互通

1. 电脑和手机都安装 Tailscale
2. 同一账号登录
3. 手机浏览器访问 `http://电脑Tailscale IP:8787`

### 5. 界面优化（用户自选）

| 设置项 | 推荐值 | 说明 |
|--------|--------|------|
| 语言 | `zh` | 中文界面 |
| 主题 | `light` 或 `dark` | 个人偏好 |
| 默认模型 | `mimo-v2.5` | 小米模型（或你喜欢的） |
| 显示 Token 用量 | 关闭 | 界面更干净 |
| 显示 TPS | 关闭 | 减少干扰 |

---

## 快速开始

### 1. 克隆并启动

```bash
git clone https://github.com/Roc8458/hermes-webui.git
cd hermes-webui

# 启动（指定绑定地址）
HERMES_WEBUI_HOST=0.0.0.0 HERMES_WEBUI_PORT=8787 python server.py
```

### 2. 手机访问

1. 确保电脑和手机都安装了 [Tailscale](https://tailscale.com/)
2. 同一账号登录
3. 手机浏览器访问：`http://电脑的Tailscale IP:8787`
4. （可选）Safari 添加到主屏幕 → PWA App

### 3. 设置中文

1. 打开 WebUI
2. 点击左下角 ⚙️ Settings
3. Language → 选择 `zh`（中文）
4. 刷新页面生效

---

## 开机自启配置

### 文件 1: Startup 入口

**路径**: `C:\Users\<用户名>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Hermes_WebUI.vbs`

```vbscript
' Hermes WebUI - 开机自启入口
Option Explicit
Dim fso, sh, target
target = "C:\Users\<你的用户名>\AppData\Local\hermes\webui-service\Hermes_WebUI.vbs"
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(target) Then WScript.Quit 0
Set sh = CreateObject("WScript.Shell")
sh.Run "wscript.exe C:\Users\<你的用户名>\AppData\Local\hermes\webui-service\Hermes_WebUI.vbs", 0, False
```

### 文件 2: Watchdog

**路径**: `C:\Users\<你的用户名>\AppData\Local\hermes\webui-service\Hermes_WebUI.vbs`

```vbscript
' Hermes WebUI Watchdog
' 每 60 秒检查，掉线自动拉起
Option Explicit
Dim sh, wmi, procs, WEBUI_CMD
Set sh = CreateObject("WScript.Shell")
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

' 单实例保护
Set procs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='wscript.exe' AND CommandLine LIKE '%Hermes_WebUI.vbs%'")
If procs.Count > 1 Then WScript.Quit 0

WEBUI_CMD = "cmd /c cd /d D:\hermes-webui && set HERMES_WEBUI_HOST=0.0.0.0 && set HERMES_WEBUI_PORT=8787 && C:\Users\<你的用户名>\AppData\Local\hermes\hermes-agent\venv\Scripts\python.exe server.py"

Function IsWebUIRunning()
  Dim p
  IsWebUIRunning = False
  On Error Resume Next
  Set p = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='python.exe' AND CommandLine LIKE '%server.py%'")
  On Error GoTo 0
  If Not p Is Nothing Then
    If p.Count > 0 Then IsWebUIRunning = True
  End If
End Function

' 启动时立即检查一次
If Not IsWebUIRunning() Then
  sh.Run WEBUI_CMD, 0, False
End If

' 主循环：每 60 秒检查
Do While True
  WScript.Sleep 60000
  If Not IsWebUIRunning() Then
    sh.Run WEBUI_CMD, 0, False
  End If
Loop
```

**注意**: 请将 `<你的用户名>` 替换为你的 Windows 用户名，`D:\hermes-webui` 替换为你的实际安装路径。

---

## 防火墙配置

### 方法 1: 命令行（推荐）

```powershell
# 需要管理员权限
netsh advfirewall firewall add rule name="Hermes WebUI" dir=in action=allow protocol=TCP localport=8787
```

### 方法 2: 图形界面

1. Windows 搜索 → 「高级安全 Windows Defender 防火墙」
2. 左侧 → 「入站规则」→ 右侧 → 「新建规则」
3. 端口 → TCP → 特定端口：`8787`
4. 允许连接 → 完成
5. 名称：`Hermes WebUI`

---

## 常见问题

### Q: 手机访问显示「拒绝连接」

**原因**: 防火墙未放行 8787 端口

**解决**: 运行上面的防火墙配置命令

### Q: 手机访问显示「无法连接到服务器」

**原因**: WebUI 未启动

**解决**:
```bash
# 手动启动
cd D:\hermes-webui
HERMES_WEBUI_HOST=0.0.0.0 python server.py
```

### Q: 电脑重启后手机还是连不上

**原因**: 开机自启未配置

**解决**: 按照「开机自启配置」章节设置

### Q: 手机和电脑同时聊天会冲突吗？

**会**。同一时间只允许一个 agent 运行在一个会话上。

**解决方案**:
- 手机聊完再切电脑
- 或者手机和电脑各自开独立会话

### Q: 能用微信/飞书访问吗？

**不能**。微信/飞书走的是 Gateway（另一个组件），和 WebUI 是独立的。

本方案只适用于 **Tailscale + 浏览器** 的方式。

---

## 致谢

- **原版作者**: [nesquena](https://github.com/nesquena) - 创建了 hermes-webui (14.8K⭐)
- **社区贡献者**: 所有参与原版开发的贡献者
- **本 Fork 维护**: [Roc8458](https://github.com/Roc8458) - 添加中文远程访问指南

---

## 📄 License

本 Fork 遵循原版 License。请查看原版仓库了解详情。
