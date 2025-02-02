' This file is part of Windows_Proxy_Toggler: https://github.com/ElectricRCAircraftGuy/Windows_Proxy_Toggler
'
' Toggle your Proxy on and off via a clickable desktop shortcut/icon
' By Gabriel Staples, June 2017
' www.ElectricRCAircraftGuy.com
' See the README at the link above.

' 强制变量声明，避免因未声明变量而产生的错误
Option Explicit

' Variables & Constants:
' 定义存储代理设置程序的路径变量
Dim ProxySettings_path, VbsScript_filename
' 定义当前脚本的文件名
VbsScript_filename = "toggle_proxy_on_off.vbs"
' 定义消息框显示的超时时间（秒），可修改此值来调整消息框显示时长
Const MESSAGE_BOX_TIMEOUT = 1
' 定义表示代理关闭状态的常量
Const PROXY_OFF = 0

' 声明变量
Dim WSHShell, proxyEnableVal, username
' 创建 WScript.Shell 对象，用于执行系统操作，如读写注册表、创建快捷方式等
Set WSHShell = WScript.CreateObject("WScript.Shell")
' 获取当前用户名，避免直接在路径中使用 "%USERNAME%" 变量导致错误
username = WSHShell.ExpandEnvironmentStrings("%USERNAME%")
' 修改存储代理设置程序的路径
ProxySettings_path = "C:\Program Files\Windows_Proxy_Toggler"

' Determine current proxy setting and toggle to opposite setting
' 从注册表中读取当前代理启用状态
proxyEnableVal = WSHShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")
' 根据当前代理状态决定调用开启或关闭代理的子例程
If proxyEnableVal = PROXY_OFF Then
    TurnProxyOn
Else
    TurnProxyOff
End If

' Subroutine to Toggle Proxy Setting to ON
' 开启代理的子例程
Sub TurnProxyOn
    ' 通过注册表项将代理设置为开启状态
    WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
    ' 创建或更新开始菜单快捷方式，传入 "on" 表示代理开启状态
    CreateOrUpdateStartMenuShortcut("on")
    ' 通过自动计时的弹出框通知用户代理已开启
    WSHShell.Popup "Internet proxy is now ON", MESSAGE_BOX_TIMEOUT, "Proxy Settings"
End Sub

' Subroutine to Toggle Proxy Setting to OFF
' 关闭代理的子例程
Sub TurnProxyOff
    ' 通过注册表项将代理设置为关闭状态
    WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
    ' 创建或更新开始菜单快捷方式，传入 "off" 表示代理关闭状态
    CreateOrUpdateStartMenuShortcut("off")
    ' 通过自动计时的弹出框通知用户代理已关闭
    WSHShell.Popup "Internet proxy is now OFF", MESSAGE_BOX_TIMEOUT, "Proxy Settings"
End Sub

' ' Subroutine to create or update a shortcut on the desktop
' ' 创建或更新桌面快捷方式的子例程
' Sub CreateOrUpdateDesktopShortcut(onOrOff)
'     ' 声明变量，用于存储快捷方式对象和图标文件名
'     Dim shortcut, iconStr
'     ' 声明变量，用于存储桌面路径
'     Dim desktopPath
'     ' 使用 SpecialFolders 方法获取桌面路径
'     desktopPath = WSHShell.SpecialFolders("Desktop")
'     ' 创建桌面快捷方式对象
'     Set shortcut = WSHShell.CreateShortcut(desktopPath & "\Proxy On-Off.lnk")
'     ' 设置快捷方式点击时要运行的目标文件路径
'     shortcut.TargetPath = ProxySettings_path & "\" & VbsScript_filename
'     ' 设置快捷方式的工作目录，确保在调用相关脚本时能正确找到文件
'     shortcut.WorkingDirectory = ProxySettings_path
'     ' 根据代理状态选择对应的图标文件名
'     If onOrOff = "on" Then
'         iconStr = "on.ico"
'     ElseIf onOrOff = "off" Then
'         iconStr = "off.ico"
'     End If
'     ' 设置快捷方式关联的图标路径
'     shortcut.IconLocation = ProxySettings_path & "\icons\" & iconStr
'     ' 保存快捷方式的设置
'     shortcut.Save
' End Sub

' Subroutine to create or update a shortcut in the start menu
' 创建或更新开始菜单快捷方式的子例程
Sub CreateOrUpdateStartMenuShortcut(onOrOff)
    ' 声明变量，用于存储快捷方式对象和图标文件名
    Dim shortcut, iconStr
    ' 声明变量，用于存储开始菜单程序文件夹的路径
    Dim startMenuPath
    ' 使用 SpecialFolders 方法获取开始菜单路径，并拼接上 "Programs" 文件夹
    startMenuPath = WSHShell.SpecialFolders("StartMenu") & "\Programs"
    ' 创建开始菜单快捷方式对象，将名字改为 Aproxy
    Set shortcut = WSHShell.CreateShortcut(startMenuPath & "\Aproxy.lnk")
    ' 设置快捷方式点击时要运行的目标文件路径
    shortcut.TargetPath = ProxySettings_path & "\" & VbsScript_filename
    ' 设置快捷方式的工作目录，确保在调用相关脚本时能正确找到文件
    shortcut.WorkingDirectory = ProxySettings_path
    ' 根据代理状态选择对应的图标文件名
    If onOrOff = "on" Then
        iconStr = "on.ico"
    ElseIf onOrOff = "off" Then
        iconStr = "off.ico"
    End If
    ' 设置快捷方式关联的图标路径
    shortcut.IconLocation = ProxySettings_path & "\icons\" & iconStr
    ' 保存快捷方式的设置
    shortcut.Save
End Sub