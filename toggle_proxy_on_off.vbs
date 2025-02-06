' This file is part of Windows_Proxy_Toggler: https://github.com/ElectricRCAircraftGuy/Windows_Proxy_Toggler
'
' Toggle your Proxy on and off via a clickable desktop shortcut/icon
' By Gabriel Staples, June 2017
' www.ElectricRCAircraftGuy.com
' See the README at the link above.

Option Explicit

' Variables & Constants:
Dim ProxySettings_path, VbsScript_filename
VbsScript_filename = "toggle_proxy_on_off.vbs"
Const MESSAGE_BOX_TIMEOUT = 1
Const PROXY_OFF = 0

Dim WSHShell, proxyEnableVal, username
Set WSHShell = WScript.CreateObject("WScript.Shell")
username = WSHShell.ExpandEnvironmentStrings("%USERNAME%")
ProxySettings_path = "C:\Program Files\Windows_Proxy_Toggler"

' Determine current proxy setting and toggle to opposite setting
proxyEnableVal = WSHShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")
If proxyEnableVal = PROXY_OFF Then
    TurnProxyOn
Else
    TurnProxyOff
End If

' Subroutine to Toggle Proxy Setting to ON
Sub TurnProxyOn
    WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
    CreateOrUpdateStartMenuShortcut("on")
    WSHShell.Popup "Internet proxy is now ON", MESSAGE_BOX_TIMEOUT, "Proxy Settings"
End Sub

' Subroutine to Toggle Proxy Setting to OFF
Sub TurnProxyOff
    WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
    CreateOrUpdateStartMenuShortcut("off")
    WSHShell.Popup "Internet proxy is now OFF", MESSAGE_BOX_TIMEOUT, "Proxy Settings"
End Sub

' ' Subroutine to create or update a shortcut on the desktop
' Sub CreateOrUpdateDesktopShortcut(onOrOff)
'     Dim shortcut, iconStr
'     Dim desktopPath
'     desktopPath = WSHShell.SpecialFolders("Desktop")
'     Set shortcut = WSHShell.CreateShortcut(desktopPath & "\Proxy On-Off.lnk")
'     shortcut.TargetPath = ProxySettings_path & "\" & VbsScript_filename
'     shortcut.WorkingDirectory = ProxySettings_path
'     If onOrOff = "on" Then
'         iconStr = "on.ico"
'     ElseIf onOrOff = "off" Then
'         iconStr = "off.ico"
'     End If
'     shortcut.IconLocation = ProxySettings_path & "\icons\" & iconStr
'     shortcut.Save
' End Sub

' Subroutine to create or update a shortcut in the start menu
Sub CreateOrUpdateStartMenuShortcut(onOrOff)
    Dim shortcut, iconStr
    Dim startMenuPath
    startMenuPath = WSHShell.SpecialFolders("StartMenu") & "\Programs"
    Set shortcut = WSHShell.CreateShortcut(startMenuPath & "\Aproxy.lnk")
    shortcut.TargetPath = ProxySettings_path & "\" & VbsScript_filename
    shortcut.WorkingDirectory = ProxySettings_path
    If onOrOff = "on" Then
        iconStr = "on.ico"
    ElseIf onOrOff = "off" Then
        iconStr = "off.ico"
    End If
    shortcut.IconLocation = ProxySettings_path & "\icons\" & iconStr
    shortcut.Save
End Sub
