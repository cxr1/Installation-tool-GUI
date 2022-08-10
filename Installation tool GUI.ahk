#SingleInstance Force
#NoEnv
SetRegView 64
SetBatchLines -1
SendMode Input
#Persistent

;强制以ANSI版本管理员权限运行
runwith("admin","A")

if (A_IsUnicode=1 and A_PtrSize=8) ; 用于修正32位AHK读取不到"C:\Windows\System32"的问题
    DllCall("Wow64DisableWow64FsRedirection")

; 加载皮肤，只支持32位AHK
hSkinH := DllCall("LoadLibrary", "Str", "SkinH.dll")
DllCall("SkinH\SkinH_AttachEx", "Str", A_ScriptDir "\自改黑红皮肤修正.she")

;查看电脑型号
command_pcname = wmic csproduct get name
pcname := cmdSilenceReturn(command_pcname)
pcname := StrReplace(pcname, A_Space, "")
pcname := StrReplace(pcname, "`r`n")
; pcname := StrReplace(pcname, "Name", "型号：     ")
pcname := SubStr(pcname, 5, StrLen(pcname))

;读取bios sn
command_sn = wmic bios get serialnumber
bios_sn := cmdSilenceReturn(command_sn)
bios_sn := StrReplace(bios_sn, A_Space, "")
bios_sn := StrReplace(bios_sn, "`r`n")
; bios_sn := StrReplace(bios_sn, "SerialNumber", "序列号：   ")
bios_sn := SubStr(bios_sn, 13, StrLen(bios_sn))

;读取主板sn
; command = wmic baseboard get serialnumber
; board_sn := cmdSilenceReturn(command)

;查看工作组
command_domain = wmic computersystem get domain
get_domain := cmdSilenceReturn(command_domain)
get_domain := StrReplace(get_domain, A_Space, "")
get_domain := StrReplace(get_domain, "`r`n")
;get_domain := StrReplace(get_domain, "Domain", "工作组/域：")
;get_domain := SubStr(get_domain, 12, StrLen(get_domain))
get_domain := SubStr(get_domain, 7, StrLen(get_domain))

;读取钉钉版本
RegRead, dd_version, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\钉钉, DisplayVersion

;静音
;Send {Volume_Mute}

;打开word
;Run, winword
;打开激活界面
;Run, ms-settings:activation,Max

; Gui Font, s10
; Gui Add, Text, x5 y200 w230 h30, %pcname%
; ;Gui Add, Text, x5 y200 w230 h30, 机器SN:
; Gui Add, Text, x5 y230 w230 h30, %bios_sn%
; ;Gui Add, Text, x5 y250 w230 h30, 工作组/域:
; Gui Add, Text, x5 y260 w230 h30, %get_domain%

Gui Font, s8
; Gui Add, Edit, x70 y218 vSN ReadOnly
; Gui Add, Edit, x70 y248 vdomain ReadOnly
GuiControl,, SN, %bios_sn%
GuiControl,, domain, %get_domain%
; Gui Add, Text, x5 y250 w230 h30, 主板_%board_sn%

Gui Font, s10
; Gui Add, Text, x5 y2 w230 h30, -----------用户管理-----------
; Gui Add, Text, x5 y130 w230 h30, -----------安装软件-----------
Gui Add, Text, x5 y20 w60 h20, 用户名:
Gui Add, Text, x5 y45 w60 h20, 目标组:
Gui Add, Edit, x60 y18 w150 h20 vUsername
Gui Add, DDL, x60 y43 vcbx w150 hwndhcbx, Administrators||Users
Gui Add, Button, x75 y75 w65 h20, 添加用户
Gui Add, Button, x145 y75 w65 h20, 查看用户
;Gui Add, Button, x5 y100 w65 h20, 加入W组
;Gui Add, Button, x75 y100 w65 h20, 加入域
Gui Add, Button, x5 y75 w65 h20, 用户和组
Gui Add, Button, x5 y145 w65 h20, 安装钉钉

Gui, Add, GroupBox, x3 y2 w210 h125 +Center, % " 用户管理 "
Gui, Add, GroupBox, x3 y130 w210 h42 +Center, % " 安装软件 "
; Gui, Add, GroupBox, x3 y185 w210 h150, 计算机信息

Gui Font, s9
Gui, Add, ListView, r20 w210 h150 gMyListView Grid NoSortHdr -Hdr NoSort -Multi, 项目|详细信息
LV_ModifyCol(1, 62)
LV_ModifyCol(2, 144)
LV_Add("", "型号", pcname)
LV_Add("", "序列号", bios_sn)
LV_Add("", "工作组/域", get_domain)
LV_Add("", "钉钉版本", dd_version)

Gui, Add, StatusBar


Gui Show, x1000 y150 w216 h355, 装机工具by CXR
Return

MyListView:
    if (A_GuiEvent = "DoubleClick")
    {
        LV_GetText(RowText, A_EventInfo, 2) ; 从行的第一个字段中获取文本.
        ToolTip %RowText%
        Sleep, 3000
        ToolTip
    }
return

Button用户和组:
    Run, lusrmgr.msc
return

Button添加用户:
    GuiControlGet, user_name ,, Username
    GuiControlGet, group ,, cbx
    command = net user %user_name% /add ;添加用户
    command1 = net localgroup %group% %user_name% /add ;添加到选择的组
    command2 = wmic.exe UserAccount Where Name="%user_name%" Set PasswordExpires="false" ;密码永不过期
    user := cmdSilenceReturn(command)
    cmdSilenceReturn(command2)
    user_add := cmdSilenceReturn(command1)
    If (user = "")
        msgbox 添加失败
    Else
        msgbox,,%user_name%, % user_add
Return

Button查看用户:
    GuiControlGet, group,, cbx
    command = net localgroup %group%
    user := cmdSilenceReturn(command)
    msgbox,,%group%, % user
return

;Button加入W组:
;Return

;Button加入域:
;Return

Button安装钉钉:
    SB_SetText("正在安装钉钉！")
    runwait 钉钉.exe /S     ;需要离线安装包
    SB_SetText("钉钉安装成功！")
    RegRead, dd_version, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\钉钉, DisplayVersion
    LV_Add("", "钉钉_新", dd_version)
return

GuiEscape:
GuiClose:
ExitApp

;强制改权限
RunWith(RunAsAdmin:="Default", ANSI_U32_U64:="Default")
{
    ; 格式化预期的模式
    switch, RunAsAdmin
    {
        case "Normal","Standard","No","0":		RunAsAdmin:=0
        case "Admin","Yes","1":								RunAsAdmin:=1
        case "Default":												RunAsAdmin:=A_IsAdmin
        default:															RunAsAdmin:=A_IsAdmin
    }
    switch, ANSI_U32_U64
    {
        case "A32","ANSI","A":								ANSI_U32_U64:="AutoHotkeyA32.exe"
        case "U32","X32","32":								ANSI_U32_U64:="AutoHotkeyU32.exe"
        case "U64","X64","64":								ANSI_U32_U64:="AutoHotkeyU64.exe"
        case "Default":												ANSI_U32_U64:="AutoHotkey.exe"
        default:															ANSI_U32_U64:="AutoHotkey.exe"
    }
    ; 获取传递给 “.ahk” 的用户参数（不是 /restart 之类传递给 “.exe” 的开关参数）
    for k, v in A_Args
    {
        if (RunAsAdmin=1)
        {
            ; 转义所有的引号与转义符号
            v:=StrReplace(v, "\", "\\")
            v:=StrReplace(v, """", "\""")
            ; 无论参数中是否有空格，都给参数两边加上引号
            ; Run       的内引号是 "
            ScriptParameters .= (ScriptParameters="") ? """" v """" : A_Space """" v """"
        }
        else
        {
            ; 转义所有的引号与转义符号
            ; 注意要转义两次 Run 和 RunAs.exe
            v:=StrReplace(v, "\", "\\")
            v:=StrReplace(v, """", "\""")
            v:=StrReplace(v, "\", "\\")
            v:=StrReplace(v, """", "\""")
            ; 无论参数中是否有空格，都给参数两边加上引号
            ; RunAs.exe 的内引号是 \"
            ScriptParameters .= (ScriptParameters="") ? "\""" v "\""" : A_Space "\""" v "\"""
        }
    }
    ; 判断当前 exe 是什么版本
    if (!A_IsUnicode)
        RunningEXE:="AutoHotkeyA32.exe"
    else if (A_PtrSize=4)
        RunningEXE:="AutoHotkeyU32.exe"
    else if (A_PtrSize=8)
        RunningEXE:="AutoHotkeyU64.exe"
    ; 运行模式与预期相同，则直接返回。 ANSI_U32_U64="AutoHotkey.exe" 代表不对 ahk 版本做要求。
    if (A_IsAdmin=RunAsAdmin and (ANSI_U32_U64="AutoHotkey.exe" or ANSI_U32_U64=RunningEXE))
        return
    ; 如果当前已经是使用 /restart 参数重启的进程，则报错避免反复重启导致死循环。
    else if (RegExMatch(DllCall("GetCommandLine", "str"), " /restart(?!\S)"))
    {
        预期权限:=(RunAsAdmin=1) ? "管理员权限" : "普通权限"
        当前权限:=(A_IsAdmin=1) ? "管理员权限" : "普通权限"
            ErrorMessage=
            (LTrim
            预期使用: %ANSI_U32_U64%
            当前使用: %RunningEXE%
            预期权限: %预期权限%
            当前权限: %当前权限%
            程序即将退出。
            )
            MsgBox 0x40030, 运行状态与预期不一致, %ErrorMessage%
        ExitApp
    }
    else
    {
        ; 获取 AutoHotkey.exe 的路径
        SplitPath, A_AhkPath, , Dir
        if (RunAsAdmin=0)
        {
            ; 强制普通权限运行
            switch, A_IsCompiled
            {
                ; %A_ScriptFullPath% 必须加引号，否则含空格的路径会被截断。%ScriptParameters% 必须不加引号，因为构造时已经加了。
                ; 工作目录不用单独指定，默认使用 A_WorkingDir 。
                case, "1": Run, RunAs.exe /trustlevel:0x20000 "\"%A_ScriptFullPath%\" /restart %ScriptParameters%",, Hide
                default: Run, RunAs.exe /trustlevel:0x20000 "\"%Dir%\%ANSI_U32_U64%\" /restart \"%A_ScriptFullPath%\" %ScriptParameters%",, Hide
            }
        }
        else
        {
            ; 强制管理员权限运行
            switch, A_IsCompiled
            {
                ; %A_ScriptFullPath% 必须加引号，否则含空格的路径会被截断。%ScriptParameters% 必须不加引号，因为构造时已经加了。
                ; 工作目录不用单独指定，默认使用 A_WorkingDir 。
                case, "1": Run, *RunAs "%A_ScriptFullPath%" /restart %ScriptParameters%
                default: Run, *RunAs "%Dir%\%ANSI_U32_U64%" /restart "%A_ScriptFullPath%" %ScriptParameters%
            }
        }
        ExitApp
    }
}

;后台静默运行cmd命令缓存文本取值
cmdSilenceReturn(command){
    CMDReturn:=""
    cmdFN:="RunAnyCtrlCMD"
    try{
        RunWait,% ComSpec " /C " command " > ""%Temp%\" cmdFN ".log""",, Hide
        FileRead, CMDReturn, %A_Temp%\%cmdFN%.log
        FileDelete,%A_Temp%\%cmdFN%.log
    }catch{}
    return CMDReturn
}
