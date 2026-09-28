#Requires AutoHotkey v2.0
#SingleInstance Force

SetTitleMatchMode 2

; ================== WPS PPT COM 翻页控制 ==================
; Adding PgUp/PgDn allows you to turn pages of PPT regardless of whether it is a front-end application,by clinging
; OBS Websocket option needs to be opened

[::
$PgUp::
{
    WPP_SlideAction("Previous")
}

; ] 键：下一页（Page Down）
]::
$PgDn::
{
    WPP_SlideAction("Next")
}

WPP_SlideAction(action)
{
    ppt := ""
    ; 1. 尝试获取正在运行的演示实例（兼容 WPS 与 MS Office）
    try {
        ; 优先检测 WPS 放映窗口是否存在
        if WinExist("WPS Presentation Slide Show - ahk_exe wpp.exe") {
            try ppt := ComObjActive("Kwpp.Application")
        }
        ; 若 WPS 放映窗口不存在或绑定失败，再检测 PowerPoint 放映窗口
        if !IsObject(ppt) && WinExist("PowerPoint - ahk_exe POWERPNT.EXE") {
            try ppt := ComObjActive("PowerPoint.Application")
        }
        ; 兜底兼容：若上述特定窗口未匹配（如窗口模式放映），按活跃 COM 实例降级探测
        if !IsObject(ppt) {
            try {
                wppApp := ComObjActive("Kwpp.Application")
                if (wppApp.SlideShowWindows.Count > 0) {
                    ppt := wppApp
                }
            } catch {
                ; 忽略异常继续尝试 PowerPoint
            }
            if !IsObject(ppt) {
                try {
                    msApp := ComObjActive("PowerPoint.Application")
                    if (msApp.SlideShowWindows.Count > 0)
                        ppt := msApp
                } catch {
                    ; 忽略异常
                }
            }
        }
    } catch {
        ppt := ""
    }

    ; 2. 检查是否有正在放映的窗口，并执行动作
    if IsObject(ppt) {
        try {
            if (ppt.SlideShowWindows.Count > 0) {
                if (action == "Next")
                    ppt.SlideShowWindows(1).View.Next()
                else
                    ppt.SlideShowWindows(1).View.Previous()
                return
            }
        }
    }

    ; 3. 若未检测到放映或 COM 未就绪，弹出轻量提示
    ToolTip "未检测到正在放映的幻灯片"
    SetTimer () => ToolTip(), -1500
}


RShift::
{
    curScene := RunStdout('obs-cmd scene current')

    if InStr(curScene, "嘉宾PPT")
        SwitchOBSScene("流程PPT")
    else if InStr(curScene, "流程PPT")
        SwitchOBSScene("嘉宾PPT")
    else
        SwitchOBSScene("流程PPT")
}

; 按顿号键（\ 键）切换到「主KV」
\::
{
    SwitchOBSScene("主KV")
}

SwitchOBSScene(sceneName)
{
    cmd := Format('obs-cmd scene switch "{1}"', sceneName)
    Run(cmd, , "Hide")
}

; 通过系统内存管道静默抓取命令行输出（零临时文件、无黑框闪烁）
RunStdout(cmd)
{
    DllCall("CreatePipe", "Ptr*", &hRead := 0, "Ptr*", &hWrite := 0, "Ptr", 0, "UInt", 0)
    DllCall("SetHandleInformation", "Ptr", hWrite, "UInt", 1, "UInt", 1)

    is64 := (A_PtrSize == 8)
    siSize := is64 ? 104 : 68
    piSize := is64 ? 24 : 16

    si := Buffer(siSize, 0)
    pi := Buffer(piSize, 0)

    NumPut("UInt", siSize, si, 0)
    NumPut("UInt", 0x101, si, is64 ? 60 : 44)  ; STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW
    NumPut("UShort", 0, si, is64 ? 64 : 48)    ; SW_HIDE（隐藏窗口）
    NumPut("Ptr", hWrite, si, is64 ? 88 : 60)  ; 标准输出重定向到管道
    NumPut("Ptr", hWrite, si, is64 ? 96 : 64)  ; 标准错误重定向到管道

    ; CREATE_NO_WINDOW = 0x08000000
    if DllCall("CreateProcessW", "Ptr", 0, "Str", cmd, "Ptr", 0, "Ptr", 0, "Int", 1, "UInt", 0x08000000, "Ptr", 0, "Ptr", 0, "Ptr", si.Ptr, "Ptr", pi.Ptr) {
        DllCall("CloseHandle", "Ptr", hWrite)
        DllCall("CloseHandle", "Ptr", NumGet(pi, 0, "Ptr"))
        DllCall("CloseHandle", "Ptr", NumGet(pi, A_PtrSize, "Ptr"))

        output := ""
        buf := Buffer(4096, 0)
        while DllCall("ReadFile", "Ptr", hRead, "Ptr", buf.Ptr, "UInt", 4096, "UInt*", &bytesRead := 0, "Ptr", 0) && bytesRead > 0 {
            output .= StrGet(buf, bytesRead, "UTF-8")
        }
        DllCall("CloseHandle", "Ptr", hRead)
        return Trim(output, " `t`r`n")
    }

    DllCall("CloseHandle", "Ptr", hWrite)
    DllCall("CloseHandle", "Ptr", hRead)
    return ""
}