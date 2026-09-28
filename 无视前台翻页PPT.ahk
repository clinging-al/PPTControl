#Requires AutoHotkey v2.0
#SingleInstance Force

SetTitleMatchMode 2

; ================== WPS PPT COM 翻页控制 ==================
; Adding PgUp/PgDn allows you to turn pages of PPT regardless of whether it is a front-end application,by clinging

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
