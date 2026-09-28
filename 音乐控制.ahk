#Requires AutoHotkey v2.0
#SingleInstance Force

#Include <UIA>

; One-click control of HirenderP1 fade-in and fade-out to play music，by clinging
SetTitleMatchMode 2
targetWin := "ahk_exe HirenderP1.exe"

; 设置为屏幕物理绝对坐标模式
CoordMode "Mouse", "Screen"

; ==============================================================================
; 右 Alt：激活窗口 -> 物理拖拽音量到0% -> 点击停止 -> 点击分组 -> 点击素材播放 -> 点击广播
; ==============================================================================
RAlt::
{
    ; 1. 严格判断窗口是否存在
    hwnd := WinExist(targetWin)
    if !hwnd {
        ToolTip "未找到 Hirender P1 窗口！"
        SetTimer () => ToolTip(), -1500
        return
    }

    WinActivate hwnd
    if !WinWaitActive(hwnd, , 1.5) {
        ToolTip "无法激活窗口"
        SetTimer () => ToolTip(), -1500
        return
    }

    ; 记录鼠标原位
    MouseGetPos &origX, &origY

    try {
        p1El := UIA.ElementFromHandle(hwnd)
        
        ; 点击停止前：先将音量拉到最低 (0%)
        SetVolume(p1El, 0)
        Sleep 50

        ; 停止按钮（做非空判断）
        stopBtn := p1El.WaitElement({Name: "停止", matchMode: 2}, 1500)
        if IsObject(stopBtn) && HasMethod(stopBtn, "Click")
            stopBtn.Click("left")

        ; 2. 定位左侧【上场音乐】分组
        groupEl := p1El.WaitElement({Type: "TreeItem", Name: "上场音乐", matchMode: 2}, 2000)
        if !IsObject(groupEl) || !HasMethod(groupEl, "Click")
            throw Error("未找到【上场音乐】分组")
        
        try groupEl.Select()
        groupEl.Click("left")

        Sleep 300

        ; 3. 定位右侧【上场音乐】素材
        itemEl := p1El.WaitElement({Type: "ListItem", Name: "上场音乐", matchMode: 2}, 2000)
        if !IsObject(itemEl) || !HasMethod(itemEl, "Click")
            throw Error("未找到【上场音乐】素材")
            
        itemEl.Click("left", 1)
		
		; 4. 点击【广播】按钮
        broadcastBtn := p1El.WaitElement({Name: "广播", matchMode: 2}, 1500)
        if !IsObject(broadcastBtn) || !HasMethod(broadcastBtn, "Click")
            throw Error("未找到【广播】按钮")
        broadcastBtn.Click("left")
        ToolTip "已触发广播，等待 1.5 秒..."
		
        ; 5. 鼠标归位
        ; MouseMove origX, origY, 0
        ToolTip "开始播放（音量已推至100%）！"
        SetTimer () => ToolTip(), -1500

    } catch as err {
        if IsSet(origX)
            MouseMove origX, origY, 0
        MsgBox "播放失败！`n错误信息：" err.Message "`n出错位置：" err.What, "错误", 16
    }
}

; ==========================================================
; 右 Ctrl：点击广播 -> 等待1.5秒 -> 点击停止 -> 物理拖拽音量回100%
; ==========================================================
RCtrl::
{
    hwnd := WinExist(targetWin)
    if !hwnd {
        ToolTip "未找到 Hirender P1 窗口！"
        SetTimer () => ToolTip(), -1500
        return
    }

    WinActivate hwnd
    if !WinWaitActive(hwnd, , 1.5) {
        ToolTip "无法激活窗口"
        SetTimer () => ToolTip(), -1500
        return
    }

    MouseGetPos &origX, &origY

    try {
        p1El := UIA.ElementFromHandle(hwnd)

        ; 1. 点击【广播】按钮
        broadcastBtn := p1El.WaitElement({Name: "广播", matchMode: 2}, 1500)
        if !IsObject(broadcastBtn) || !HasMethod(broadcastBtn, "Click")
            throw Error("未找到【广播】按钮")
        broadcastBtn.Click("left")
        ToolTip "已触发广播，等待 2 秒..."

        ; 2. 等待 2 秒
        Sleep 2000

        ; 3. 点击【停止】按钮
        stopBtn := p1El.WaitElement({Name: "停止", matchMode: 2}, 1500)
        if !IsObject(stopBtn) || !HasMethod(stopBtn, "Click")
            throw Error("未找到【停止】按钮")
        stopBtn.Click("left")

        Sleep 200

        ; 4. 物理拖拽滑块复位回 100%
        SetVolume(p1El, 100)

        ; 5. 还原鼠标
        ; MouseMove origX, origY, 0
        ToolTip "广播停止完成，音量已推回100！"
        SetTimer () => ToolTip(), -1500

    } catch as err {
        Click "Up Left"  ; 异常时防卡键
        if IsSet(origX)
            MouseMove origX, origY, 0
        MsgBox "停止操作失败！`n错误信息：" err.Message "`n出错位置：" err.What, "错误", 16
    }
}

; ==========================================================
; 核心函数：判断滑块音量，支持物理拖动至指定百分比 (默认 100%)
; ==========================================================
SetVolume(p1El, targetPercent := 100)
{
    try {
        sliderEl := p1El.WaitElement({Type: "Slider"}, 1500)
        if !IsObject(sliderEl)
            return

        ; 判断当前 Value 已经达到目标值则直接返回
        try {
            if (targetPercent >= 100 && Float(sliderEl.Value) >= 100)
                return
            if (targetPercent <= 0 && Float(sliderEl.Value) <= 0)
                return
        }

        sPos := sliderEl.GetPos()
        
        centerX := sPos.x + (sPos.w // 2)
        topY    := sPos.y + 6            ; 最顶端 100% 极限位置
        botY    := sPos.y + sPos.h - 6   ; 最底端 0% 极限位置

        ; 目标 Y 坐标
        targetY := (targetPercent <= 0) ? botY : topY

        startX := centerX
        startY := 0

        ; 读取当前 Value 百分比推算手柄所在高度
        try {
            curVal := Float(sliderEl.Value)
            startY := botY - ((botY - topY) * (curVal / 100.0))
        } catch {
            ; 兜底：若目标是拉到 0，则默认从偏上开始抓；目标是 100 则从偏下抓
            startY := (targetPercent <= 0) ? (topY + 10) : (botY - 10)
        }

        ; 如果已经在目标位置附近（相差不到 4 像素），无需重复拖拽
        if (Abs(startY - targetY) < 4)
            return

        ; 执行物理拖拽
        MouseMove startX, startY, 0
        Click "Down Left"
        Sleep 30
        MouseMove centerX, targetY, 2
        Sleep 30
        Click "Up Left"
        Sleep 50
    } catch {
        Click "Up Left" ; 确保发生异常时不会卡死鼠标左键
    }
}
