#Requires AutoHotkey v2.0

#SingleInstance Force

SetTitleMatchMode(2)

; Declare globals
global POWERPOINT_APP := "PowerPoint.Application"
global PPT_EXE         := "ahk_exe POWERPNT.EXE"
global SLIDESHOW_CLASS := "ahk_class screenClass"
global PRESENTER_CLASS := "ahk_class PodiumParent"

try {
    ; Attempt to get the window title of the active PowerPoint window
    isPowerPointActive() {
        return WinExist(SLIDESHOW_CLASS " " PPT_EXE) && WinExist(PRESENTER_CLASS " " PPT_EXE)
    }

    ; Only applies the modifications if PowerPoint is active
    #HotIf isPowerPointActive()

    ; Down Arrow - Advance slide or start presenting from the first slide
    *Down:: {
        ComObjActive(POWERPOINT_APP).SlideShowWindows.Item(1).View.Next()
        return
    }

    ; Up Arrow - Go back one slide
    *Up:: {
        ComObjActive(POWERPOINT_APP).SlideShowWindows.Item(1).View.Previous()
        return
    }

    ; Ignore Volume Up
    *Volume_Up:: {
        return
    }

    ; Ignore Volume Down
    *Volume_Down:: {
        return
    }

    ; Ignore Tab
    *Tab:: {
        return
    }

    #HotIf
} catch error as e {
    MsgBox("An error occurred: " e.Message)
}