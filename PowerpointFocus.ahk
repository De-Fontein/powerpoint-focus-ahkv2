#Requires AutoHotkey v2.0

#SingleInstance Force

SetTitleMatchMode(2)

; Attempt to get the window title of the active PowerPoint window
isPowerPointActive() {
    if WinExist("ahk_class screenClass") && WinExist("PowerPoint Presenter View") {
        return true
    }
    return false
}

; Only applies the modifications if PowerPoint is active
#HotIf isPowerPointActive()

; Down Arrow - Advance slide or start presenting from the first slide
*Down:: {
    ControlSend("{PgDn}", "PowerPoint Presenter View")
    return
}

; Up Arrow - Go back one slide
*Up:: {
    ControlSend("{PgUp}", "PowerPoint Presenter View")
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