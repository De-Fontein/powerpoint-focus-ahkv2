#Requires AutoHotkey v2.0

#SingleInstance Force

SetTitleMatchMode(2)

; Declare globals
global SCREEN_CLASS := "ahk_class screenClass"
global POWERPOINT_CLASS := "PowerPoint Presenter View"

try {
    ; Attempt to get the window title of the active PowerPoint window
    isPowerPointActive() {
        return WinExist(SCREEN_CLASS) && WinExist(POWERPOINT_CLASS)
    }

    ; Only applies the modifications if PowerPoint is active
    #HotIf isPowerPointActive()

    ; Down Arrow - Advance slide or start presenting from the first slide
    *Down:: {
        ControlSend("{PgDn}", POWERPOINT_CLASS)
        return
    }

    ; Up Arrow - Go back one slide
    *Up:: {
        ControlSend("{PgUp}", POWERPOINT_CLASS)
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
} catch error {
    MsgBox("An error occurred: " error.Message)
}