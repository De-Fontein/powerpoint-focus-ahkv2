#Requires AutoHotkey v2.0

#SingleInstance Force

SetTitleMatchMode(2)

; Attempt to get the window title of the active PowerPoint window
getWindowTitle() {
    if WinExist("ahk_class screenClass") || WinActive("PowerPoint Presenter View") {
        return "PowerPoint Presenter View"
    }
    return ""
}

; Bind to Down Arrow - Advance slide or start presenting from the first slide
*Down:: {
    title := getWindowTitle()
    if (title != "") {
        ControlSend("{PgDn}", title)
    } else {
        Send("{Down}")
    }

    return
}

; Bind to Up Arrow - Go back one slide
*Up:: {
    title := getWindowTitle()
    if (title != "") {
        ControlSend("{PgUp}", title)
    } else {
        Send("{Up}")
    }
    return
}

; Ignore Volume Up when a PowerPoint slide show is active
*Volume_Up:: {
    title := getWindowTitle()
    if title == "" {
        Send("{Volume_Up}")
    }
    return
}

; Ignore Volume Down when a PowerPoint slide show is active
*Volume_Down:: {
    title := getWindowTitle()
    if (title == "") {
        Send("{Volume_Down}")
    }
    return
}

; Ignore Tab while powerpoint is open.
*Tab:: {
    title := getWindowTitle()
    if (title == "") {
        Send("{Tab}")
    }
    return
}