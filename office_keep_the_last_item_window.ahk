; This is AHK v2 script, it won't work on AHK v1.
#SingleInstance Force
#Warn	; Enable warnings to assist with detecting common errors.
SendMode "Input"	; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir A_ScriptDir	; Ensures a consistent starting directory.


; All translations can be found on https://www.microsoft.com/en-us/language/Search?&searchTerm=You%20placed%20a%20picture%20on%20the%20Clipboard.&langID=591&Source=true&productid=undefined

; English language
phrase1:= "You placed a picture on the Clipboard. Do you want this picture to be available to other applications after you quit"
phrase2:= "Do you want to keep the last item you copied?\n\nIf you do, it may take a bit longer to exit."
phrase_yes:= "Yes"

; ; Polish language
; phrase1:= "W Schowku znajduje się obraz. Czy ten obraz ma być dostępny dla innych aplikacji po zakończeniu pracy z programem"
; phrase2:= "Czy chcesz zachować ostatni skopiowany element?\n\nJeśli tak, zakończenie działania może potrwać trochę dłużej."
; phrase_yes:= "Tak"

Loop
{
    HWND := WinWaitActive("Microsoft ")
    window_text := WinGetText("ahk_id " HWND)

    if InStr(window_text, phrase1, "Off") or InStr(window_text, phrase2, "Off")
    {
        Controls := WinGetControls(HWND)
        for control in Controls
        {   
            MsgBox(control " " ControlGetText(control)) ; for debugging, comment this line if it works ok
            if InStr(control, "Button") ; check if control is of type button
            {
                if InStr(ControlGetText(control), phrase_yes)
                {
                    MsgBox("Button found") ; for debugging, comment this line if it works ok
                    ControlClick(control)
                    break
                }
            }
        }
    }
    else
    {
        MsgBox "Not found" ; for debugging, comment this line if it works ok
    }
}
