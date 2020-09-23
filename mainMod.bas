Attribute VB_Name = "mainMod"
' Another Prev Instance Example, using subclassing
' Thomas Greenwood 02/02/04
' thomasgreenwood@2die4.com

' Win32 Api Declares
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' ^ Used to change window proc
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' ^ Used to call default window proc for window
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
' ^ Send Message to window (Proc)
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
' ^ Used, funnily enough...to flash the window :D
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' ^ Used to find the window if we have a prev instance
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
' ^ Check to see if we have a valid window.

Private Const WM_CUSTOM As Long = &H22 'Simply Defines my custom message
Private Const GWL_WNDPROC As Long = (-4) ' Used when creating the Form Hook
Public Const NormalTit As String = "Only One" ' The caption to look for

Private PrevProc As Long 'Stores the location of the Old Window Proc
' For when we have finished with the subclass
' This should always be restored on queryunload or unload

Public Sub HookForm(Obj As Object)
    'Set the Window Proc For the Form to our custom Proc Below
    PrevProc = SetWindowLong(Obj.hwnd, GWL_WNDPROC, AddressOf MainWindowProc)
    ' PrevProc: Store the old address
End Sub


Public Sub UnHookForm(Obj As Object)
    ' Restore the Old Window Proc
    SetWindowLong Obj.hwnd, GWL_WNDPROC, PrevProc

End Sub

Public Function MainWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' This function is called when the form recieves a message
' Whatever processing is done here should be very fast!!
    If uMsg = WM_CUSTOM Then ' If its our custom message, let the user know a prevInstance is opened
        FlashWindow hwnd, True ' Flash the Title Bar and TaskBar
        oneForm.SetFocus ' Set Focus to the main form (dont think it does anything due to flash)
    End If
    ' Once done, call the old window proc so all other messages are dealt with
    MainWindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
End Function

Public Sub CheckInstance()
    If App.PrevInstance = True Then
        ' If there is another app running...
        prevhwnd = FindWindow(vbNullString, NormalTit)
        ' Find the window
        If IsWindow(prevhwnd) Then
        ' If we found the window..
        ' Send out custom message to the Window Proc so the previnstance knows were here
            If SendMessage(prevhwnd, WM_CUSTOM, 1, 1) <> 0 Then
                ' If theres a problem sending the message, tell the user
                MsgBox "Message was not 0"
            End If
            End ' End this instance
        End If
    End If
End Sub
