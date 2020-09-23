VERSION 5.00
Begin VB.Form oneForm 
   Caption         =   "Only One"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "oneForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Me.Caption = "Not Me" ' Change the caption so we dont find the wrong window
    CheckInstance ' Check for the prev instance
    ' If we get here there is no prev instance so
    HookForm Me ' Hook the form so we can do our notification
    Me.Caption = NormalTit ' Restore the main caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnHookForm Me ' Unhook the form (restore the window proc)
End Sub
