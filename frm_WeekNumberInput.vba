VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_WeekNumberInput 
   Caption         =   "Enter week number"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3135
   OleObjectBlob   =   "frm_WeekNumberInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_WeekNumberInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_weeknumberCancel_Click()

Unload Me

End Sub

Private Sub cmd_weeknumberOK_Click()

If IsNumeric(txt_weeknumberinput.Value) And txt_weeknumberinput.Value > 0 And txt_weeknumberinput.Value < 54 Then
    Call TasklistFormat(txt_weeknumberinput.Value)
Else: MsgBox "Please enter a whole number between 1 and 53"
End If

Unload Me

End Sub
