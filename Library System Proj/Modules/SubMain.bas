Attribute VB_Name = "SubMain"
Option Explicit
'initializations are located at this sub procedure

Public Sub Main()
'detects if application is already open
If App.PrevInstance = True Then
    MsgBox "Student Library System v1.0 is already open.", vbOKOnly + vbInformation, "Library System"
    End
End If

Main_On = False
frmSplash.Show

End Sub
