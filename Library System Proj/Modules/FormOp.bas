Attribute VB_Name = "FormOp"
Option Explicit


'API declarations for dragging form
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'variables for ADODB
Dim Connect As New ADODB.Connection

Dim rsOverDueBooks As New ADODB.Recordset
Dim rsDueBooks As New ADODB.Recordset
Dim rsDate As New ADODB.Recordset
Dim rsBooks As New ADODB.Recordset
Dim rsBorBooks As New ADODB.Recordset


Public Sub ConDB()
    On Error Resume Next
    Connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\Lib_Dbase.mdb;Persist Security Info=False;Jet OLEDB:Database Password= crimson119"
End Sub

Public Sub OverDueCount()
    
    Call ConDB
    
    rsOverDueBooks.Open "Select * From Book_Loan where Date_Due < '" & Date & "' and Status = 0", Connect, adOpenStatic, adLockOptimistic
    OverDue = rsOverDueBooks.RecordCount
    
    Set rsOverDueBooks = Nothing
    Set Connect = Nothing
    
End Sub

Public Sub DueCount()
    
    Call ConDB
    
    rsDueBooks.Open "Select * From Book_Loan where Date_Due = '" & Date & "' and Status = 0", Connect, adOpenStatic, adLockOptimistic
    DueBooks = rsDueBooks.RecordCount
    
    Set rsDueBooks = Nothing
    Set Connect = Nothing
    
End Sub

Public Sub BorrowedCount()
    
    Call ConDB
    
    rsBorBooks.Open "Select * From Book where StatusID = 2", Connect, adOpenStatic, adLockOptimistic
    BorrowedBooks = rsBorBooks.RecordCount
    
    Set rsBorBooks = Nothing
    Set Connect = Nothing
    
End Sub

Public Function ChkAcount(str As String) As Boolean
    
    Call ConDB
    
    rsBorBooks.Open "Select * From Book_Loan where Borrower_Id = '" & str & "' and Status = 0", Connect, adOpenStatic, adLockOptimistic
    If Not rsBorBooks.RecordCount = 0 Then
        ChkAcount = True
    Else
        ChkAcount = False
    End If
    
    Set rsBorBooks = Nothing
    Set Connect = Nothing

End Function


Public Sub TotalCount()
    
    Call ConDB
    
    rsBooks.Open "Select * From Book", Connect, adOpenStatic, adLockOptimistic
    TotalBooks = rsBooks.RecordCount
    
    Set rsBooks = Nothing
    Set Connect = Nothing
    
End Sub

Public Function AppDir() As String
    If Right$(App.Path, 1) = "\" Then
        AppDir = App.Path
    Else
        AppDir = App.Path & "\"
    End If
End Function

'PROCEDURE TO CENTER A CHILD FORM ONTO A PARENT FORM
Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form) 'used for the frmInsignia

    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
    Childfrm.Top = (Parentfrm.ScaleHeight \ 2) - (Childfrm.Height \ 2)

End Sub

Public Sub ConnectToDb(adoObj As Adodc, AdoRec As String) 'for table Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\Lib_Dbase.mdb;Persist Security Info=False; Jet OLEDB:Database Password = crimson119"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdTable
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub

Public Sub SQLDB(adoObj As Adodc, AdoRec As String) 'for SQL Recordsource

    'Loads the database and provides the database password
    adoObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\Lib_Dbase.mdb;Persist Security Info=False; Jet OLEDB:Database Password = crimson119"
    
    'Sets the command type to Table
    adoObj.CommandType = adCmdText
    
    'Loads the source table of info
    adoObj.RecordSource = AdoRec

    'refreshes database status
    adoObj.Refresh
End Sub

Public Sub FormDrag(frmName As Form) 'procedure to drag a no-titlebar form
    ReleaseCapture
    Call SendMessage(frmName.hwnd, &HA1, 2, 0&)
End Sub

Public Sub NoRec(frmName As Form) 'displays a message confirmation box that states there are no existing records

    Call Status("Ready")
    
    If MsgBox("There is no existing borrower's record." & vbCrLf + vbCrLf & "Would you like to create a new record?", vbYesNo + vbExclamation, "Library System") = vbNo Then
        Unload frmName
    Else
        Unload frmName
        frmAddStud.Show vbModal
    End If
    
End Sub

Public Sub Status(Stat As String) 'displays text in the first panel of the status bar
    MDIMain.stbMain.Panels(1).Text = Stat
End Sub

Public Sub Missing()
    MsgBox "Required fields missing. Please provide information to all the fields.", vbOKCancel + vbExclamation, "Library System"
End Sub

