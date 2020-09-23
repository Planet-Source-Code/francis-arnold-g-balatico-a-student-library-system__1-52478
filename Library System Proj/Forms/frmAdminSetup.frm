VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAdminSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome New Librarian!"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdminSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "LIBRARIAN SECURITY SETTINGS"
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   75
      TabIndex        =   15
      Top             =   2985
      Width           =   5415
      Begin VB.TextBox txtUName 
         Height          =   330
         Left            =   1860
         TabIndex        =   9
         Top             =   255
         Width           =   3285
      End
      Begin VB.TextBox txtPass 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   675
         Width           =   3285
      End
      Begin VB.TextBox txtConPass 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1110
         Width           =   3285
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&User Name:"
         Height          =   195
         Left            =   735
         TabIndex        =   8
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   750
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&onfirm Password:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   1170
         Width           =   1575
      End
   End
   Begin VB.Frame fmeLibInfo 
      Appearance      =   0  'Flat
      Caption         =   "LIBRARIAN INFO"
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   75
      TabIndex        =   14
      Top             =   945
      Width           =   5415
      Begin VB.TextBox txtMname 
         Height          =   330
         Left            =   1875
         TabIndex        =   3
         Top             =   675
         Width           =   3285
      End
      Begin VB.TextBox txtContact 
         Height          =   330
         Left            =   1875
         TabIndex        =   7
         Top             =   1530
         Width           =   3285
      End
      Begin VB.TextBox txtLname 
         Height          =   330
         Left            =   1875
         TabIndex        =   5
         Top             =   1110
         Width           =   3285
      End
      Begin VB.TextBox txtFname 
         Height          =   330
         Left            =   1875
         TabIndex        =   1
         Top             =   255
         Width           =   3285
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Middle Name:"
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   750
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Contact #:"
         Height          =   195
         Left            =   840
         TabIndex        =   6
         Top             =   1590
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Last Name:"
         Height          =   195
         Left            =   780
         TabIndex        =   4
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&First Name:"
         Height          =   195
         Left            =   795
         TabIndex        =   0
         Top             =   330
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc AdoAdmin 
      Height          =   420
      Left            =   75
      Top             =   5205
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Project1.lvButtons_H cmdAdminOk 
      Height          =   420
      Left            =   2745
      TabIndex        =   16
      Top             =   4650
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   741
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   11891757
      cFHover         =   11891757
      cBhover         =   14846764
      cGradient       =   14846764
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin Project1.lvButtons_H cmdAdminExit 
      Height          =   420
      Left            =   4125
      TabIndex        =   17
      Top             =   4650
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   741
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   11891757
      cFHover         =   11891757
      cBhover         =   14846764
      cGradient       =   14846764
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin MSAdodcLib.Adodc AdoInsti 
      Height          =   420
      Left            =   1305
      Top             =   5205
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAdminSetup.frx":0CCA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   750
      TabIndex        =   18
      Top             =   135
      Width           =   4710
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   195
      Picture         =   "frmAdminSetup.frx":0D76
      Top             =   180
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmAdminSetup.frx":1A40
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5565
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   75
      Picture         =   "frmAdminSetup.frx":4AE0
      Stretch         =   -1  'True
      Top             =   4575
      Width           =   5415
   End
End
Attribute VB_Name = "frmAdminSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdminExit_Click()
On Error GoTo ErrHandler 'check if the database already contains at least one data item
    'refreshes database status
    AdoAdmin.Refresh

    AdoAdmin.Recordset.MoveFirst 'error occurs if database is empty else exits the module
    
    frmAdminSetup.Hide
    
    AdoAdmin.Refresh
   
    Unload frmAdminSetup
Exit Sub

ErrHandler: 'handles the error
    If MsgBox("There are still no Librarian Accounts. Close and exit application?", vbOKCancel + vbQuestion, "Library System") = vbOK Then
        End
    Else
        Exit Sub
    End If

End Sub

Private Sub cmdAdminOk_Click()
On Error GoTo ErrHandler 'handles expected unsuccessful data entry error

'form field validation
If Trim(txtFname.Text) = "" Or Trim(txtMname.Text) = "" Or Trim(txtLname.Text) = "" _
    Or Trim(txtContact.Text) = "" Or Trim(txtUName.Text) = "" Or Trim(txtPass.Text) = "" Or _
    Trim(txtConPass.Text) = "" Then
    
    MsgBox "Required field missing. Please fill up ALL the fields.", vbOKOnly + vbExclamation, "Library System"
    
    'checks the missing field and focuses on it
    If Trim(txtFname.Text) = "" Then
        txtFname.Text = ""
        txtFname.SetFocus
    ElseIf Trim(txtMname.Text) = "" Then
        txtMname.Text = ""
        txtMname.SetFocus
    ElseIf Trim(txtLname.Text) = "" Then
        txtLname.Text = ""
        txtLname.SetFocus
    ElseIf Trim(txtContact.Text) = "" Then
        txtContact.Text = ""
        txtContact.SetFocus
    ElseIf Trim(txtUName.Text) = "" Then
        txtUName.Text = ""
        txtUName.SetFocus
    ElseIf Trim(txtPass.Text) = "" Then
        txtPass.Text = ""
        txtPass.SetFocus
    Else
        txtConPass.Text = ""
        txtConPass.SetFocus
    End If

    Exit Sub
End If

If IsNumeric(txtContact.Text) = False Then
    MsgBox "Cannot accept non-numeric input for contact number.", vbOKOnly + vbExclamation, "Library System"
    txtContact.SetFocus
    SendKeys HiLyt
    Exit Sub
End If
'if all fields are ok then transfer data to the database
'checks if the password typed is similar with the password confirmation
If Trim(txtPass.Text) = Trim(txtConPass.Text) Then
    
    AdoAdmin.Refresh
    AdoAdmin.Recordset.AddNew

    With AdoAdmin.Recordset
        .Fields(0) = txtUName.Text
        .Fields(1) = txtPass.Text
        .Fields(2) = txtFname.Text
        .Fields(3) = txtMname.Text
        .Fields(4) = txtLname.Text
        .Fields(5) = txtContact.Text
        
    End With
    AdoAdmin.Recordset.Update
    
    'confirms that data has already been entered to the database
    AdoAdmin.Refresh
 
    
    AdoAdmin.Recordset.MoveFirst 'will generate an error if data has not been entered
    
    txtFname.Text = ""
    txtLname.Text = ""
    txtMname.Text = ""
    txtContact.Text = ""
    txtUName.Text = ""
    txtPass.Text = ""
    txtConPass.Text = ""
    
    'message confirmation that data entry is successful and asks continuation..
    If MsgBox("Administrator Data entry successful. Enter another record?", vbYesNo + vbQuestion, "Library System") = vbYes Then
        AdoAdmin.Refresh
        Exit Sub
    Else
        frmAdminSetup.Hide
                       
        AdoAdmin.Refresh
       
        Unload frmAdminSetup
        Exit Sub
    End If
    
Else
    MsgBox "The password and password confirmation inputted are not the same." & vbCrLf + vbCrLf & "Please re-confirm password", vbOKOnly + vbExclamation, "Library System"
    txtConPass.Text = ""
    txtConPass.SetFocus
    Exit Sub
End If
Exit Sub

ErrHandler:
    MsgBox "User name already exists. Please choose a different user name.", vbOKOnly, "Library system"
    txtUName.SetFocus
    SendKeys HiLyt
    Exit Sub
End Sub

Private Sub Form_Load()

Call ConnectToDb(AdoAdmin, "Admin")
On Error GoTo ErrHandler
Call ConnectToDb(AdoInsti, "Setup")
    AdoInsti.Refresh
    AdoInsti.Recordset.MoveFirst
Exit Sub

ErrHandler:
   MsgBox "WELCOME! It seems that this is the first time that you would use this system." & vbCrLf + vbCrLf & "Please enter your institution's name", vbOKOnly + vbInformation, "Library System - Setup System"
        frmInsti.Show vbModal
    
End Sub

    

Private Sub txtConPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAdminOk.SetFocus
    End If
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtUName.SetFocus
    SendKeys HiLyt
End If

End Sub

Private Sub txtFname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMname.SetFocus
    SendKeys HiLyt
End If
End Sub


Private Sub txtLname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtContact.SetFocus
    SendKeys HiLyt
End If
End Sub

Private Sub txtMname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtLname.SetFocus
    SendKeys HiLyt
End If

End Sub


Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtConPass.SetFocus
    SendKeys HiLyt
End If

End Sub

Private Sub txtUName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPass.SetFocus
    SendKeys HiLyt
End If

End Sub
