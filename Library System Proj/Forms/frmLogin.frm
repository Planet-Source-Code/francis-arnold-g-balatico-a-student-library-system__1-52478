VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Library System Login"
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.lvButtons_H cmdLogOk 
      Height          =   405
      Left            =   1125
      TabIndex        =   4
      Top             =   1635
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   714
      Caption         =   "&Enter"
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
   Begin MSAdodcLib.Adodc ADOLog 
      Height          =   480
      Left            =   60
      Top             =   2610
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   847
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "DBase Login"
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
   Begin VB.TextBox txtLogPass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B5742D&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1230
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1980
   End
   Begin VB.TextBox txtLogUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B5742D&
      Height          =   315
      Left            =   1230
      MaxLength       =   10
      TabIndex        =   1
      Top             =   690
      Width           =   1980
   End
   Begin Project1.lvButtons_H cmdLogExit 
      Height          =   405
      Left            =   2235
      TabIndex        =   5
      Top             =   1635
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   714
      Caption         =   "E&xit"
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
      Height          =   480
      Left            =   1305
      Top             =   2610
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   847
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "DBase Login"
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
   Begin VB.Label lblDrag 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   75
      TabIndex        =   6
      Top             =   45
      Width           =   3240
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   330
      TabIndex        =   2
      Top             =   1245
      Width           =   885
   End
   Begin VB.Label lblLogUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User &Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   750
      Width           =   1005
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdLogExit_Click()
'Exits the application if msgbox returns OK

If MsgBox("Exiting the application. Proceed?", vbOKCancel + vbQuestion, "Library System") = vbOK Then
    End
Else
'Cancel operation
    Exit Sub
End If

End Sub

Private Sub cmdLogOk_Click()
    Dim tempUser As String
    Dim tempPass As String
    
    On Error GoTo NotFound 'traps the error generated when no similar records are found
    
    ADOLog.Refresh
    'searches the current table for the username
    ADOLog.Recordset.Find ("Admin_Name = '" & UCase(txtLogUser.Text) & "'")
    
    tempUser = UCase(ADOLog.Recordset.Fields("Admin_Name")) 'temporarily store username for comparison
    tempPass = ADOLog.Recordset.Fields("Admin_Pass") 'temporarily store password for comparison
    
    If tempUser = UCase(txtLogUser.Text) Then 'if found then it validates the password
        If tempPass = txtLogPass.Text Then 'if password is valid it executes next command...
            
            'set values for global variables
            LibFName = ADOLog.Recordset.Fields("Fname")
            LibMname = ADOLog.Recordset.Fields("Mname")
            LibLname = ADOLog.Recordset.Fields("Lname")
            LibUser = tempUser
            LibPass = tempPass
            On Error Resume Next
            Fines = AdoInsti.Recordset.Fields("Fine")
 
            MaxBooks = AdoInsti.Recordset.Fields("MaxBooks")
            
            Call DueCount
            Call OverDueCount
            Call BorrowedCount
            Call TotalCount
            
            On Error Resume Next
            AdoInsti.Refresh
            AdoInsti.Recordset.MoveFirst
            
            On Error Resume Next
            LibInsti = AdoInsti.Recordset.Fields("Institution")
            
                If Trim(LibInsti) = "" Then
                    LibInsti = "Student Library System"
                End If
            
            frmInsignia.lblName.Caption = LibFName & " " & LibLname 'assigns the librarian name to FrmInsignia
            
            MDIMain.Show
            
            Unload frmLogin
        Else '...else it notifies user that it is invalid
            MsgBox "Invalid Password. Access Denied.", vbOKOnly + vbExclamation, "Library System: Login"
            txtLogPass.SetFocus
            SendKeys HiLyt
        End If
    Else 'if username is invalid, user is notified
        MsgBox "Invalid Username. Access Denied.", vbOKOnly + vbExclamation, "Library System: Login"
        txtLogUser.Text = ""
    End If
    Exit Sub
    
NotFound: 'Notifies the user that the username provided does not exist
        MsgBox "User name does not exist. Please enter a valid User Name.", vbOKOnly + vbExclamation, "Library System"
        txtLogUser.SetFocus
        SendKeys HiLyt
        ADOLog.Refresh
End Sub



Private Sub Form_Load()
On Error GoTo ErrHandler 'we will use an intentional error to facilitate new user input


frmSplash.lblSplashStat.Caption = "Accessing database..."
    
   Call ConnectToDb(ADOLog, "Admin")
   Call ConnectToDb(AdoInsti, "Setup")
   
frmSplash.lblSplashStat.Caption = "Initialization complete!"
    
    'refreshes database status
    ADOLog.Refresh
    
    'intentionally create an error situation if there is no record in the Dbase
    ADOLog.Recordset.MoveFirst
    
    frmLogin.Show
Exit Sub


ErrHandler: 'handles the error by asking the new user to input initial settings
    frmSplash.Hide
    
    Load frmAdminSetup
    
    frmSplash.lblSplashStat.Caption = "Preparing initial setup..."
    
    frmAdminSetup.Show vbModal
        
    frmLogin.Show
    
    ADOLog.Refresh
    
End Sub

Private Sub lblDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call FormDrag(Me)
End Sub

Private Sub txtLogPass_GotFocus()
    SendKeys HiLyt
End Sub

Private Sub txtLogPass_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then 'if enter is pressed execute next command
        Call cmdLogOk_Click
    End If
End Sub

Private Sub txtLogUser_GotFocus()
     SendKeys HiLyt
End Sub

Private Sub txtLogUser_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then 'if enter is pressed execute next command
        txtLogPass.SetFocus
        SendKeys HiLyt
    End If
End Sub

