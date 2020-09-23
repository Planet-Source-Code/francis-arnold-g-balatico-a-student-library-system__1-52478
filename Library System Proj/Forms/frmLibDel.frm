VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLibDel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Librarian Profile"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   203
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   236
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1410
      Top             =   3135
   End
   Begin MSAdodcLib.Adodc AdoLibDel 
      Height          =   330
      Left            =   30
      Top             =   3105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Corresponding Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   30
      TabIndex        =   3
      Top             =   1680
      Width           =   3450
      Begin VB.TextBox txtPass 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   135
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   270
         Width           =   3195
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Enter Librarian Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   3450
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   2625
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   435
         Left            =   2790
         TabIndex        =   2
         Top             =   195
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   767
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
         cBhover         =   11891757
         cGradient       =   11891757
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "frmLibDel.frx":0000
         cBack           =   16777215
      End
   End
   Begin Project1.lvButtons_H cmdClose 
      Height          =   405
      Left            =   2055
      TabIndex        =   5
      Top             =   2505
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
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
   Begin Project1.lvButtons_H cmdDel 
      Height          =   405
      Left            =   675
      TabIndex        =   6
      Top             =   2505
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Delete"
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
      ImgAlign        =   1
      Image           =   "frmLibDel.frx":0CDA
      cBack           =   12632256
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   225
      Picture         =   "frmLibDel.frx":19B4
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Deleting an existing librarian profile requires the password of the profile to be deleted."
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   840
      TabIndex        =   7
      Top             =   135
      Width           =   2565
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmLibDel.frx":267E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3540
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   30
      Picture         =   "frmLibDel.frx":571E
      Stretch         =   -1  'True
      Top             =   2430
      Width           =   3450
   End
End
Attribute VB_Name = "frmLibDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private UserName As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
On Error Resume Next


    If txtPass.Text = AdoLibDel.Recordset.Fields(1) Then
        If MsgBox("This will delete the requested profile. Proceed?", vbYesNo + vbQuestion, "Library System") = vbYes Then
            Status "Deleting profile..."
            UserName = UCase(txtSearch.Text)
            AdoLibDel.Recordset.Delete
            AdoLibDel.Refresh
            txtPass.Text = ""
            txtPass.Enabled = False
            txtPass.BackColor = &H808080
            txtSearch.Text = ""
            cmdDel.Enabled = False
            Status "Ready"
            
            If AdoLibDel.Recordset.RecordCount = 0 Then
                Me.Hide
                MsgBox "There are no Librarian Profiles available. Logging out initiated." & vbCrLf + vbCrLf & "To continue system access, please create a new Librarian Profile.", vbOKOnly + vbInformation, "Library System"
                MDIMain.Hide
                frmAdminSetup.Show vbModal
                frmLogin.Show
                Unload Me
                Exit Sub
            End If
                
            If UCase(UserName) = UCase(LibUser) Then
                Me.Hide
                MsgBox "The profile deleted is currently being used." & vbCrLf + vbCrLf & "Automatic Logout initiated. To continue use, please log in with a different Librarian Profile.", vbOKOnly + vbInformation, "Library System"
                MDIMain.Hide
                frmLogin.Show
                Unload Me
                Exit Sub
            End If
                        
        Else
            Exit Sub
        End If
    Else
        MsgBox "The password entered does not match with the selected profile password. Deletion denied.", vbOKOnly + vbExclamation, "Library System"
        txtPass.SetFocus
        SendKeys HiLyt
        
        Exit Sub
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo NotFound
    AdoLibDel.Refresh
    AdoLibDel.Recordset.Find ("Admin_Name = '" & UCase(txtSearch.Text) & "'")
    
    UserName = AdoLibDel.Recordset.Fields("Admin_Name")
    
    cmdDel.Enabled = True
    txtPass.Enabled = True
    txtPass.BackColor = vbWhite
    txtPass.SetFocus
    Exit Sub

NotFound:
    MsgBox "User name does not exist. Please enter a valid User Name.", vbOKOnly, "Library System"
    cmdDel.Enabled = False
    txtSearch.SetFocus
    SendKeys HiLyt
    AdoLibDel.Refresh
End Sub

Private Sub Form_Load()
    Call ConnectToDb(AdoLibDel, "Admin")
End Sub

Private Sub Timer1_Timer()
    
        If Trim(txtPass.Text) = "" Then
            cmdDel.Enabled = False
        Else
            cmdDel.Enabled = True
        End If
        
        If Trim(txtSearch.Text) = "" Then
            cmdSearch.Enabled = False
        Else
            cmdSearch.Enabled = True
        End If
        
End Sub



Private Sub txtSearch_Change()
    txtPass.Text = ""
    txtPass.Enabled = False
    txtPass.BackColor = &H808080
    cmdDel.Enabled = False
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
End Sub
