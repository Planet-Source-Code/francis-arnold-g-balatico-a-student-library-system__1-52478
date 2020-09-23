VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLibEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Librarian"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Change Password"
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
      Height          =   1140
      Left            =   60
      TabIndex        =   9
      Top             =   3660
      Width           =   6195
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1920
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   660
         Width           =   4110
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1920
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   240
         Width           =   4110
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "New &Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   450
         TabIndex        =   11
         Top             =   330
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&onfirm Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   750
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Search Librarian"
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
      Height          =   750
      Left            =   60
      TabIndex        =   1
      Top             =   855
      Width           =   6195
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
         Left            =   1935
         TabIndex        =   2
         Top             =   240
         Width           =   3195
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   435
         Left            =   5205
         TabIndex        =   3
         Top             =   210
         Width           =   825
         _ExtentX        =   1455
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
         Image           =   "frmLibEdit.frx":0000
         cBack           =   16777215
      End
      Begin VB.Label lblBorId 
         AutoSize        =   -1  'True
         Caption         =   "Librarian &ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   510
         TabIndex        =   4
         Top             =   330
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Librarian Information"
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
      Height          =   2010
      Left            =   60
      TabIndex        =   0
      Top             =   1635
      Width           =   6195
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   360
         Index           =   1
         Left            =   1920
         TabIndex        =   17
         Top             =   255
         Width           =   4110
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   360
         Index           =   2
         Left            =   1920
         TabIndex        =   16
         Top             =   675
         Width           =   4110
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   360
         Index           =   3
         Left            =   1920
         TabIndex        =   15
         Top             =   1095
         Width           =   4110
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   360
         Index           =   4
         Left            =   1920
         TabIndex        =   14
         Top             =   1515
         Width           =   4110
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&First Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   795
         TabIndex        =   21
         Top             =   345
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Middle Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   20
         Top             =   765
         Width           =   1170
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Last Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   795
         TabIndex        =   19
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co&ntact #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   855
         TabIndex        =   18
         Top             =   1590
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc AdoEdit 
      Height          =   390
      Left            =   285
      Top             =   5685
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   688
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
   Begin Project1.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   4845
      TabIndex        =   5
      Top             =   5655
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
   Begin Project1.lvButtons_H cmdSave 
      Height          =   405
      Left            =   3465
      TabIndex        =   6
      Top             =   5655
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Update"
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
      ImgSize         =   32
      Enabled         =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H cmdReload 
      Height          =   405
      Left            =   2055
      TabIndex        =   7
      Top             =   5655
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Reload"
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
      ImgSize         =   32
      Enabled         =   0   'False
      cBack           =   12632256
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Authorize Update"
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
      Left            =   60
      TabIndex        =   22
      Top             =   4815
      Width           =   6195
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   1920
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   270
         Width           =   4110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Old Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   24
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   60
      Picture         =   "frmLibEdit.frx":0CDA
      Stretch         =   -1  'True
      Top             =   5580
      Width           =   6195
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLibEdit.frx":3661
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   780
      TabIndex        =   8
      Top             =   240
      Width           =   5280
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   195
      Picture         =   "frmLibEdit.frx":36FB
      Top             =   195
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   -15
      Picture         =   "frmLibEdit.frx":43C5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6345
   End
End
Attribute VB_Name = "frmLibEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReload_Click()
    AdoEdit.Refresh
    txtInfo(1).Text = AdoEdit.Recordset.Fields("Fname")
    txtInfo(2).Text = AdoEdit.Recordset.Fields("Mname")
    txtInfo(3).Text = AdoEdit.Recordset.Fields("Lname")
    txtInfo(4).Text = AdoEdit.Recordset.Fields("Contact")
End Sub

Private Sub cmdSave_Click()
    Dim Fields As String
    Dim counter As Byte
    Dim pass As String
    
    pass = AdoEdit.Recordset.Fields("Admin_Pass")
    AdoEdit.Refresh
    
        If txtInfo(7).Text = pass Then
        
        'field validation
            For counter = 1 To 4
                If Trim(txtInfo(counter).Text) = "" Then
                    MsgBox "Required field missing. Please provide all required information.", vbOKOnly + vbExclamation, "Library System"
                    txtInfo(counter).SetFocus
                    Exit Sub
                End If
            Next counter
        
            If Not txtInfo(1).Text = AdoEdit.Recordset.Fields("Fname") Then
                AdoEdit.Recordset.Fields("Fname") = txtInfo(1).Text
                Fields = Fields & " First Name"
            End If
            
            If Not txtInfo(2).Text = AdoEdit.Recordset.Fields("Mname") Then
                AdoEdit.Recordset.Fields("Mname") = txtInfo(2).Text
                Fields = Fields & " Middle Name"
            End If
            
            If Not txtInfo(3).Text = AdoEdit.Recordset.Fields("Lname") Then
                AdoEdit.Recordset.Fields("Lname") = txtInfo(3).Text
                Fields = Fields & " Last Name"
            End If
            
            If Not txtInfo(4).Text = AdoEdit.Recordset.Fields("Contact") Then
                AdoEdit.Recordset.Fields("Contact") = txtInfo(4).Text
                Fields = Fields & " Contact Number"
            End If
            
            If Trim(txtInfo(5).Text) <> "" Then
                If Not Trim(txtInfo(5).Text) = Trim(txtInfo(6).Text) Then
                    MsgBox "Password change failed. New password and confirmation password not the same.", vbOKOnly + vbExclamation, "Library System"
                Else
                    AdoEdit.Recordset.Fields("Admin_Pass") = Trim(txtInfo(5).Text)
                    Fields = Fields & " Password"
                    
                End If
            End If
            
            If Trim(Fields) = "" Then
                Exit Sub
            End If
            
            AdoEdit.Recordset.Update
            
            MsgBox "Change succesful on the following fields: " & Fields & ".", vbOKOnly + vbInformation, "Library System"
            Call ClearAll
            txtSearch.SetFocus
            SendKeys HiLyt
        Else
            MsgBox "Password invalid. Authorization failed.", vbOKOnly + vbExclamation, "Library System"
            txtInfo(7).SetFocus
            SendKeys HiLyt
            Exit Sub
        End If
    
End Sub

Private Sub cmdSearch_Click()
    Dim counter As Byte
    On Error GoTo NotFound
        AdoEdit.Refresh
        AdoEdit.Recordset.Find ("Admin_Name = '" & Trim(txtSearch.Text) & "'")
         
        txtInfo(1).Text = AdoEdit.Recordset.Fields("Fname")
        txtInfo(2).Text = AdoEdit.Recordset.Fields("Mname")
        txtInfo(3).Text = AdoEdit.Recordset.Fields("Lname")
        txtInfo(4).Text = AdoEdit.Recordset.Fields("Contact")
        
        For counter = 1 To 7
            txtInfo(counter).BackColor = vbWhite
            txtInfo(counter).Enabled = True
        Next counter
        
        cmdReload.Enabled = True
        cmdSave.Enabled = True
        txtInfo(1).SetFocus
        SendKeys HiLyt
        
        Exit Sub
NotFound:
        MsgBox "The librarian profile you requested could not be found.", vbOKOnly + vbExclamation, "Library System"
        
        Call ClearAll
        txtSearch.SetFocus
        SendKeys HiLyt
End Sub

Private Sub Form_Load()
    Call ConnectToDb(AdoEdit, "Admin")
End Sub

Private Sub ClearAll()
    Dim counter As Byte
    
    For counter = 1 To 7
        txtInfo(counter).Text = ""
        txtInfo(counter).BackColor = &HE0E0E0
        txtInfo(counter).Enabled = False
    Next counter
    
    cmdReload.Enabled = False
    cmdSave.Enabled = False
    
    
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    SendKeys HiLyt
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Not Index = 7 Then
        txtInfo(Index + 1).SetFocus
    End If
    
    If KeyAscii = 13 And Index = 7 Then
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtSearch_Change()
    ClearAll
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
End Sub
