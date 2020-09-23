VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddStud 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Borrower Record"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "frmAddStud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   1410
      Top             =   4545
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoAddStud 
      Height          =   390
      Left            =   150
      Top             =   4590
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "NEW BORROWER INFORMATION"
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
      Height          =   2955
      Left            =   60
      TabIndex        =   15
      Top             =   930
      Width           =   6735
      Begin VB.TextBox txtBorInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D3F2D&
         Height          =   360
         Index           =   0
         Left            =   3435
         TabIndex        =   1
         Top             =   255
         Width           =   3135
      End
      Begin VB.PictureBox picID 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00B5742D&
         Height          =   1530
         Left            =   180
         ScaleHeight     =   1470
         ScaleWidth      =   1470
         TabIndex        =   16
         Top             =   300
         Width           =   1530
         Begin VB.Image imgID 
            Height          =   1500
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Photo"
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
            TabIndex        =   17
            Top             =   615
            Width           =   1110
         End
      End
      Begin VB.TextBox txtBorInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D3F2D&
         Height          =   360
         Index           =   5
         Left            =   3435
         TabIndex        =   11
         Top             =   2430
         Width           =   3135
      End
      Begin VB.TextBox txtBorInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D3F2D&
         Height          =   360
         Index           =   4
         Left            =   3435
         TabIndex        =   9
         Top             =   1995
         Width           =   3135
      End
      Begin VB.TextBox txtBorInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D3F2D&
         Height          =   360
         Index           =   3
         Left            =   3435
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtBorInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D3F2D&
         Height          =   360
         Index           =   2
         Left            =   3435
         TabIndex        =   5
         Top             =   1125
         Width           =   3135
      End
      Begin VB.TextBox txtBorInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D3F2D&
         Height          =   360
         Index           =   1
         Left            =   3435
         TabIndex        =   3
         Top             =   690
         Width           =   3135
      End
      Begin Project1.lvButtons_H cmdBrowsePic 
         Height          =   570
         Left            =   195
         TabIndex        =   12
         Top             =   1875
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1005
         Caption         =   "Browse &Photo"
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
         Image           =   "frmAddStud.frx":0CCA
         ImgSize         =   24
         cBack           =   16777215
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contact &Number:"
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
         Index           =   5
         Left            =   1935
         TabIndex        =   10
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Course:"
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
         Left            =   2715
         TabIndex        =   8
         Top             =   2085
         Width           =   660
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Index           =   3
         Left            =   2400
         TabIndex        =   6
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Index           =   2
         Left            =   2205
         TabIndex        =   4
         Top             =   1215
         Width           =   1170
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   2400
         TabIndex        =   2
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblborInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Borrower ID:"
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
         Left            =   2295
         TabIndex        =   0
         Top             =   345
         Width           =   1080
      End
   End
   Begin Project1.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   5310
      TabIndex        =   14
      Top             =   4035
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
      Left            =   3945
      TabIndex        =   13
      Top             =   4035
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
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
      cBack           =   12632256
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddStud.frx":19A4
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1050
      TabIndex        =   18
      Top             =   210
      Width           =   5535
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   345
      Picture         =   "frmAddStud.frx":1A44
      Top             =   180
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   15
      Picture         =   "frmAddStud.frx":270E
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   6840
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   60
      Picture         =   "frmAddStud.frx":57AE
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   6735
   End
End
Attribute VB_Name = "frmAddStud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowsePic_Click()
    dlgBrowse.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
    dlgBrowse.ShowOpen
    imgID.Picture = LoadPicture(dlgBrowse.FileName)
End Sub

Private Sub cmdCancel_Click()
Me.Hide
Call Status("Refreshing database. Please wait...")
AdoAddStud.Refresh
Call Status("Ready")
Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo DupErr
    
    AdoAddStud.Refresh
    
    Dim counter As Integer

    For counter = 0 To 5
        If Trim(txtBorInfo(counter).Text) = "" Then
            Missing
            txtBorInfo(counter).SetFocus
            Exit Sub
        End If
        
    Next counter
    
      
    If IsNumeric(txtBorInfo(5).Text) = False Then
        MsgBox "Cannot accept non-numeric input for contact number." & vbCrLf + vbCrLf & " Please change accordingly.", vbOKOnly + vbExclamation, "Library System"
        txtBorInfo(5).SetFocus
        SendKeys HiLyt
        Exit Sub
    End If
    
    If Trim(dlgBrowse.FileName) = "" Then
        MsgBox "Please provide an appropriate picture for identification.", vbOKOnly + vbExclamation, "Library System"
        Exit Sub
    End If
    
    'allocate a new recordset for the info
    AdoAddStud.Recordset.AddNew
    
    'enter the info to the database fields
    
        AdoAddStud.Recordset.Fields("BorId") = txtBorInfo(0).Text
        AdoAddStud.Recordset.Fields("LName") = txtBorInfo(3).Text
        AdoAddStud.Recordset.Fields("FName") = txtBorInfo(1).Text
        AdoAddStud.Recordset.Fields("MName") = txtBorInfo(2).Text
        AdoAddStud.Recordset.Fields("Course") = txtBorInfo(4).Text
        AdoAddStud.Recordset.Fields("Contact") = txtBorInfo(5).Text
        AdoAddStud.Recordset.Fields("Date_Reg") = Date
        
        AdoAddStud.Recordset.Fields("Pic") = dlgBrowse.FileName
        
        'update the database
        AdoAddStud.Recordset.Update
        MsgBox "New borrower entry successful!", vbOKOnly + vbInformation, "Library System"
        
        'clears the filename content ready for the next input
        dlgBrowse.FileName = ""
        
    'clears the ID Pic
    imgID.Picture = LoadPicture("")
    
    'clears contents of the fields
    For counter = 0 To 5
        txtBorInfo(counter).Text = ""
    Next counter
    
    'returns the focus to the first field
        txtBorInfo(0).SetFocus
    
    Exit Sub

DupErr:

MsgBox "The Borrower's ID you have entered already exists." & vbCrLf + vbCrLf & " Please specify a different ID.", vbOKOnly + vbExclamation, "Library System"
txtBorInfo(0).SetFocus
SendKeys HiLyt
End Sub

Private Sub Form_Load()
    Call Status("Loading. Please Wait...")
    On Error Resume Next
    Call ConnectToDb(AdoAddStud, "Borrower")
    
    txtBorInfo(0).SetFocus
    Call Status("Ready")
End Sub


Private Sub txtBorInfo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 And Index <> 5 Then
    txtBorInfo(Index + 1).SetFocus
    SendKeys HiLyt
End If

If KeyAscii = 13 And Index = 5 Then
cmdSave.SetFocus
End If

End Sub
