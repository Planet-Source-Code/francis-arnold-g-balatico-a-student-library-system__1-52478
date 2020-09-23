VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEditBor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Borrower Profile"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6840
   ControlBox      =   0   'False
   Icon            =   "frmEditBor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   2040
      Top             =   4860
   End
   Begin MSComDlg.CommonDialog dlgPic 
      Left            =   1590
      Top             =   4830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoEdit 
      Height          =   390
      Left            =   270
      Top             =   4905
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
      Caption         =   "BORROWER INFORMATION"
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
      Height          =   2985
      Left            =   60
      TabIndex        =   18
      Top             =   1770
      Width           =   6690
      Begin VB.PictureBox picPrev 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00B5742D&
         Height          =   1530
         Left            =   180
         ScaleHeight     =   1470
         ScaleWidth      =   1470
         TabIndex        =   19
         Top             =   375
         Width           =   1530
         Begin VB.Image imgPrev 
            Height          =   1500
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   1500
         End
         Begin VB.Label Label2 
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
            TabIndex        =   20
            Top             =   660
            Width           =   1110
         End
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
         Index           =   5
         Left            =   3120
         TabIndex        =   12
         Top             =   2460
         Width           =   3435
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
         Left            =   3120
         TabIndex        =   10
         Top             =   1935
         Width           =   3435
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
         Left            =   3120
         TabIndex        =   8
         Top             =   1395
         Width           =   3435
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
         Left            =   3120
         TabIndex        =   6
         Top             =   870
         Width           =   3435
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
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   345
         Width           =   3435
      End
      Begin Project1.lvButtons_H cmdBrowsePic 
         Height          =   540
         Left            =   180
         TabIndex        =   13
         Top             =   1965
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   953
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
         Image           =   "frmEditBor.frx":0CCA
         ImgSize         =   24
         Enabled         =   0   'False
         cBack           =   16777215
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
         Left            =   2145
         TabIndex        =   11
         Top             =   2535
         Width           =   915
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&ourse:"
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
         TabIndex        =   9
         Top             =   2010
         Width           =   660
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
         Left            =   2085
         TabIndex        =   7
         Top             =   1470
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
         Left            =   1905
         TabIndex        =   5
         Top             =   960
         Width           =   1170
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
         Left            =   2100
         TabIndex        =   3
         Top             =   435
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "SEARCH BORROWER"
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
      Height          =   840
      Left            =   60
      TabIndex        =   17
      Top             =   900
      Width           =   6690
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
         Left            =   2190
         TabIndex        =   1
         Top             =   285
         Width           =   2895
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   435
         Left            =   5205
         TabIndex        =   2
         Top             =   255
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   767
         Caption         =   "&Search"
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
         ImgAlign        =   1
         Image           =   "frmEditBor.frx":19A4
         cBack           =   16777215
      End
      Begin VB.Label lblBorId 
         AutoSize        =   -1  'True
         Caption         =   "&Enter Borrower ID"
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
         Left            =   285
         TabIndex        =   0
         Top             =   375
         Width           =   1830
      End
   End
   Begin Project1.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   5355
      TabIndex        =   16
      Top             =   4875
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
      Left            =   3960
      TabIndex        =   14
      Top             =   4875
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
      Left            =   2550
      TabIndex        =   15
      Top             =   4875
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
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   135
      Picture         =   "frmEditBor.frx":267E
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEditBor.frx":3348
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   720
      TabIndex        =   21
      Top             =   135
      Width           =   6000
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmEditBor.frx":3443
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6840
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   45
      Picture         =   "frmEditBor.frx":64E3
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   6690
   End
End
Attribute VB_Name = "frmEditBor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowsePic_Click() 'browse the borrower's picture
    
    dlgPic.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
    dlgPic.ShowOpen
    
    If dlgPic.FileName <> "" Then
        imgPrev.Picture = LoadPicture(dlgPic.FileName)
    End If
    
End Sub

Private Sub cmdCancel_Click()

Unload Me

End Sub

'reloads the unsaved info
Private Sub cmdReload_Click()
    Dim counter As Integer
    
    If MsgBox("This will reload the current data. Any unsaved data will be lost. Proceed?", vbYesNo + vbQuestion, "Library System") = vbNo Then
    Exit Sub
    Else
    
    imgPrev.Picture = LoadPicture("") 'clears picture
    dlgPic.FileName = "" 'clears the filename property
    
    'refresh dbase
    AdoEdit.Refresh
    
    'finds the record to edit
    AdoEdit.Recordset.Find ("BorId = '" & UCase(txtSearch.Text) & "'")
    
    
        For counter = 1 To 5
            txtInfo(counter).Text = AdoEdit.Recordset.Fields(counter)
        Next counter
    
    On Error Resume Next
        imgPrev.Picture = LoadPicture(AdoEdit.Recordset.Fields(6))
        
        'enables the textfields and pictures and command buttons
        For counter = 1 To 5
            txtInfo(counter).Enabled = True
        Next counter
                
        cmdBrowsePic.Enabled = True
        
        cmdSave.Enabled = True
        cmdReload.Enabled = True
        
        txtInfo(1).SetFocus
        SendKeys HiLyt
    End If
End Sub

Private Sub cmdSave_Click() 'updates the current record
    Dim counter As Integer
    
    On Error GoTo Err
    
    For counter = 1 To 5 'validates fields
        If txtInfo(counter).Text = "" Then
            MsgBox "Required field missing. Please fill ALL fields", vbOKOnly + vbExclamation, "Library System"
            txtInfo(counter).SetFocus
            Exit Sub
        End If
    Next counter
    
    If imgPrev.Picture = 0 Then 'checks existence of picture
        MsgBox "Please provide a photo for identification", vbOKOnly + vbExclamation, "Library System"
        Exit Sub
    End If
    
    
    If MsgBox("Save changes to the current record?", vbOKCancel + vbQuestion, "Library System") = vbCancel Then
        Exit Sub
    Else
    
        
        AdoEdit.Recordset.Fields("Fname") = txtInfo(1).Text
        AdoEdit.Recordset.Fields("Mname") = txtInfo(2).Text
        AdoEdit.Recordset.Fields("Lname") = txtInfo(3).Text
        AdoEdit.Recordset.Fields("Course") = txtInfo(4).Text
        AdoEdit.Recordset.Fields("Contact") = txtInfo(5).Text
        
        If dlgPic.FileName <> "" Then
            AdoEdit.Recordset.Fields("Pic") = dlgPic.FileName
        End If
        
        AdoEdit.Recordset.Update
        Call Status("Updating record. Please wait...")
        AdoEdit.Refresh
        Call Status("Ready")
            
        If MsgBox("Record successfully updated. Continue editing records?", vbYesNo + vbQuestion, "Library system") = vbYes Then
            txtSearch.SetFocus
            SendKeys HiLyt
            Exit Sub
        Else
            Unload Me
        End If
           
    End If
    Exit Sub
    
Err:
     
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler
    
        txtSearch.Text = Trim(txtSearch.Text)
        If Trim(txtSearch.Text) = "" Then
            Exit Sub
        End If
    
    Dim counter As Byte
    
    'refresh dbase
    AdoEdit.Refresh
    
    'finds the record to edit
    AdoEdit.Recordset.Find ("BorId = '" & UCase(txtSearch.Text) & "'")
    
            txtInfo(1).Text = AdoEdit.Recordset.Fields("Fname")
            txtInfo(2).Text = AdoEdit.Recordset.Fields("Mname")
            txtInfo(3).Text = AdoEdit.Recordset.Fields("Lname")
            txtInfo(4).Text = AdoEdit.Recordset.Fields("Course")
            txtInfo(5).Text = AdoEdit.Recordset.Fields("Contact")
    
    On Error Resume Next
        imgPrev.Picture = LoadPicture(AdoEdit.Recordset.Fields("Pic"))
        
        'enables the textfields and pictures and command buttons
        For counter = 1 To 5
            txtInfo(counter).Enabled = True
            txtInfo(counter).BackColor = vbWhite
        Next counter
                
        cmdBrowsePic.Enabled = True
        
        cmdSave.Enabled = True
        cmdReload.Enabled = True
        
        txtInfo(1).SetFocus
        SendKeys HiLyt
        
    Exit Sub
    
ErrHandler:
    'states the error
    MsgBox "The record you requested could not be found.", vbOKOnly + vbExclamation, "Library System"
    
    'clears the textboxes
    For counter = 1 To 5
        txtInfo(counter).Text = ""
    Next counter
    
    'clears the picture
    imgPrev.Picture = LoadPicture("")
    
    'disables the textfields and pictures and command buttons
        For counter = 1 To 5
            txtInfo(counter).Enabled = False
            txtInfo(counter).BackColor = &HE0E0E0
        Next counter
                
        cmdBrowsePic.Enabled = False
        
        cmdSave.Enabled = False
        cmdReload.Enabled = False
    
    'sends the focus to the search box
    txtSearch.SetFocus
    SendKeys HiLyt
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
        
        Call Status("Loading student information...")
        Call ConnectToDb(AdoEdit, "Borrower")
        AdoEdit.Refresh
        
        AdoEdit.Recordset.MoveFirst
        
        Call Status("Ready")
        
        dlgPic.FileName = ""
        Exit Sub
    
ErrHandler:
       Call NoRec(Me)

End Sub




Private Sub Timer1_Timer()
    If Trim(txtSearch.Text) = "" Then
        cmdSearch.Enabled = False
    Else
        cmdSearch.Enabled = True
    End If
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    'automatic shifting of focus when the enter key is pressed
    If KeyAscii = 13 Then
        If Index <> 5 Then
            txtInfo(Index + 1).SetFocus
            SendKeys HiLyt
        Else
            cmdBrowsePic.SetFocus
        End If
    End If

End Sub

Private Sub txtSearch_Change()
    Dim counter As Byte
    
    For counter = 1 To 5
        txtInfo(counter).Text = ""
    Next counter
    
    'clears the picture
    imgPrev.Picture = LoadPicture("")
    
    'disables the textfields and pictures and command buttons
        For counter = 1 To 5
            txtInfo(counter).Enabled = False
            txtInfo(counter).BackColor = &HE0E0E0
        Next counter
                
        cmdBrowsePic.Enabled = False
        
        cmdSave.Enabled = False
        cmdReload.Enabled = False

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
End Sub

