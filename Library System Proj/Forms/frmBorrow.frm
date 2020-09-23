VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBorrow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrow Books"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   2655
      Top             =   3720
   End
   Begin MSAdodcLib.Adodc AdoTitle 
      Height          =   330
      Left            =   1410
      Top             =   3765
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoBook 
      Height          =   330
      Left            =   165
      Top             =   3765
      Visible         =   0   'False
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Book Details"
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
      Height          =   1560
      Left            =   45
      TabIndex        =   8
      Top             =   1500
      Width           =   5535
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00875B25&
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   1050
         ScaleHeight     =   1125
         ScaleWidth      =   4350
         TabIndex        =   13
         Top             =   285
         Width           =   4380
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   4
            Left            =   75
            TabIndex        =   17
            Top             =   855
            Width           =   120
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   75
            TabIndex        =   16
            Top             =   600
            Width           =   120
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   75
            TabIndex        =   15
            Top             =   330
            Width           =   120
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   75
            TabIndex        =   14
            Top             =   75
            Width           =   120
         End
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   375
         Width           =   450
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   420
         TabIndex        =   11
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edition:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   390
         TabIndex        =   10
         Top             =   900
         Width           =   660
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   1155
         Width           =   855
      End
   End
   Begin Project1.lvButtons_H cmdClose 
      Height          =   405
      Left            =   4170
      TabIndex        =   6
      Top             =   3720
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   45
      TabIndex        =   18
      Top             =   2985
      Width           =   5535
      Begin MSComCtl2.DTPicker dtpDueDate 
         Height          =   345
         Left            =   1080
         TabIndex        =   4
         Top             =   195
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   11891757
         CalendarForeColor=   16777215
         CalendarTitleBackColor=   8208173
         CalendarTitleForeColor=   781309
         CalendarTrailingForeColor=   8421504
         Format          =   24444929
         CurrentDate     =   38061
      End
      Begin Project1.lvButtons_H cmdBorrow 
         Height          =   405
         Left            =   4125
         TabIndex        =   5
         Top             =   165
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   714
         Caption         =   "&Borrow"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Due Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   285
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   60
      TabIndex        =   19
      Top             =   840
      Width           =   5550
      Begin VB.TextBox txtBookID 
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
         Left            =   1410
         TabIndex        =   1
         Top             =   195
         Width           =   3375
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   435
         Left            =   4860
         TabIndex        =   2
         Top             =   150
         Width           =   570
         _ExtentX        =   1005
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
         Image           =   "frmBorrow.frx":0000
         cBack           =   16777215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Book Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   150
         TabIndex        =   0
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   45
      Picture         =   "frmBorrow.frx":0CDA
      Stretch         =   -1  'True
      Top             =   3645
      Width           =   5535
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "frmBorrow.frx":3661
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the book number of the book to be borrowed."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   360
      Width           =   3630
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmBorrow.frx":432B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5625
   End
End
Attribute VB_Name = "frmBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBorrow_Click()
If frmStudProf.adoFilter.Recordset.RecordCount = MaxBooks Then
    MsgBox "Maximum allowed books to be borrowed has been reached.", vbOKOnly + vbInformation, "Library System"
    Exit Sub
End If

    frmStudProf.adoFilter.Refresh
    With frmStudProf.adoFilter.Recordset
        .AddNew
        .Fields("Book_Id") = txtBookID.Text
        .Fields("Call_ID") = AdoTitle.Recordset.Fields("CallID")
        .Fields("Book_Title") = AdoTitle.Recordset.Fields("Title")
        .Fields("Borrower_Id") = frmStudProf.ADOStud.Recordset.Fields("BorId")
        .Fields("Borrower_FName") = frmStudProf.ADOStud.Recordset.Fields("FName")
        .Fields("Borrower_LName") = frmStudProf.ADOStud.Recordset.Fields("LName")
        .Fields("Date_Borrowed") = Date
        
        If frmStudProf.optCirculation(0).Value = True Then
            .Fields("Date_Due") = Date
        ElseIf frmStudProf.optCirculation(1).Value = True Then
            .Fields("Date_Due") = dtpDueDate.Value
        End If
                
        .Fields("Status") = 0
        .Update
    End With
    
        AdoBook.Recordset.Fields("StatusID") = 2
        AdoBook.Recordset.Update
       
        Call ClearMe
        txtBookID.SetFocus
        SendKeys HiLyt
        
End Sub

Private Sub cmdClose_Click()
    With frmStudProf
        If .adoFilter.Recordset.RecordCount >= MaxBooks Then
            .cmdControl(0).Enabled = False
            .optCirculation(0).Enabled = False
            .optCirculation(1).Enabled = False
        Else
            .cmdControl(0).Enabled = True
            .optCirculation(0).Enabled = True
            .optCirculation(1).Enabled = True
        End If
    End With
    
    Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error GoTo NotFound
    
    Status "Searching..."
    AdoBook.Refresh
    AdoTitle.Refresh
    
    AdoBook.Recordset.Find ("BookID = '" & Trim(txtBookID.Text) & "'")
    
    If AdoBook.Recordset.Fields("StatusID") = 2 Then 'Checks if book is already borrowed
        MsgBox "Book already borrowed. Please specify a different book.", vbOKOnly + vbInformation, "Library System"
        txtBookID.SetFocus
        SendKeys HiLyt
        Exit Sub
    End If
    
    AdoTitle.Recordset.Find ("CallID = '" & Trim(AdoBook.Recordset.Fields("CallID")) & "'")
    
    On Error Resume Next
    
    If Trim(AdoTitle.Recordset.Fields("Title")) = "" Then
        lblInfo(1).Caption = "--"
    Else
        lblInfo(1).Caption = AdoTitle.Recordset.Fields("Title")
    End If
    
    If Trim(AdoTitle.Recordset.Fields("Author")) = "" Then
        lblInfo(2).Caption = "--"
    Else
        lblInfo(2).Caption = AdoTitle.Recordset.Fields("Author")
    End If
    
    If Trim(AdoBook.Recordset.Fields("Edition")) = "" Then
        lblInfo(3).Caption = "--"
    Else
        lblInfo(3).Caption = AdoBook.Recordset.Fields("Edition")
    End If
    
    If Trim(AdoBook.Recordset.Fields("Publisher")) = "" Then
        lblInfo(4).Caption = "--"
    Else
        lblInfo(4).Caption = AdoBook.Recordset.Fields("Publisher")
    End If
    
    cmdBorrow.Enabled = True
    cmdBorrow.SetFocus
    Status "Ready"
Exit Sub

NotFound:
    MsgBox "The book specified could not be found. Please enter a different book number.", vbOKOnly + vbExclamation, "Library System"
    txtBookID.SetFocus
    SendKeys HiLyt
    Status "Ready"
End Sub



Private Sub dtpDueDate_Change()
    If dtpDueDate.Value < Date Then
        MsgBox "Due date must not be earlier than current date.", vbOKOnly + vbExclamation, "Library System"
        dtpDueDate.Value = Date
    End If
End Sub



Private Sub Form_Load()
   Status "Loading..."
   Call ConnectToDb(AdoBook, "Book")
   Call ConnectToDb(AdoTitle, "Title")
   If AdoBook.Recordset.RecordCount = 0 Then
        MsgBox "There are no registered books to be borrowed. Register books first.", vbOKOnly + vbExclamation, "Library System"
        Unload Me
   End If
   Status "Ready"
End Sub

Private Sub ClearMe()
    Dim counter As Byte
    
    For counter = 1 To 4
        lblInfo(counter).Caption = "--"
    Next counter
    
    cmdBorrow.Enabled = False
End Sub


Private Sub Timer1_Timer()
    If Trim(txtBookID.Text) = "" Then
        cmdSearch.Enabled = False
    Else
        cmdSearch.Enabled = True
    End If
    
    
End Sub

Private Sub txtBookID_Change()
    ClearMe
End Sub

Private Sub txtBookID_GotFocus()
    SendKeys HiLyt
End Sub

Private Sub txtBookID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
End Sub
