VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Reports"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab stbPrint 
      Height          =   2325
      Left            =   15
      TabIndex        =   0
      Top             =   870
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   4101
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      ForeColor       =   16744448
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Borrowers"
      TabPicture(0)   =   "frmPrint.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeBor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Books"
      TabPicture(1)   =   "frmPrint.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeBooks"
      Tab(1).ControlCount=   1
      Begin VB.Frame fmeBooks 
         Caption         =   "Registration Date Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   -74880
         TabIndex        =   8
         Top             =   660
         Width           =   4185
         Begin MSComCtl2.DTPicker dtpFrom2 
            Height          =   345
            Left            =   780
            TabIndex        =   9
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   11891757
            CalendarForeColor=   16777215
            CalendarTitleBackColor=   8208173
            CalendarTitleForeColor=   781309
            CalendarTrailingForeColor=   8421504
            Format          =   24379393
            CurrentDate     =   38065
         End
         Begin MSComCtl2.DTPicker dtpTo2 
            Height          =   345
            Left            =   2565
            TabIndex        =   10
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   11891757
            CalendarForeColor=   16777215
            CalendarTitleBackColor=   8208173
            CalendarTitleForeColor=   781309
            CalendarTrailingForeColor=   8421504
            Format          =   24379393
            CurrentDate     =   38065
         End
         Begin Project1.lvButtons_H cmdPrev 
            Height          =   405
            Index           =   1
            Left            =   2670
            TabIndex        =   13
            Top             =   975
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   714
            Caption         =   "&Preview"
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
            cBack           =   16777215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "FROM"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   555
            Width           =   465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "TO"
            Height          =   195
            Left            =   2235
            TabIndex        =   11
            Top             =   555
            Width           =   225
         End
      End
      Begin VB.Frame fmeBor 
         Caption         =   "Registration Date Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   4185
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   345
            Left            =   780
            TabIndex        =   4
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   11891757
            CalendarForeColor=   16777215
            CalendarTitleBackColor=   8208173
            CalendarTitleForeColor=   781309
            CalendarTrailingForeColor=   8421504
            Format          =   24379393
            CurrentDate     =   38065
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   345
            Left            =   2565
            TabIndex        =   7
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   11891757
            CalendarForeColor=   16777215
            CalendarTitleBackColor=   8208173
            CalendarTitleForeColor=   781309
            CalendarTrailingForeColor=   8421504
            Format          =   24379393
            CurrentDate     =   38065
         End
         Begin Project1.lvButtons_H cmdPrev 
            Height          =   405
            Index           =   0
            Left            =   2670
            TabIndex        =   14
            Top             =   975
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   714
            Caption         =   "&Preview"
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
            cBack           =   16777215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "TO"
            Height          =   195
            Left            =   2235
            TabIndex        =   6
            Top             =   555
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "FROM"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   555
            Width           =   465
         End
      End
   End
   Begin Project1.lvButtons_H cmdClose 
      Height          =   405
      Left            =   2805
      TabIndex        =   2
      Top             =   3315
      Width           =   1410
      _ExtentX        =   2487
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
      ImgAlign        =   1
      cBack           =   16777215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   15
      Picture         =   "frmPrint.frx":0038
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4425
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Print reports here. Choose report category below."
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   795
      TabIndex        =   1
      Top             =   330
      Width           =   3525
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   225
      Picture         =   "frmPrint.frx":29BF
      Top             =   195
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmPrint.frx":3689
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdPrev_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
   
    Case 0
        DataEnvironment1.rscomBorrowers.Open
        DataEnvironment1.rscomBorrowers.Filter = ""
        DataEnvironment1.rscomBorrowers.Filter = "Date_Reg >= '" & dtpFrom.Value & "' and Date_Reg <= '" & dtpTo.Value & "'"
        dtrBorrowers.Show vbModal, Me
        DataEnvironment1.rscomBorrowers.Close
    Case 1
        DataEnvironment1.rscomBooks.Open
        DataEnvironment1.rscomBooks.Filter = ""
        DataEnvironment1.rscomBooks.Filter = "DateReg >= '" & dtpFrom2.Value & "' and DateReg <= '" & dtpTo2.Value & "'"
        dtrBooks.Show vbModal, Me
        DataEnvironment1.rscomBooks.Close
    End Select
End Sub



Private Sub dtpTo_Change()
    If dtpTo.Value < dtpFrom.Value Then
            MsgBox "Ending date value cannot be less than Beginning date value. Adjust accordingly.", vbOKOnly + vbExclamation, "Library System"
            dtpTo.Value = dtpFrom.Value
            dtpTo.SetFocus
            Exit Sub
    End If
End Sub


Private Sub dtpTo2_Change()
    If dtpTo2.Value < dtpFrom2.Value Then
            MsgBox "Ending date value cannot be less than Beginning date value. Adjust accordingly.", vbOKOnly + vbExclamation, "Library System"
            dtpTo2.Value = dtpFrom2.Value
            dtpTo2.SetFocus
            Exit Sub
    End If
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub

Private Sub stbPrint_Click(PreviousTab As Integer)
    If stbPrint.Tab = 0 Then
        fmeBooks.Enabled = False
        fmeBor.Enabled = True
    Else
        fmeBooks.Enabled = True
        fmeBor.Enabled = False
    End If
End Sub

