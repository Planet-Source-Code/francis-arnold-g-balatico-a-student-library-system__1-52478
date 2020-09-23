VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStudProf 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library Transaction Panel"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   Icon            =   "frmStudProf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   5130
      Top             =   5625
   End
   Begin MSAdodcLib.Adodc AdoBooks 
      Height          =   375
      Left            =   2760
      Top             =   6285
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Borrow Type"
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
      Height          =   1965
      Left            =   5940
      TabIndex        =   19
      Top             =   3540
      Width           =   1650
      Begin VB.OptionButton optCirculation 
         BackColor       =   &H00875B25&
         Caption         =   "Loan"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   21
         Top             =   825
         Width           =   1245
      End
      Begin VB.OptionButton optCirculation 
         BackColor       =   &H00875B25&
         Caption         =   "Circulation"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   20
         Top             =   480
         Value           =   -1  'True
         Width           =   1245
      End
      Begin Project1.lvButtons_H cmdControl 
         Height          =   420
         Index           =   0
         Left            =   150
         TabIndex        =   22
         Top             =   1470
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   741
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
         cBhover         =   11891757
         cGradient       =   11891757
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   16777215
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000BEBFD&
         FillColor       =   &H00875B25&
         FillStyle       =   0  'Solid
         Height          =   900
         Left            =   165
         Top             =   375
         Width           =   1380
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5565
      Top             =   5625
   End
   Begin MSAdodcLib.Adodc ADOStud 
      Height          =   390
      Left            =   315
      Top             =   6270
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Search Borrower"
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
      Height          =   720
      Left            =   120
      TabIndex        =   14
      Top             =   885
      Width           =   7470
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
         Top             =   225
         Width           =   3600
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   435
         Left            =   6000
         TabIndex        =   2
         Top             =   180
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
         Image           =   "frmStudProf.frx":0CCA
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
         Top             =   285
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Borrower Information"
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
      Height          =   1875
      Left            =   120
      TabIndex        =   3
      Top             =   1635
      Width           =   7470
      Begin VB.PictureBox picStud 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00B5742D&
         Height          =   1530
         Left            =   135
         ScaleHeight     =   1470
         ScaleWidth      =   1470
         TabIndex        =   15
         Top             =   225
         Width           =   1530
         Begin VB.Image imgPicStud 
            Height          =   1500
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   1500
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Photo"
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
            Left            =   315
            TabIndex        =   16
            Top             =   630
            Width           =   810
         End
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   2805
         TabIndex        =   12
         Top             =   1395
         Width           =   75
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contact #:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   1950
         TabIndex        =   11
         Top             =   1425
         Width           =   750
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   2805
         TabIndex        =   9
         Top             =   1035
         Width           =   75
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   2160
         TabIndex        =   8
         Top             =   1065
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   2805
         TabIndex        =   7
         Top             =   660
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Borrower ID:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   1815
         TabIndex        =   6
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   2235
         TabIndex        =   5
         Top             =   345
         Width           =   465
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   2805
         TabIndex        =   4
         Top             =   315
         Width           =   75
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00875B25&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000BEBFD&
         FillColor       =   &H00875B25&
         Height          =   300
         Left            =   2730
         Top             =   285
         Width           =   4620
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00875B25&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000BEBFD&
         FillColor       =   &H00875B25&
         Height          =   300
         Left            =   2730
         Top             =   645
         Width           =   4620
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00875B25&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000BEBFD&
         FillColor       =   &H00875B25&
         Height          =   300
         Left            =   2730
         Top             =   1005
         Width           =   4620
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00875B25&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000BEBFD&
         FillColor       =   &H00875B25&
         Height          =   300
         Left            =   2730
         Top             =   1365
         Width           =   4620
      End
   End
   Begin MSAdodcLib.Adodc adoFilter 
      Height          =   390
      Left            =   1545
      Top             =   6270
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin Project1.lvButtons_H cmdControl 
      Height          =   420
      Index           =   2
      Left            =   6090
      TabIndex        =   17
      Top             =   5625
      Width           =   1380
      _ExtentX        =   2434
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
      ImgAlign        =   1
      cBack           =   16777215
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Borrowed Books"
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
      Height          =   1965
      Left            =   120
      TabIndex        =   10
      Top             =   3540
      Width           =   5760
      Begin MSDataGridLib.DataGrid dtgBooks 
         Height          =   1230
         Left            =   60
         TabIndex        =   13
         Top             =   210
         Visible         =   0   'False
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   2170
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   14811135
         Enabled         =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         RowDividerStyle =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin Project1.lvButtons_H cmdControl 
         Height          =   420
         Index           =   1
         Left            =   4320
         TabIndex        =   23
         Top             =   1470
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   741
         Caption         =   "&Return"
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
         Enabled         =   0   'False
         cBack           =   16777215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum allowed:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2415
         TabIndex        =   27
         Top             =   1575
         Width           =   1290
      End
      Begin VB.Label lblMax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   3795
         TabIndex        =   26
         Top             =   1575
         Width           =   120
      End
      Begin VB.Label lblNumBooks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1830
         TabIndex        =   25
         Top             =   1575
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total books borrowed:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   1575
         Width           =   1590
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00875B25&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000BEBFD&
         FillColor       =   &H00875B25&
         Height          =   360
         Left            =   60
         Top             =   1500
         Width           =   4110
      End
      Begin VB.Shape Shape2 
         Height          =   1230
         Left            =   60
         Top             =   210
         Width           =   5640
      End
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStudProf.frx":19A4
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   915
      TabIndex        =   18
      Top             =   135
      Width           =   6705
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   270
      Picture         =   "frmStudProf.frx":1A82
      Top             =   195
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmStudProf.frx":274C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7710
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   120
      Picture         =   "frmStudProf.frx":57EC
      Stretch         =   -1  'True
      Top             =   5550
      Width           =   7470
   End
End
Attribute VB_Name = "frmStudProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Bor_Opt As Integer
    'for borrowing type
    '0 = Circulation
    '1 = Loan



Private Sub cmdControl_Click(Index As Integer)
On Error GoTo Err
    If Index = 0 Then
        Load frmBorrow
        
        If optCirculation(0).Value = True Then
            If CheckCirculationDue() = True Then
                MsgBox "Borrower has past due book issued for circulation. Return book first.", vbOKOnly + vbExclamation, "Library System"
                Exit Sub
            End If
            frmBorrow.dtpDueDate.Value = Date
            frmBorrow.dtpDueDate.Enabled = False
        ElseIf optCirculation(1).Value = True Then
            frmBorrow.dtpDueDate.Value = Date
            frmBorrow.dtpDueDate.Enabled = True
        End If
                
        frmBorrow.Show vbModal, Me
        Call SetDataGrid
    ElseIf Index = 1 Then
        Call ReturnBooks
        Call SetDataGrid
        
    ElseIf Index = 2 Then
        Call DueCount
        Call OverDueCount
        Call BorrowedCount
        Unload Me
    End If
Exit Sub
Err:

End Sub

Private Sub cmdSearch_Click()
    Dim counter As Integer
    Timer1.Enabled = True
    ADOStud.Refresh
    
    dtgBooks.ClearFields
    
    On Error GoTo NotFound
        
        txtSearch.Text = Trim(txtSearch.Text)
        If Trim(txtSearch.Text) = "" Then
            Exit Sub
        End If
        
        Status "Searching..."
               
        'searches the current table for the username
        ADOStud.Recordset.Find ("BorId = '" & UCase(txtSearch.Text) & "'")
        
        lblInfo(1).Caption = ADOStud.Recordset.Fields("LName") & ", " & ADOStud.Recordset.Fields("FName") & " " & ADOStud.Recordset.Fields("MName") 'name labels
        lblInfo(2).Caption = ADOStud.Recordset.Fields("BorId") 'label for Borrower's ID
        lblInfo(3).Caption = ADOStud.Recordset.Fields("Course") 'label for course
        lblInfo(4).Caption = ADOStud.Recordset.Fields("Contact") 'label for contact number
        
        
            
        On Error Resume Next 'if there are no pictures to load or cannot be found, resume operation
        imgPicStud.Picture = LoadPicture(ADOStud.Recordset.Fields("Pic"))
    
        Call SetDataGrid 'Setups Data Grid
        dtgBooks.Visible = True
        
        
        If adoFilter.Recordset.RecordCount = 0 Then
            
            cmdControl(1).Enabled = False
        Else
            
            cmdControl(1).Enabled = True
        End If
            
        If adoFilter.Recordset.RecordCount >= MaxBooks Then
            cmdControl(0).Enabled = False
            optCirculation(0).Enabled = False
            optCirculation(1).Enabled = False
        Else
            cmdControl(0).Enabled = True
            optCirculation(0).Enabled = True
            optCirculation(1).Enabled = True
        End If
    
    Status "Ready"
    Exit Sub
    
    
NotFound:     'performs operation if no such record is found
    MsgBox "The ID you requested could not be found.", vbOKOnly + vbExclamation, "Library System"
        For counter = 1 To 4
            lblInfo(counter).Caption = ""
        Next counter
        Status "Ready"
    imgPicStud.Picture = LoadPicture("")
    cmdControl(0).Enabled = False
    cmdControl(1).Enabled = False
    dtgBooks.ClearFields
    txtSearch.SetFocus
    SendKeys HiLyt
End Sub

Private Sub Form_Load()
    Call Status("Loading Transaction Form. Please Wait.....")
    On Error GoTo ErrHandler
    'connects to the dbase
    Call ConnectToDb(ADOStud, "Borrower")
    Call ConnectToDb(AdoBooks, "Book")
     'Loads the database and provides the database password *Note:This database serves as a filter
        adoFilter.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Lib_Dbase.mdb;Persist Security Info=False; Jet OLEDB:Database Password = crimson119"
        
     'Sets the command type to Table
        adoFilter.CommandType = adCmdText
    
    ADOStud.Recordset.MoveFirst
    Call Status("Ready")
    Bor_Opt = 0
    Exit Sub
    
ErrHandler:
    Call NoRec(Me)
End Sub





Private Sub optCirculation_Click(Index As Integer)
    '0 = Circulation
    '1 = Loan
    Bor_Opt = Index
    
    If Index = 1 Then
        If adoFilter.Recordset.RecordCount > 0 Then
            MsgBox "To loan a book, all borrowed books must be returned to the library first.", vbOKOnly + vbInformation, "Library System"
            optCirculation(0).Value = True
        End If
    End If
    
End Sub

Private Sub Timer1_Timer() 'disables/enables the search button if criteria is unmet/met
        
    On Error Resume Next
    If adoFilter.Recordset.RecordCount = 0 Then
        cmdControl(1).Enabled = False
    Else
        cmdControl(1).Enabled = True
    End If
    
    If Bor_Opt = 0 Then
        If adoFilter.Recordset.RecordCount = MaxBooks Then
            cmdControl(0).Enabled = False
        Else
            cmdControl(0).Enabled = True
        End If
    ElseIf Bor_Opt = 1 Then
        If adoFilter.Recordset.RecordCount <> 0 Then
            cmdControl(0).Enabled = False
        Else
            cmdControl(0).Enabled = True
        End If
    End If
    
    lblMax.Caption = MaxBooks
    lblNumBooks.Caption = adoFilter.Recordset.RecordCount
End Sub

Private Sub Timer2_Timer()
    If Trim(txtSearch.Text) = "" Then
        cmdSearch.Enabled = False
    Else
        If IsNumeric(txtSearch.Text) = False Then
            cmdSearch.Enabled = False
        Else
            cmdSearch.Enabled = True
        End If
    End If
End Sub

Private Sub txtSearch_Change()
    Dim counter As Integer
    Timer1.Enabled = False
    cmdControl(0).Enabled = False
    cmdControl(1).Enabled = False
    adoFilter.RecordSource = ""
        For counter = 1 To 4
            lblInfo(counter).Caption = ""
        Next counter
        
        optCirculation(0).Enabled = False
        optCirculation(1).Enabled = False
        
        imgPicStud.Picture = LoadPicture("")
        
        dtgBooks.ClearFields
        dtgBooks.Visible = False
        
End Sub

Private Sub txtSearch_GotFocus()
    SendKeys HiLyt
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdSearch_Click
        txtSearch.SetFocus
        SendKeys HiLyt
    End If

End Sub

Private Sub SetDataGrid()
On Error Resume Next
 'filter the books to show only the ones borrowed by the Current Borrower
    adoFilter.RecordSource = "SELECT * FROM Book_Loan WHERE Borrower_Id = '" & Trim(txtSearch.Text) & "' and Status = 0 ORDER BY Date_Borrowed"
    adoFilter.Refresh
    
    Set dtgBooks.DataSource = adoFilter
    
    'ensures that the column formatting remains the same
    With dtgBooks
        dtgBooks.Columns(0).DataField = "Book_Id"
        dtgBooks.Columns(0).Caption = "Book ID"
        dtgBooks.Columns(0).Width = 1500
        
        dtgBooks.Columns(1).DataField = "Call_ID"
        dtgBooks.Columns(1).Caption = "Call ID"
        dtgBooks.Columns(1).Width = 1500
        
        dtgBooks.Columns(2).DataField = "Book_Title"
        dtgBooks.Columns(2).Caption = "Title"
        dtgBooks.Columns(2).Width = 3000
        
        dtgBooks.Columns(3).DataField = "Borrower_Id"
        dtgBooks.Columns(3).Caption = "Borrower ID"
        dtgBooks.Columns(3).Width = 0
        
        dtgBooks.Columns(4).DataField = "Borrower_FName"
        dtgBooks.Columns(4).Caption = "First Name"
        dtgBooks.Columns(4).Width = 0
        
        dtgBooks.Columns(5).DataField = "Borrower_LName"
        dtgBooks.Columns(5).Caption = "Last Name"
        dtgBooks.Columns(5).Width = 0
        
        dtgBooks.Columns(6).DataField = "Date_Borrowed"
        dtgBooks.Columns(6).Caption = "Date Borrowed"
        dtgBooks.Columns(6).Width = 1500
        
        dtgBooks.Columns(7).DataField = "Date_Due"
        dtgBooks.Columns(7).Caption = "Due Date"
        dtgBooks.Columns(7).Width = 1500
        
        dtgBooks.Columns(8).DataField = "Date_Returned"
        dtgBooks.Columns(8).Caption = "Date Returned"
        dtgBooks.Columns(8).Width = 0
        
        dtgBooks.Columns(9).DataField = "Days_Past_Due"
        dtgBooks.Columns(9).Caption = "Days Past Due"
        dtgBooks.Columns(9).Width = 0
               
        dtgBooks.Columns(10).DataField = "Fines"
        dtgBooks.Columns(10).Caption = "Fine"
        dtgBooks.Columns(10).Width = 0
        
        dtgBooks.Columns(11).DataField = "Status"
        dtgBooks.Columns(11).Caption = "Status"
        dtgBooks.Columns(11).Width = 0
    End With
    adoFilter.Recordset.Sort = "Book_Id ASC"
End Sub

Private Sub ReturnBooks()

Dim tmpWord As String
Dim tmpDaysPast As Integer

On Error Resume Next
    Status "Returning book to library..."
    AdoBooks.Refresh
    AdoBooks.Recordset.Find ("BookId = '" & Trim(adoFilter.Recordset.Fields("Book_Id")) & "'")
    AdoBooks.Recordset.Fields("StatusID") = 1
    AdoBooks.Recordset.Update
    
    With adoFilter.Recordset
        .Fields("Date_Returned") = Date
        
        If DateDiff("d", .Fields("Date_Due"), Date) <= 0 Then
            tmpDaysPast = 0
            .Fields("Days_Past_Due") = tmpDaysPast
        Else
            tmpDaysPast = DateDiff("d", .Fields("Date_Due"), Date)
            .Fields("Days_Past_Due") = tmpDaysPast
        End If
        
        .Fields("Fines") = tmpDaysPast * Fines
        .Fields("Status") = 1
        .Update
    End With
    

    
If tmpDaysPast > 1 Then
    tmpWord = "days"
Else
    tmpWord = "day"
End If

    If tmpDaysPast > 0 Then
        MsgBox tmpDaysPast & " " & tmpWord & " overdue. Collect " & (DateDiff("d", adoFilter.Recordset.Fields("Date_Due"), Date)) * Fines & " from Borrower.", vbOKOnly + vbInformation, "Library System"
    End If
    
    adoFilter.Refresh
    
    cmdControl(0).Enabled = True
    optCirculation(0).Enabled = True
    optCirculation(1).Enabled = True
    
    Status "Ready"
End Sub

Private Function CheckCirculationDue() As Boolean
Dim Strdate As String
Strdate = Date
CheckCirculationDue = False
    If Not adoFilter.Recordset.RecordCount = 0 Then
        adoFilter.Recordset.MoveFirst
        Do While adoFilter.Recordset.EOF = False
            If Not adoFilter.Recordset.Fields("Date_Borrowed") = Strdate Then
                CheckCirculationDue = True
                Exit Function
            Else
                CheckCirculationDue = False
            End If
            adoFilter.Recordset.MoveNext
        Loop
    End If
End Function

