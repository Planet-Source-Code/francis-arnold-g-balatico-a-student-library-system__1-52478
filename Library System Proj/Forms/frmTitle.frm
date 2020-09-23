VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTitle 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Title Manager"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4515
      Top             =   5100
   End
   Begin MSAdodcLib.Adodc AdoTitle 
      Height          =   375
      Left            =   135
      Top             =   5130
      Visible         =   0   'False
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
   Begin TabDlg.SSTab tabTitle 
      Height          =   3495
      Left            =   30
      TabIndex        =   21
      Top             =   945
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   6165
      _Version        =   393216
      TabOrientation  =   1
      Tab             =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Add Title"
      TabPicture(0)   =   "frmTitle.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Edit Title"
      TabPicture(1)   =   "frmTitle.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Delete Title"
      TabPicture(2)   =   "frmTitle.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fmeDel"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdSearchDel"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtIdDel"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtIdDel 
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
         Height          =   420
         Left            =   1380
         TabIndex        =   24
         Top             =   150
         Width           =   3630
      End
      Begin VB.Frame Frame2 
         Caption         =   "Title Information"
         Height          =   3000
         Left            =   -74940
         TabIndex        =   23
         Top             =   65
         Width           =   6540
         Begin VB.TextBox txtEdCallID 
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
            Left            =   1335
            TabIndex        =   11
            Top             =   390
            Width           =   3630
         End
         Begin VB.TextBox txtEdAuthor 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
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
            Left            =   1335
            TabIndex        =   18
            Top             =   2265
            Width           =   3630
         End
         Begin VB.TextBox txtEdISBN 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
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
            Left            =   1335
            TabIndex        =   16
            Top             =   1740
            Width           =   3630
         End
         Begin VB.TextBox txtEdTitle 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
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
            Left            =   1350
            TabIndex        =   14
            Top             =   1215
            Width           =   3630
         End
         Begin Project1.lvButtons_H cmdUpdate 
            Height          =   405
            Left            =   5145
            TabIndex        =   19
            Top             =   2235
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
            Enabled         =   0   'False
            cBack           =   16777215
         End
         Begin Project1.lvButtons_H cmdSearch 
            Height          =   405
            Left            =   5145
            TabIndex        =   12
            Top             =   375
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   714
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
            Image           =   "frmTitle.frx":0054
            cBack           =   16777215
         End
         Begin VB.Line Line1 
            X1              =   30
            X2              =   6495
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Call Number:"
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
            TabIndex        =   10
            Top             =   510
            Width           =   1095
         End
         Begin VB.Label Label8 
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
            Left            =   630
            TabIndex        =   17
            Top             =   2385
            Width           =   630
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN:"
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
            Left            =   750
            TabIndex        =   15
            Top             =   1845
            Width           =   510
         End
         Begin VB.Label Label6 
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
            Left            =   810
            TabIndex        =   13
            Top             =   1335
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Title Information"
         Height          =   3000
         Left            =   -74940
         TabIndex        =   22
         Top             =   65
         Width           =   6540
         Begin VB.TextBox txtCallID 
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
            Left            =   1335
            TabIndex        =   1
            Top             =   435
            Width           =   3630
         End
         Begin VB.TextBox txtAuthor 
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
            Left            =   1335
            TabIndex        =   7
            Top             =   2010
            Width           =   3630
         End
         Begin VB.TextBox txtISBN 
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
            Left            =   1335
            TabIndex        =   5
            Top             =   1485
            Width           =   3630
         End
         Begin VB.TextBox txtTitle 
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
            Left            =   1335
            TabIndex        =   3
            Top             =   960
            Width           =   3630
         End
         Begin Project1.lvButtons_H cmdAdd 
            Height          =   405
            Left            =   5145
            TabIndex        =   8
            Top             =   435
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   714
            Caption         =   "&Add"
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
         Begin Project1.lvButtons_H cmdClear 
            Height          =   405
            Left            =   5145
            TabIndex        =   9
            Top             =   1965
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   714
            Caption         =   "C&lear"
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
            Enabled         =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Call Number:"
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
            TabIndex        =   0
            Top             =   555
            Width           =   1095
         End
         Begin VB.Label Label4 
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
            Left            =   630
            TabIndex        =   6
            Top             =   2130
            Width           =   630
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ISBN:"
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
            Left            =   750
            TabIndex        =   4
            Top             =   1590
            Width           =   510
         End
         Begin VB.Label Label2 
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
            Left            =   810
            TabIndex        =   2
            Top             =   1080
            Width           =   450
         End
      End
      Begin Project1.lvButtons_H cmdSearchDel 
         Height          =   405
         Left            =   5160
         TabIndex        =   25
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   714
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
         Image           =   "frmTitle.frx":0D2E
         cBack           =   16777215
      End
      Begin VB.Frame fmeDel 
         Caption         =   "Books under this Title"
         Height          =   2445
         Left            =   60
         TabIndex        =   27
         Top             =   645
         Width           =   6540
         Begin MSDataGridLib.DataGrid dtgBooks 
            Bindings        =   "frmTitle.frx":1A08
            Height          =   1080
            Left            =   75
            TabIndex        =   29
            Top             =   840
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   1905
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   14585129
            ForeColor       =   16777215
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   1
            FormatLocked    =   -1  'True
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
               Weight          =   700
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
               ScrollBars      =   2
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  DividerStyle    =   1
                  ColumnWidth     =   5804.788
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin Project1.lvButtons_H cmdDel 
            Height          =   405
            Left            =   5130
            TabIndex        =   28
            Top             =   1950
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
            Image           =   "frmTitle.frx":1A1F
            Enabled         =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label lblAuthor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00875B25&
            Height          =   195
            Left            =   900
            TabIndex        =   36
            Top             =   540
            Width           =   135
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00875B25&
            Height          =   195
            Left            =   900
            TabIndex        =   35
            Top             =   270
            Width           =   135
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
            Left            =   165
            TabIndex        =   34
            Top             =   540
            Width           =   630
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
            Left            =   345
            TabIndex        =   33
            Top             =   270
            Width           =   450
         End
         Begin VB.Label lblBooks 
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
            ForeColor       =   &H00875B25&
            Height          =   195
            Left            =   2430
            TabIndex        =   32
            Top             =   2055
            Width           =   120
         End
         Begin VB.Label lblNumBooks 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "# of books under this title:"
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
            TabIndex        =   31
            Top             =   2055
            Width           =   2265
         End
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Call Number:"
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
         Left            =   210
         TabIndex        =   26
         Top             =   270
         Width           =   1095
      End
   End
   Begin Project1.lvButtons_H cmdClose 
      Height          =   405
      Left            =   5220
      TabIndex        =   20
      Top             =   4560
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
   Begin MSAdodcLib.Adodc AdoBooks 
      Height          =   375
      Left            =   1350
      Top             =   5130
      Visible         =   0   'False
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
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   315
      Picture         =   "frmTitle.frx":26F9
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTitle.frx":33C3
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1005
      TabIndex        =   30
      Top             =   240
      Width           =   5625
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmTitle.frx":3453
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6705
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   30
      Picture         =   "frmTitle.frx":64F3
      Stretch         =   -1  'True
      Top             =   4485
      Width           =   6660
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ID As String

Private Sub cmdAdd_Click()
On Error GoTo Duplicate_Err
    'field validation
    
    Status "Validating fields..."
    If Trim(txtCallID.Text) = "" Then
        Missing
        txtCallID.SetFocus
        Exit Sub
    ElseIf Trim(txtTitle.Text) = "" Then
        Missing
        txtTitle.SetFocus
        Exit Sub
    ElseIf Trim(txtISBN.Text) = "" Then
        Missing
        txtISBN.SetFocus
        Exit Sub
    ElseIf Trim(txtAuthor.Text) = "" Then
        Missing
        txtAuthor.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtCallID.Text) = False Then
        MsgBox "Cannot acept non-numeric input for Call Number.", vbOKOnly + vbExclamation, "Library System"
        txtCallID.SetFocus
        SendKeys HiLyt
        Exit Sub
    End If
    
    
    'setup database for data transfer and update
    Call Status("Adding new title...")
    AdoTitle.Refresh
    AdoTitle.Recordset.AddNew
    
    'transfer data
    AdoTitle.Recordset.Fields("CallID") = txtCallID.Text
    AdoTitle.Recordset.Fields("Title") = txtTitle.Text
    AdoTitle.Recordset.Fields("ISBN") = txtISBN.Text
    AdoTitle.Recordset.Fields("Author") = txtAuthor.Text
    
    'save transferred data
    AdoTitle.Recordset.Update
    
    'validate save
    MsgBox "New title added.", vbInformation + vbOKOnly, "Library System"
    
        'clear fields ready for new input
        txtCallID.Text = ""
        txtTitle.Text = ""
        txtISBN.Text = ""
        txtAuthor.Text = ""
        txtCallID.SetFocus
    Call Status("Ready")
Exit Sub

Duplicate_Err:
    MsgBox "The call ID specified already exists. Please specify a different call ID.", vbExclamation + vbOKOnly, "Library System"
    txtCallID.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

Private Sub cmdClear_Click()
    If MsgBox("Clear all fields?", vbYesNo + vbQuestion, "Library System") = vbYes Then
        txtCallID.Text = ""
        txtTitle.Text = ""
        txtISBN.Text = ""
        txtAuthor.Text = ""
        txtCallID.SetFocus
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    If MsgBox("This will delete '" & AdoTitle.Recordset.Fields("Title") & "' and all records associated with it." & vbCrLf & "Proceed?", vbQuestion + vbYesNo, "Library System") = vbYes Then
        Call Status("Deleting...")
        AdoTitle.Recordset.Delete
        AdoTitle.Refresh
        MsgBox "Title deleted.", vbOKCancel + vbInformation, "Library System"
        txtIdDel.Text = ""
        txtIdDel.SetFocus
        Call Status("Ready")
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo NotFound
    
        If AdoTitle.Recordset.RecordCount = 0 Then
            MsgBox "There are no existing titles to search.", vbOKOnly + vbExclamation, "Library System"
            Exit Sub
        End If
    
    AdoTitle.Refresh
    Call Status("Searching...")
    AdoTitle.Recordset.Find ("CallID = '" & txtEdCallID.Text & "'")
    
    'assign values
    txtEdTitle.Text = AdoTitle.Recordset.Fields("Title")
    txtEdISBN.Text = AdoTitle.Recordset.Fields("ISBN")
    txtEdAuthor.Text = AdoTitle.Recordset.Fields("Author")
    
    'enable textboxes
    txtEdTitle.Enabled = True
    txtEdTitle.BackColor = vbWhite
    
    txtEdISBN.Enabled = True
    txtEdISBN.BackColor = vbWhite
    
    txtEdAuthor.Enabled = True
    txtEdAuthor.BackColor = vbWhite
    
    
    cmdUpdate.Enabled = True
    
    txtEdTitle.SetFocus
    SendKeys HiLyt
    
    Call Status("Ready")
    Exit Sub

NotFound:
    MsgBox "Call ID not found. Please specify an existing call ID.", vbOKOnly + vbExclamation, "Library System"
    Call DisAbler
    txtEdCallID.SetFocus
    SendKeys HiLyt
End Sub

Private Sub cmdSearchDel_Click()

On Error GoTo NotFound

    If AdoTitle.Recordset.RecordCount = 0 Then 'prompts when no records are available for title
        MsgBox "There are no existing titles to search.", vbOKOnly + vbExclamation, "Library System"
        Exit Sub
    End If

    AdoTitle.Refresh
    Call Status("Searching...")
    AdoTitle.Recordset.Find ("CallID = '" & txtIdDel.Text & "'")

    'assigns Call ID to memory
    ID = AdoTitle.Recordset.Fields("CallID")
    
      
    'set SQL Statement
    Call SQLDB(AdoBooks, "SELECT Book.BookId, Book.CallId FROM Book WHERE (((Book.CallId)='" & ID & "'));")
    AdoBooks.Refresh
    
    'format data grid
    With dtgBooks
    .Refresh
    .Columns(0).Caption = "Book Number"
    .Columns(0).DataField = "BookId"
    .Columns(0).Width = 5800
    
    .Columns(1).Caption = "Call ID"
    .Columns(1).DataField = "CallId"
    .Columns(1).Width = 0
    End With
    
    On Error Resume Next
     'assigns values for labels
    If Trim(AdoTitle.Recordset.Fields("Title")) = "" Then
        lblTitle.Caption = "--"
    Else
        lblTitle.Caption = UCase(Trim(AdoTitle.Recordset.Fields("Title")))
    End If
    
    If Trim(AdoTitle.Recordset.Fields("Author")) = "" Then
        lblAuthor.Caption = "--"
    Else
        lblAuthor.Caption = UCase(Trim(AdoTitle.Recordset.Fields("Author")))
    End If
    
    Call Status("Ready")
    'enable delete button
    cmdDel.Enabled = True

Exit Sub

NotFound:
    
    MsgBox "Call ID not found. Please specify an existing call ID.", vbOKOnly + vbExclamation, "Library System"
    txtIdDel.SetFocus
    SendKeys HiLyt
End Sub



Private Sub cmdUpdate_Click()

    Dim upd_fields As String
    upd_fields = "" 'sets updated strings to null
    
    'field validation
    If Trim(txtEdTitle.Text) = "" Then
        txtEdTitle.SetFocus
        Missing
        Exit Sub
    ElseIf Trim(txtEdISBN.Text) = "" Then
        txtEdISBN.SetFocus
        Missing
        Exit Sub
    ElseIf Trim(txtEdAuthor.Text) = "" Then
        txtEdAuthor.SetFocus
        Missing
        Exit Sub
    End If
    
    'actual editing
    Call Status("Updating fields...")
    If AdoTitle.Recordset.Fields("Title") <> txtEdTitle.Text Then
    AdoTitle.Recordset.Fields("Title") = txtEdTitle.Text
    upd_fields = upd_fields & " Title"
    End If
    
    If AdoTitle.Recordset.Fields("ISBN") <> txtEdISBN.Text Then
    AdoTitle.Recordset.Fields("ISBN") = txtEdISBN.Text
    upd_fields = upd_fields & " ISBN"
    End If
    
    If AdoTitle.Recordset.Fields("Author") <> txtEdAuthor.Text Then
    AdoTitle.Recordset.Fields("Author") = txtEdAuthor.Text
    upd_fields = upd_fields & " Author"
    End If
    
    'save edited data
    AdoTitle.Recordset.Update
    Call Status("Ready")
    'confirm save
    MsgBox upd_fields & " successfully edited.", vbInformation + vbOKOnly, "Library System"
   
    
End Sub

Private Sub Form_Load()
    Call Status("Loading...")
    Call ConnectToDb(AdoTitle, "Title")
    Call Status("Ready")
End Sub




Private Sub tabTitle_Click(PreviousTab As Integer)
On Error Resume Next

    Select Case PreviousTab
        Case 0
        txtCallID.Text = ""
        txtTitle.Text = ""
        txtISBN.Text = ""
        txtAuthor.Text = ""
        Case 1
        txtEdCallID.Text = ""
        Call DisAbler
        Case 2
            txtIdDel.Text = ""
            ID = ""
            Call SQLDB(AdoBooks, "SELECT Book.BookId, Book.CallId FROM Book WHERE (((Book.CallId)='" & ID & "'));")
            AdoBooks.Refresh
        
        With dtgBooks
            .Refresh
            .Columns(0).Caption = "Book ID"
            .Columns(0).DataField = "BookId"
            .Columns(0).Width = 5800
        
            .Columns(1).Caption = "Call ID"
            .Columns(1).DataField = "CallId"
            .Columns(1).Width = 0
         End With
        
        
    End Select

    If tabTitle.Tab = 0 Then
        txtCallID.SetFocus
    ElseIf tabTitle.Tab = 1 Then
        txtEdCallID.SetFocus
    ElseIf tabTitle.Tab = 2 Then
        txtIdDel.SetFocus
    End If
    
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
    'disable buttons when not needed
    If Trim(txtCallID.Text) = "" And Trim(txtTitle.Text) = "" And Trim(txtISBN.Text) = "" And Trim(txtAuthor.Text) = "" Then
        cmdClear.Enabled = False
    Else
        cmdClear.Enabled = True
    End If
        
    If Trim(txtEdCallID.Text) = "" Then
        cmdSearch.Enabled = False
    Else
        cmdSearch.Enabled = True
    End If
        
    If Trim(txtIdDel.Text) = "" Then
        cmdSearchDel.Enabled = False
    Else
        cmdSearchDel.Enabled = True
    End If
    
    lblBooks.Caption = AdoBooks.Recordset.RecordCount
    
    If tabTitle.Tab = 0 Then
    'enable needed
        txtCallID.Enabled = True
        txtTitle.Enabled = True
        txtISBN.Enabled = True
        txtAuthor.Enabled = True
        cmdAdd.Enabled = True

    'disable rest
        txtEdCallID.Enabled = False
        cmdSearch.Enabled = False
        txtIdDel.Enabled = False
        cmdSearchDel.Enabled = False
        
    ElseIf tabTitle.Tab = 1 Then
    'enable needed
        txtEdCallID.Enabled = True

    'disable rest
        txtCallID.Enabled = False
        txtTitle.Enabled = False
        txtISBN.Enabled = False
        txtAuthor.Enabled = False
        cmdAdd.Enabled = False
        txtIdDel.Enabled = False
        cmdSearchDel.Enabled = False
    ElseIf tabTitle.Tab = 2 Then
    'enable needed
        txtIdDel.Enabled = True

    'disable rest
        txtCallID.Enabled = False
        txtTitle.Enabled = False
        txtISBN.Enabled = False
        txtAuthor.Enabled = False
        cmdAdd.Enabled = False
        txtEdCallID.Enabled = False
        cmdSearch.Enabled = False
    End If
End Sub


Private Sub txtAuthor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAdd.SetFocus
        SendKeys HiLyt
    End If
End Sub

Private Sub txtCallID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTitle.SetFocus
        SendKeys HiLyt
    End If
End Sub

Public Sub DisAbler()
        'disable textboxes
    txtEdTitle.Enabled = False
    txtEdTitle.Text = ""
    txtEdTitle.BackColor = &HC0C0C0
    
    txtEdISBN.Enabled = False
    txtEdISBN.Text = ""
    txtEdISBN.BackColor = &HC0C0C0
    
    txtEdAuthor.Enabled = False
    txtEdAuthor.Text = ""
    txtEdAuthor.BackColor = &HC0C0C0
    
    
    cmdUpdate.Enabled = False

End Sub

Private Sub txtEdAuthor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdUpdate.SetFocus
    End If
End Sub

Private Sub txtEdCallID_Change()
    Call DisAbler
End Sub

Private Sub txtEdCallID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
End Sub

Private Sub txtEdISBN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEdAuthor.SetFocus
        SendKeys HiLyt
    End If
End Sub

Private Sub txtEdTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEdISBN.SetFocus
        SendKeys HiLyt
    End If
End Sub

Private Sub txtIdDel_Change()
On Error Resume Next
        ID = ""
            Call SQLDB(AdoBooks, "SELECT Book.BookId, Book.CallId FROM Book WHERE (((Book.CallId)='" & ID & "'));")
            AdoBooks.Refresh
        
        lblTitle.Caption = "--"
        lblAuthor.Caption = "--"
        
        With dtgBooks
            .Refresh
            .Columns(0).Caption = "Book ID"
            .Columns(0).DataField = "BookId"
            .Columns(0).Width = 5800
        
            .Columns(1).Caption = "Call ID"
            .Columns(1).DataField = "CallId"
            .Columns(1).Width = 0
         End With
End Sub

Private Sub txtIdDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearchDel_Click
    End If
End Sub

Private Sub txtISBN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAuthor.SetFocus
        SendKeys HiLyt
    End If
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtISBN.SetFocus
        SendKeys HiLyt
    End If
End Sub

