VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBooks 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Manager"
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
      Left            =   2730
      Top             =   5115
   End
   Begin MSAdodcLib.Adodc AdoBooks 
      Height          =   390
      Left            =   225
      Top             =   5145
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
   Begin TabDlg.SSTab tabBooks 
      Height          =   3495
      Left            =   15
      TabIndex        =   18
      Top             =   945
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   6165
      _Version        =   393216
      TabOrientation  =   1
      Tab             =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Add Book"
      TabPicture(0)   =   "frmBooks.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Edit Book"
      TabPicture(1)   =   "frmBooks.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Delete Book"
      TabPicture(2)   =   "frmBooks.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdSearchDel"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtIdDel"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fmeDel"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame fmeDel 
         Caption         =   "Book information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   60
         TabIndex        =   27
         Top             =   645
         Width           =   6540
         Begin VB.PictureBox picContainer 
            Appearance      =   0  'Flat
            BackColor       =   &H00875B25&
            ForeColor       =   &H80000008&
            Height          =   1665
            Left            =   1470
            ScaleHeight     =   1635
            ScaleWidth      =   4890
            TabIndex        =   36
            Top             =   240
            Width           =   4920
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
               Index           =   0
               Left            =   75
               TabIndex        =   42
               Top             =   45
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
               TabIndex        =   41
               Top             =   315
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
               TabIndex        =   40
               Top             =   570
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
               TabIndex        =   39
               Top             =   840
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
               Index           =   4
               Left            =   75
               TabIndex        =   38
               Top             =   1095
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
               Index           =   5
               Left            =   75
               TabIndex        =   37
               Top             =   1365
               Width           =   120
            End
         End
         Begin Project1.lvButtons_H cmdDel 
            Height          =   405
            Left            =   5100
            TabIndex        =   29
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
            Image           =   "frmBooks.frx":0054
            Enabled         =   0   'False
            cBack           =   16777215
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Received:"
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
            Left            =   75
            TabIndex        =   35
            Top             =   315
            Width           =   1350
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
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
            TabIndex        =   34
            Top             =   1635
            Width           =   615
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
            Left            =   570
            TabIndex        =   33
            Top             =   1365
            Width           =   855
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
            Left            =   765
            TabIndex        =   32
            Top             =   1110
            Width           =   660
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
            Left            =   795
            TabIndex        =   31
            Top             =   840
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
            Left            =   975
            TabIndex        =   30
            Top             =   585
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Book Information"
         Height          =   3000
         Left            =   -74940
         TabIndex        =   26
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
            Left            =   1365
            TabIndex        =   3
            Top             =   1035
            Width           =   3630
         End
         Begin VB.TextBox txtEdition 
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
            Left            =   1365
            TabIndex        =   5
            Top             =   1560
            Width           =   3630
         End
         Begin VB.TextBox txtPublisher 
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
            Left            =   1365
            TabIndex        =   7
            Top             =   2085
            Width           =   3630
         End
         Begin VB.TextBox txtBookID 
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
            Left            =   1365
            TabIndex        =   1
            Top             =   510
            Width           =   3630
         End
         Begin Project1.lvButtons_H cmdAdd 
            Height          =   405
            Left            =   5130
            TabIndex        =   8
            Top             =   495
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
            Left            =   5130
            TabIndex        =   9
            Top             =   2085
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
         Begin VB.Label Label2 
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
            Left            =   225
            TabIndex        =   2
            Top             =   1155
            Width           =   1095
         End
         Begin VB.Label Label3 
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
            Left            =   660
            TabIndex        =   4
            Top             =   1665
            Width           =   660
         End
         Begin VB.Label Label4 
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
            Left            =   465
            TabIndex        =   6
            Top             =   2205
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Book Number:"
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
            Left            =   105
            TabIndex        =   0
            Top             =   630
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Book Information"
         Height          =   3000
         Left            =   -74940
         TabIndex        =   25
         Top             =   60
         Width           =   6540
         Begin VB.TextBox txtEdCallID 
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
            Left            =   1410
            TabIndex        =   14
            Top             =   1215
            Width           =   3630
         End
         Begin VB.TextBox txtEdEdition 
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
            Left            =   1410
            TabIndex        =   16
            Top             =   1740
            Width           =   3630
         End
         Begin VB.TextBox txtEdPublisher 
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
            Left            =   1410
            TabIndex        =   24
            Top             =   2265
            Width           =   3630
         End
         Begin VB.TextBox txtEdBookID 
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
            Left            =   1410
            TabIndex        =   11
            Top             =   390
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
            Image           =   "frmBooks.frx":0D2E
            cBack           =   16777215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            X1              =   15
            X2              =   6495
            Y1              =   1035
            Y2              =   1035
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   15
            X2              =   6495
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label Label6 
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
            Left            =   240
            TabIndex        =   13
            Top             =   1335
            Width           =   1095
         End
         Begin VB.Label Label7 
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
            Left            =   675
            TabIndex        =   15
            Top             =   1845
            Width           =   660
         End
         Begin VB.Label Label8 
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
            Left            =   480
            TabIndex        =   17
            Top             =   2385
            Width           =   855
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Book Number:"
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
            Left            =   120
            TabIndex        =   10
            Top             =   510
            Width           =   1215
         End
      End
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
         Left            =   1515
         TabIndex        =   21
         Top             =   150
         Width           =   3555
      End
      Begin Project1.lvButtons_H cmdSearchDel 
         Height          =   405
         Left            =   5160
         TabIndex        =   22
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
         Image           =   "frmBooks.frx":1A08
         cBack           =   16777215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Number:"
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
         TabIndex        =   20
         Top             =   270
         Width           =   1215
      End
   End
   Begin Project1.lvButtons_H cmdClose 
      Height          =   405
      Left            =   5205
      TabIndex        =   23
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
   Begin MSAdodcLib.Adodc AdoTitle 
      Height          =   390
      Left            =   1485
      Top             =   5145
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
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   15
      Picture         =   "frmBooks.frx":26E2
      Stretch         =   -1  'True
      Top             =   4485
      Width           =   6660
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBooks.frx":5069
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1095
      TabIndex        =   28
      Top             =   240
      Width           =   5040
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   405
      Picture         =   "frmBooks.frx":50F6
      Top             =   195
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmBooks.frx":5DC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6705
   End
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
On Error GoTo DuppErr
    Status "Validating fields..."
    'perform field validations
       
    If Trim(txtBookID.Text) = "" Then
        Missing
        txtBookID.SetFocus
        Exit Sub
    ElseIf Trim(txtCallID.Text) = "" Then
        Missing
        txtCallID.SetFocus
        Exit Sub
    ElseIf Trim(txtEdition.Text) = "" Then
        Missing
        txtEdition.SetFocus
        Exit Sub
    ElseIf Trim(txtPublisher.Text) = "" Then
        Missing
        txtPublisher.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtBookID.Text) = False Then
        MsgBox "Cannot accept non-numeric input for Book Number.", vbOKOnly + vbExclamation, "Library System"
        txtBookID.SetFocus
        SendKeys HiLyt
        Status "Ready"
        Exit Sub
    End If
    
    If IsNumeric(txtCallID.Text) = False Then
        MsgBox "Cannot accept non-numeric input for Call Number.", vbOKOnly + vbExclamation, "Library System"
        txtCallID.SetFocus
        SendKeys HiLyt
        Status "Ready"
        Exit Sub
    End If
    
    If (CallExist(txtCallID.Text) = False) Then  'checks existence of call number
        MsgBox "The Call Number specified does not exist. Please specify an existing Call Number.", vbOKOnly + vbExclamation, "Library System"
        txtCallID.SetFocus
        SendKeys HiLyt
        Status "Ready"
        Exit Sub
    End If
    
        
    Status "Saving information..."
    'transfer information to database
    AdoBooks.Refresh
    
    With AdoBooks.Recordset
        .AddNew
        .Fields("BookId") = Trim(txtBookID.Text)
        .Fields("CallId") = Trim(txtCallID.Text)
        .Fields("Edition") = Trim(txtEdition.Text)
        .Fields("Publisher") = Trim(txtPublisher.Text)
        .Fields("DateReg") = Date
        .Update
    End With
    
    MsgBox "Book added to library.", vbOKOnly + vbInformation, "Library System"
    
    'clear input fields
    Call ClearAll
    Status "Ready"
    
Exit Sub

NoCallID:
    MsgBox "The Call Number specified does exists. Please specify an existing Call Number.", vbOKOnly + vbExclamation, "Library System"
    txtCallID.SetFocus
    SendKeys HiLyt
    AdoBooks.Recordset.CancelUpdate
Exit Sub

DuppErr:
    MsgBox "The Book Number specified already exists. Please specify a different Book Number.", vbOKOnly + vbExclamation, "Library System"
    txtBookID.SetFocus
    SendKeys HiLyt
    AdoBooks.Recordset.CancelUpdate
End Sub

Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdClose_Click()
    Call TotalCount
    Unload Me
End Sub


Private Sub cmdDel_Click()
On Error Resume Next
Dim counter As Integer
    If AdoBooks.Recordset.Fields("StatusID") = 2 Then
        MsgBox "Canot delete book that has been borrowed.", vbOKOnly + vbExclamation, "Library System"
        Exit Sub
    End If
    
    If MsgBox("Deleting selected book. Proceed?", vbYesNo + vbQuestion, "Library System") = vbYes Then
        AdoBooks.Recordset.Delete
        For counter = 0 To 5
            lblInfo(counter).Caption = "--"
        Next counter
        cmdDel.Enabled = False
        txtIdDel.SetFocus
        SendKeys HiLyt
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdSearch_Click()
On Error GoTo NotFound
        If Trim(txtEdBookID.Text) = "" Then
            Exit Sub
        End If
        
        If AdoBooks.Recordset.RecordCount = 0 Then
            MsgBox "There are no existing titles to search.", vbOKOnly + vbExclamation, "Library System"
            Exit Sub
        End If
    
    AdoBooks.Refresh
    Call Status("Searching...")
    AdoBooks.Recordset.Find ("BookID = '" & Trim(txtEdBookID.Text) & "'")
    
    'assign values
    txtEdCallID.Text = AdoBooks.Recordset.Fields("CallID")
    txtEdEdition.Text = AdoBooks.Recordset.Fields("Edition")
    txtEdPublisher.Text = AdoBooks.Recordset.Fields("Publisher")
    
    'enable textboxes
    txtEdCallID.Enabled = True
    txtEdCallID.BackColor = vbWhite
    
    txtEdEdition.Enabled = True
    txtEdEdition.BackColor = vbWhite
    
    txtEdPublisher.Enabled = True
    txtEdPublisher.BackColor = vbWhite
    
    
    cmdUpdate.Enabled = True
    
    txtEdCallID.SetFocus
    SendKeys HiLyt
    
    Call Status("Ready")
    Exit Sub

NotFound:
    MsgBox "Book Number not found. Please specify an existing book number.", vbOKOnly + vbExclamation, "Library System"
    ClearAll
    txtEdBookID.SetFocus
    SendKeys HiLyt
    Status "Ready"
End Sub

Private Sub cmdSearchDel_Click()
Dim Stat As Integer

On Error GoTo NotFound
    
        If AdoBooks.Recordset.RecordCount = 0 Then
            MsgBox "There are no existing titles to search.", vbOKOnly + vbExclamation, "Library System"
            Exit Sub
        End If
    
    AdoBooks.Refresh
    AdoTitle.Refresh
    Call Status("Searching...")
    'finds appropriate Book
    AdoBooks.Recordset.Find ("BookID = '" & Trim(txtIdDel.Text) & "'")
    'finds appropriate Title
    AdoTitle.Recordset.Find ("CallId = '" & Trim(AdoBooks.Recordset.Fields("CallId")) & "'")
    
    On Error Resume Next
    If Trim(AdoBooks.Recordset.Fields("DateReg")) = "" Then
        lblInfo(0).Caption = "--"
    Else
        lblInfo(0).Caption = UCase(AdoBooks.Recordset.Fields("DateReg"))
    End If
    
    If Trim(AdoTitle.Recordset.Fields("Title")) = "" Then
        lblInfo(1).Caption = "--"
    Else
        lblInfo(1).Caption = UCase(AdoTitle.Recordset.Fields("Title"))
    End If
    
    If Trim(AdoTitle.Recordset.Fields("Author")) = "" Then
        lblInfo(2).Caption = "--"
    Else
        lblInfo(2).Caption = UCase(AdoTitle.Recordset.Fields("Author"))
    End If
    
    If Trim(AdoBooks.Recordset.Fields("Edition")) = "" Then
        lblInfo(3).Caption = "--"
    Else
        lblInfo(3).Caption = UCase(AdoBooks.Recordset.Fields("Edition"))
    End If
    
    If AdoBooks.Recordset.Fields("Publisher") = "" Then
        lblInfo(4).Caption = "--"
    Else
        lblInfo(4).Caption = UCase(AdoBooks.Recordset.Fields("Publisher"))
    End If
    
    Stat = Val(AdoBooks.Recordset.Fields("StatusID"))
        
    Select Case Stat
        Case 1
            lblInfo(5).Caption = "IN"
        Case 2
            lblInfo(5).Caption = "BORROWED"
        Case 3
            lblInfo(5).Caption = "LOST"
    End Select
    
    cmdDel.Enabled = True
    cmdDel.SetFocus
    Status "Ready"
    
Exit Sub

NotFound:
    MsgBox "Book Number not found. Please specify an existing book number.", vbOKOnly + vbExclamation, "Library System"
    ClearAll
    txtIdDel.SetFocus
    Status "Ready"
    SendKeys HiLyt
End Sub

Private Sub cmdUpdate_Click()
    Dim Fields As String

    'field validation
    Status "Validating fields..."
   

    If Trim(txtEdCallID.Text) = "" Then
        Missing
        txtEdCallID.SetFocus
    ElseIf Trim(txtEdEdition.Text) = "" Then
        Missing
        txtEdEdition.SetFocus
    ElseIf Trim(txtEdPublisher.Text) = "" Then
        Missing
        txtEdPublisher.SetFocus
    End If
    
    If IsNumeric(txtEdCallID.Text) = False Then
        MsgBox "Cannot accept non-numeric input for Call Number.", vbOKOnly + vbExclamation, "Library System"
        txtEdCallID.SetFocus
        SendKeys HiLyt
        Exit Sub
    End If
    
    'checks valid Call Number
    If (CallExist(txtEdCallID.Text) = False) Then  'checks existence of call number
        MsgBox "The Call Number specified does not exist. Please specify an existing Call Number.", vbOKOnly + vbExclamation, "Library System"
        txtEdCallID.SetFocus
        SendKeys HiLyt
        Status "Ready"
        Exit Sub
    End If
    
    Status "Updating information..."
    With AdoBooks.Recordset
        If Not .Fields("CallID") = Trim(txtEdCallID.Text) Then
            .Fields("CallID") = Trim(txtEdCallID.Text)
            Fields = Fields & " Call ID"
        End If
        
        If Not .Fields("Edition") = Trim(txtEdEdition.Text) Then
            .Fields("Edition") = Trim(txtEdEdition.Text)
            Fields = Fields & " Edition"
        End If
        
        If Not .Fields("Publisher") = Trim(txtEdPublisher.Text) Then
            .Fields("Publisher") = Trim(txtEdPublisher.Text)
            Fields = Fields & " Publisher"
        End If
        
        If Trim(Fields) = "" Then
            Exit Sub
        Else
            .Update
            MsgBox "Update successful on the following: " & Fields & ".", vbOKOnly + vbInformation, "Library System"
        End If
    End With
    AdoBooks.Refresh
    
    Status "Ready"
    
    MsgBox "Record successfully edited.", vbOKOnly + vbInformation, "Library System"
    ClearAll
End Sub



Private Sub Form_Load()
    Status "Loading..."
    Call ConnectToDb(AdoBooks, "Book")
    Call ConnectToDb(AdoTitle, "Title")
    Status "Ready"
End Sub

Public Sub ClearAll()
On Error Resume Next
Dim counter As Integer

    txtBookID.Text = ""
    txtCallID.Text = ""
    txtEdition.Text = ""
    txtPublisher.Text = ""
    
    txtEdBookID.Text = ""
    txtEdCallID.Text = ""
    txtEdEdition.Text = ""
    txtEdPublisher.Text = ""
    
    txtEdCallID.Enabled = False
    txtEdCallID.BackColor = &HC0C0C0
    txtEdEdition.Enabled = False
    txtEdEdition.BackColor = &HC0C0C0
    txtEdPublisher.Enabled = False
    txtEdPublisher.BackColor = &HC0C0C0
    
    txtIdDel.Text = ""
    
For counter = 0 To 5
    lblInfo(counter).Caption = "--"
Next counter
    
End Sub



Private Sub tabBooks_Click(PreviousTab As Integer)
On Error Resume Next
AdoBooks.Refresh
    Select Case PreviousTab
        Case 0
        ClearAll
        Case 1
        ClearAll
        Case 2
        ClearAll
    End Select
    
     If tabBooks.Tab = 0 Then
         txtBookID.SetFocus
     ElseIf tabBooks.Tab = 1 Then
         txtEdBookID.SetFocus
     ElseIf tabBooks.Tab = 2 Then
         txtIdDel.SetFocus
     End If
End Sub

Public Function CallExist(SearchItem As String) As Boolean
'this function verifies the existence of the Call Number to be assigned
    On Error GoTo ErrHandler
    Dim temp As String
    
    AdoTitle.Refresh
    
    AdoTitle.Recordset.Find ("CallID = '" & Trim(SearchItem) & "'")
    temp = AdoTitle.Recordset.Fields("CallID") 'assigns to a vaiable to validate existence
    
    CallExist = True
Exit Function

ErrHandler:
'Call Number does not exist
    CallExist = False
End Function

Private Sub Timer1_Timer()
    If Trim(txtBookID.Text) = "" And Trim(txtCallID.Text) = "" And Trim(txtEdition.Text) = "" _
        And Trim(txtPublisher.Text) = "" Then
            cmdClear.Enabled = False
    Else
            cmdClear.Enabled = True
    End If
    
    If Trim(txtEdBookID.Text) = "" Then
        cmdSearch.Enabled = False
    Else
        cmdSearch.Enabled = True
    End If
    
    If Trim(txtIdDel.Text) = "" Then
        cmdSearchDel.Enabled = False
    Else
        cmdSearchDel.Enabled = True
    End If
    
    If tabBooks.Tab = 0 Then
    'enable needed fields
        txtBookID.Enabled = True
        txtCallID.Enabled = True
        txtEdition.Enabled = True
        txtPublisher.Enabled = True
        cmdAdd.Enabled = True
      
    'disable unnecessary ones
        txtEdBookID.Enabled = False
        cmdSearch.Enabled = False
        txtIdDel.Enabled = False
        cmdSearchDel.Enabled = False
        
    ElseIf tabBooks.Tab = 1 Then
    'enable needed fields
        txtEdBookID.Enabled = True
       
    'disable unnecessary ones
        txtIdDel.Enabled = False
        cmdSearchDel.Enabled = False
        txtBookID.Enabled = False
        txtCallID.Enabled = False
        txtEdition.Enabled = False
        txtPublisher.Enabled = False
        cmdAdd.Enabled = False
    ElseIf tabBooks.Tab = 2 Then
    'enable needed fields
        txtIdDel.Enabled = True
        
    'disable unnecesary ones
        
        txtEdBookID.Enabled = False
        cmdSearch.Enabled = False
        txtBookID.Enabled = False
        txtCallID.Enabled = False
        txtEdition.Enabled = False
        txtPublisher.Enabled = False
        cmdAdd.Enabled = False
    End If
End Sub


Private Sub txtBookID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCallID.SetFocus
        SendKeys HiLyt
    End If
End Sub


Private Sub txtCallID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEdition.SetFocus
        SendKeys HiLyt
    End If
    
End Sub
Private Sub txtEdBookID_Change()
On Error Resume Next
    
    txtEdCallID.Text = ""
    txtEdEdition.Text = ""
    txtEdPublisher.Text = ""
    
    txtEdCallID.Enabled = False
    txtEdCallID.BackColor = &HC0C0C0
    txtEdEdition.Enabled = False
    txtEdEdition.BackColor = &HC0C0C0
    txtEdPublisher.Enabled = False
    txtEdPublisher.BackColor = &HC0C0C0
    
End Sub

Private Sub txtEdBookID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
End Sub

Private Sub txtEdCallID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEdEdition.SetFocus
        SendKeys HiLyt
    End If
    
End Sub


Private Sub txtEdEdition_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEdPublisher.SetFocus
        SendKeys HiLyt
    End If

End Sub

Private Sub txtEdition_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPublisher.SetFocus
        SendKeys HiLyt
    End If

End Sub


Private Sub txtEdPublisher_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdUpdate.SetFocus
    End If

End Sub

Private Sub txtIdDel_Change()
On Error Resume Next
Dim counter As Integer
    For counter = 0 To 5
        lblInfo(counter).Caption = "--"
    Next counter
    cmdDel.Enabled = False
End Sub

Private Sub txtIdDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearchDel_Click
    End If
End Sub

Private Sub txtPublisher_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAdd.SetFocus
    End If

End Sub
