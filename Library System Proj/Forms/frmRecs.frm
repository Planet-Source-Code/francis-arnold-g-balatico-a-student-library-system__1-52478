VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRecs 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record Drawer"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   ControlBox      =   0   'False
   Icon            =   "frmRecs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6405
      Top             =   5685
   End
   Begin TabDlg.SSTab tabRecs 
      Height          =   5580
      Left            =   53
      TabIndex        =   0
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   9843
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8208173
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Borrower Records"
      TabPicture(0)   =   "frmRecs.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picBorrower"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Book Records"
      TabPicture(1)   =   "frmRecs.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picBooks"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Due Books"
      TabPicture(2)   =   "frmRecs.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picDue"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picBooks 
         BackColor       =   &H00B5742D&
         Height          =   5175
         Left            =   -74940
         ScaleHeight     =   5115
         ScaleWidth      =   9705
         TabIndex        =   53
         Top             =   60
         Width           =   9765
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H00B5742D&
            Caption         =   "Selection options"
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
            Height          =   1260
            Left            =   7395
            TabIndex        =   73
            Top             =   15
            Width           =   2265
            Begin VB.ComboBox cmbBooks 
               Height          =   315
               ItemData        =   "frmRecs.frx":0060
               Left            =   135
               List            =   "frmRecs.frx":0073
               Style           =   2  'Dropdown List
               TabIndex        =   74
               Top             =   525
               Width           =   1995
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Order by:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   165
               TabIndex        =   75
               Top             =   255
               Width           =   720
            End
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00B5742D&
            Caption         =   "Information"
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
            Height          =   2700
            Left            =   30
            TabIndex        =   59
            Top             =   0
            Width           =   7335
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Registered:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   105
               TabIndex        =   77
               Top             =   2325
               Width           =   1200
            End
            Begin VB.Label lblBookInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   6
               Left            =   1440
               TabIndex        =   76
               Top             =   2325
               Width           =   120
            End
            Begin VB.Label lblBookInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   5
               Left            =   1440
               TabIndex        =   69
               Top             =   1935
               Width           =   120
            End
            Begin VB.Label lblBookInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   4
               Left            =   1440
               TabIndex        =   68
               Top             =   1515
               Width           =   120
            End
            Begin VB.Label lblBookInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   3
               Left            =   1440
               TabIndex        =   67
               Top             =   285
               Width           =   120
            End
            Begin VB.Label lblBookInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   2
               Left            =   1440
               TabIndex        =   66
               Top             =   1110
               Width           =   120
            End
            Begin VB.Label lblBookInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   1
               Left            =   1440
               TabIndex        =   65
               Top             =   690
               Width           =   120
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Publisher:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   600
               TabIndex        =   64
               Top             =   1935
               Width           =   705
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Call Number:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   405
               TabIndex        =   63
               Top             =   1110
               Width           =   900
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Author:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   765
               TabIndex        =   62
               Top             =   1515
               Width           =   540
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Title:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   975
               TabIndex        =   61
               Top             =   285
               Width           =   330
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Book Number:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   300
               TabIndex        =   60
               Top             =   690
               Width           =   1005
            End
            Begin VB.Shape Shape7 
               BackColor       =   &H00404040&
               BackStyle       =   1  'Opaque
               Height          =   2355
               Left            =   1305
               Top             =   225
               Width           =   5880
            End
         End
         Begin VB.PictureBox picBookHide 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1680
            Left            =   15
            Picture         =   "frmRecs.frx":00B5
            ScaleHeight     =   1650
            ScaleWidth      =   9630
            TabIndex        =   58
            Top             =   2985
            Width           =   9660
         End
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H00B5742D&
            Caption         =   "Search"
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
            Height          =   1380
            Left            =   7395
            TabIndex        =   54
            Top             =   1320
            Width           =   2265
            Begin VB.TextBox txtBookSearch 
               Height          =   315
               Left            =   150
               TabIndex        =   55
               Top             =   495
               Width           =   1965
            End
            Begin Project1.lvButtons_H cmdSearchBook 
               Height          =   405
               Left            =   165
               TabIndex        =   56
               Top             =   870
               Width           =   1950
               _ExtentX        =   3440
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
               Image           =   "frmRecs.frx":44E45
               cBack           =   16777215
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Book ID:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   165
               TabIndex        =   57
               Top             =   255
               Width           =   630
            End
         End
         Begin MSAdodcLib.Adodc AdoBooks 
            Height          =   330
            Left            =   8415
            Top             =   4710
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
         Begin MSDataGridLib.DataGrid dtgBookList 
            Height          =   1905
            Left            =   15
            TabIndex        =   70
            Top             =   2760
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   3360
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   781309
            HeadLines       =   1
            RowHeight       =   15
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
         Begin VB.Label lblNumBooks 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1830
            TabIndex        =   72
            Top             =   4785
            Width           =   90
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Books:"
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
            Left            =   285
            TabIndex        =   71
            Top             =   4785
            Width           =   1485
         End
         Begin VB.Shape Shape8 
            BackStyle       =   1  'Opaque
            FillColor       =   &H007D3F2D&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   15
            Top             =   4680
            Width           =   9660
         End
      End
      Begin VB.PictureBox picBorrower 
         BackColor       =   &H00B5742D&
         Height          =   5175
         Left            =   60
         ScaleHeight     =   5115
         ScaleWidth      =   9705
         TabIndex        =   28
         Top             =   60
         Width           =   9765
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00B5742D&
            Caption         =   "Search"
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
            Height          =   1380
            Left            =   7395
            TabIndex        =   47
            Top             =   1320
            Width           =   2265
            Begin VB.TextBox txtSearch 
               Height          =   315
               Left            =   150
               TabIndex        =   49
               Top             =   495
               Width           =   1965
            End
            Begin Project1.lvButtons_H cmdSearch 
               Height          =   405
               Left            =   165
               TabIndex        =   50
               Top             =   870
               Width           =   1950
               _ExtentX        =   3440
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
               Image           =   "frmRecs.frx":45B1F
               cBack           =   16777215
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Borrower ID:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   165
               TabIndex        =   48
               Top             =   255
               Width           =   885
            End
         End
         Begin VB.PictureBox picBorHide 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   15
            Picture         =   "frmRecs.frx":467F9
            ScaleHeight     =   1395
            ScaleWidth      =   9630
            TabIndex        =   42
            Top             =   2985
            Width           =   9660
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00B5742D&
            Caption         =   "Information"
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
            Height          =   2700
            Left            =   30
            TabIndex        =   31
            Top             =   0
            Width           =   7335
            Begin VB.PictureBox Picture1 
               AutoRedraw      =   -1  'True
               BackColor       =   &H007D3F2D&
               Height          =   1545
               Left            =   5475
               ScaleHeight     =   1485
               ScaleWidth      =   1485
               TabIndex        =   51
               Top             =   570
               Width           =   1545
               Begin VB.Image imgPic 
                  Height          =   1470
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   1470
               End
               Begin VB.Label Label12 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Photo"
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
                  Left            =   495
                  TabIndex        =   52
                  Top             =   660
                  Width           =   510
               End
            End
            Begin VB.Shape Shape5 
               BackColor       =   &H007D3F2D&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000BEBFD&
               BorderWidth     =   2
               Height          =   1650
               Left            =   5430
               Top             =   525
               Width           =   1650
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Borrower ID:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   375
               TabIndex        =   41
               Top             =   480
               Width           =   930
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Course:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   735
               TabIndex        =   40
               Top             =   1305
               Width           =   570
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Number:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   105
               TabIndex        =   39
               Top             =   1710
               Width           =   1200
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Borrower Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   105
               TabIndex        =   38
               Top             =   885
               Width           =   1200
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Registered:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   105
               TabIndex        =   37
               Top             =   2130
               Width           =   1200
            End
            Begin VB.Label lblBorInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   1
               Left            =   1440
               TabIndex        =   36
               Top             =   480
               Width           =   120
            End
            Begin VB.Label lblBorInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   2
               Left            =   1440
               TabIndex        =   35
               Top             =   885
               Width           =   120
            End
            Begin VB.Label lblBorInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   3
               Left            =   1440
               TabIndex        =   34
               Top             =   1305
               Width           =   120
            End
            Begin VB.Label lblBorInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   4
               Left            =   1440
               TabIndex        =   33
               Top             =   1710
               Width           =   120
            End
            Begin VB.Label lblBorInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   5
               Left            =   1440
               TabIndex        =   32
               Top             =   2130
               Width           =   120
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00404040&
               BackStyle       =   1  'Opaque
               Height          =   2355
               Left            =   1305
               Top             =   225
               Width           =   5880
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00B5742D&
            Caption         =   "Selection options"
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
            Height          =   1260
            Left            =   7395
            TabIndex        =   29
            Top             =   15
            Width           =   2265
            Begin VB.ComboBox cmbBorOrder 
               Height          =   315
               ItemData        =   "frmRecs.frx":8B589
               Left            =   135
               List            =   "frmRecs.frx":8B599
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   525
               Width           =   1995
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Order by:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   165
               TabIndex        =   30
               Top             =   255
               Width           =   720
            End
         End
         Begin MSAdodcLib.Adodc AdoBor 
            Height          =   330
            Left            =   8415
            Top             =   4710
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
         Begin MSDataGridLib.DataGrid dtgBorGrid 
            Height          =   1905
            Left            =   15
            TabIndex        =   43
            Top             =   2760
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   3360
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   781309
            HeadLines       =   1
            RowHeight       =   15
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
               ScrollBars      =   3
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Number of Records:"
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
            Left            =   120
            TabIndex        =   45
            Top             =   4785
            Width           =   1650
         End
         Begin VB.Label lblRec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1830
            TabIndex        =   44
            Top             =   4785
            Width           =   90
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            FillColor       =   &H007D3F2D&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   15
            Top             =   4680
            Width           =   9660
         End
      End
      Begin VB.PictureBox picDue 
         BackColor       =   &H00B5742D&
         Height          =   5175
         Left            =   -74940
         ScaleHeight     =   5115
         ScaleWidth      =   9705
         TabIndex        =   1
         Top             =   60
         Width           =   9765
         Begin MSAdodcLib.Adodc AdoDue 
            Height          =   330
            Left            =   8415
            Top             =   4710
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
            BackColor       =   &H00B5742D&
            Caption         =   "Selection options"
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
            Height          =   2685
            Left            =   7395
            TabIndex        =   21
            Top             =   15
            Width           =   2265
            Begin VB.ComboBox cmbOrder 
               Height          =   315
               ItemData        =   "frmRecs.frx":8B5C9
               Left            =   210
               List            =   "frmRecs.frx":8B5DC
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   1860
               Width           =   1890
            End
            Begin VB.OptionButton optAll 
               BackColor       =   &H00B5742D&
               Caption         =   "All"
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   225
               TabIndex        =   24
               Top             =   525
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optDaysPast 
               BackColor       =   &H00B5742D&
               Caption         =   "With days past due"
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   225
               TabIndex        =   23
               Top             =   1155
               Width           =   1965
            End
            Begin VB.OptionButton optDue 
               BackColor       =   &H00B5742D&
               Caption         =   "Due today"
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   225
               TabIndex        =   22
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Order by:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   26
               Top             =   1650
               Width           =   720
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00B5742D&
            Caption         =   "Information"
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
            Height          =   2700
            Left            =   30
            TabIndex        =   5
            Top             =   0
            Width           =   7335
            Begin VB.Label lblDueInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   6
               Left            =   1440
               TabIndex        =   20
               Top             =   2295
               Width           =   120
            End
            Begin VB.Label lblDueInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   5
               Left            =   1440
               TabIndex        =   19
               Top             =   1980
               Width           =   120
            End
            Begin VB.Label lblDueInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   4
               Left            =   1440
               TabIndex        =   18
               Top             =   1650
               Width           =   120
            End
            Begin VB.Label lblDueInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   3
               Left            =   1440
               TabIndex        =   17
               Top             =   1335
               Width           =   120
            End
            Begin VB.Label lblDueInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   2
               Left            =   1440
               TabIndex        =   16
               Top             =   1020
               Width           =   120
            End
            Begin VB.Label lblDueInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   1
               Left            =   1440
               TabIndex        =   15
               Top             =   690
               Width           =   120
            End
            Begin VB.Label lblDueInfo 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H000BEBFD&
               Height          =   210
               Index           =   0
               Left            =   1440
               TabIndex        =   14
               Top             =   375
               Width           =   120
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Days Past Due:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   195
               TabIndex        =   13
               Top             =   2295
               Width           =   1110
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Due:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   600
               TabIndex        =   12
               Top             =   1980
               Width           =   705
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Borrower Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   105
               TabIndex        =   11
               Top             =   690
               Width           =   1200
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Borrowed:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   150
               TabIndex        =   10
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Title:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   975
               TabIndex        =   9
               Top             =   1335
               Width           =   330
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Book Number:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   300
               TabIndex        =   8
               Top             =   1020
               Width           =   1005
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Borrower ID:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   375
               TabIndex        =   7
               Top             =   375
               Width           =   930
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00404040&
               BackStyle       =   1  'Opaque
               Height          =   2355
               Left            =   1305
               Top             =   225
               Width           =   5880
            End
         End
         Begin VB.PictureBox picHide 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   15
            Picture         =   "frmRecs.frx":8B618
            ScaleHeight     =   1395
            ScaleWidth      =   9630
            TabIndex        =   27
            Top             =   2985
            Width           =   9660
         End
         Begin MSDataGridLib.DataGrid dtgBooks 
            Height          =   1905
            Left            =   15
            TabIndex        =   3
            Top             =   2760
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   3360
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   781309
            HeadLines       =   1
            RowHeight       =   15
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
               ScrollBars      =   3
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
         Begin VB.Label lblUnRet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1830
            TabIndex        =   6
            Top             =   4785
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unreturned Books:"
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
            Left            =   195
            TabIndex        =   4
            Top             =   4785
            Width           =   1560
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            FillColor       =   &H007D3F2D&
            FillStyle       =   0  'Solid
            Height          =   405
            Left            =   15
            Top             =   4680
            Width           =   9660
         End
      End
   End
   Begin Project1.lvButtons_H cmdClose 
      Height          =   405
      Left            =   8475
      TabIndex        =   2
      Top             =   5685
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
      Left            =   45
      Picture         =   "frmRecs.frx":D03A8
      Stretch         =   -1  'True
      Top             =   5610
      Width           =   9900
   End
End
Attribute VB_Name = "frmRecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fieldName As String
Private fieldName2 As String
Private fieldName3 As String


Private Sub AdoBooks_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    Call AssignBookVal
End Sub


Private Sub AdoBor_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    Call AssignBorVal
End Sub

Private Sub AdoDue_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    Call AssignDueVal
End Sub



Private Sub cmbBooks_Click()
On Error Resume Next
    If AdoBooks.Recordset.RecordCount > 1 Then
        Call ConvertToFieldBooks
        Call SQLDB(AdoBooks, "SELECT Book.BookId, Title.CallId, Title.Title, Title.Author, Book.Publisher, Book.DateReg FROM Title INNER JOIN Book ON Title.CallId = Book.CallId ORDER by " & fieldName3)
        Call AssignBookVal
        Call SetBookGrid
    
    End If
End Sub

Private Sub cmbBorOrder_Click()
    On Error Resume Next
    If AdoBor.Recordset.RecordCount > 1 Then
        Call ConvertToFieldsBor
        Call SQLDB(AdoBor, "Select * from Borrower ORDER by " & fieldName2)
        Call AssignBorVal
    
        Call BorGrid
    End If
End Sub

Private Sub cmbOrder_Click()

    If optAll Then
        Call ConvertToFields
        Call SQLDB(AdoDue, "Select * from Book_Loan where Status = 0 ORDER by " & fieldName)
        
        Call Setduegrid
    End If
    
    If optDaysPast Then
        Call ConvertToFields
        Call SQLDB(AdoDue, "Select * from Book_Loan where Date_Due < '" & Date & "' and Status = 0 ORDER by " & fieldName)
        
        Call Setduegrid
    End If
    
    If optDue Then
        Call ConvertToFields
        Call SQLDB(AdoDue, "Select * from Book_Loan where Date_Due = '" & Date & "' and Status = 0 ORDER by " & fieldName)
        Call Setduegrid
    End If

End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdSearch_Click()
    On Error GoTo NotFound
    Dim temp As String
    
        AdoBor.Refresh
        AdoBor.Recordset.Find ("BorId = '" & txtSearch.Text & "'")
        temp = AdoBor.Recordset.Fields(1)
        
        txtSearch.SetFocus
        Call BorGrid
        SendKeys HiLyt
    Exit Sub

NotFound:
    MsgBox "The record you requested could not be found.", vbOKOnly + vbExclamation, "Library System"
    txtSearch.SetFocus
    SendKeys HiLyt
End Sub

Private Sub cmdSearchBook_Click()
On Error GoTo NotFound
Dim temp As String

AdoBooks.Refresh
AdoBooks.Recordset.Find ("BookId = '" & txtBookSearch.Text & "'")
temp = AdoBooks.Recordset.Fields(1)

txtBookSearch.SetFocus
Call SetBookGrid
SendKeys HiLyt
Exit Sub

NotFound:
    MsgBox "The record you requested could not be found.", vbOKOnly + vbExclamation, "Library System"
    txtBookSearch.SetFocus
    SendKeys HiLyt
End Sub


Private Sub Form_Load()
On Error Resume Next
    Status "Loading records. Please wait..."
    cmbOrder.ListIndex = 0
    cmbBorOrder.ListIndex = 0
    cmbBooks.ListIndex = 0
    
    Call ConvertToFields
    Call SQLDB(AdoDue, "Select * from Book_Loan where Status = 0 ORDER by " & fieldName)
    Call AssignDueVal
    Call Setduegrid
    
    Call ConvertToFieldsBor
    Call SQLDB(AdoBor, "Select * from Borrower ORDER by " & fieldName2)
    Call AssignBorVal
    Call BorGrid
    
    Call ConvertToFieldBooks
    Call SQLDB(AdoBooks, "SELECT Book.BookId, Title.CallId, Title.Title, Title.Author, Book.Publisher, Book.DateReg FROM Title INNER JOIN Book ON Title.CallId = Book.CallId ORDER by " & fieldName3)
    Call AssignBookVal
    Call SetBookGrid
    
    Status "Ready"
End Sub

Private Sub Setduegrid()
'ensures that the column formatting remains the same
Set dtgBooks.DataSource = AdoDue

    With dtgBooks
        .Columns(0).DataField = "Book_Id"
        .Columns(0).Caption = "Book ID"
        .Columns(0).Width = 1500
                
        .Columns(1).DataField = "Call_ID"
        .Columns(1).Caption = "Call ID"
        .Columns(1).Width = 1500
        
        .Columns(2).DataField = "Book_Title"
        .Columns(2).Caption = "Title"
        .Columns(2).Width = 3000
        
        .Columns(3).DataField = "Borrower_Id"
        .Columns(3).Caption = "Borrower ID"
        .Columns(3).Width = 1500
        
        .Columns(4).DataField = "Borrower_FName"
        .Columns(4).Caption = "First Name"
        .Columns(4).Width = 1800
        
        .Columns(5).DataField = "Borrower_LName"
        .Columns(5).Caption = "Last Name"
        .Columns(5).Width = 1800
        
        .Columns(6).DataField = "Date_Borrowed"
        .Columns(6).Caption = "Date Borrowed"
        .Columns(6).Width = 1500
        
        .Columns(7).DataField = "Date_Due"
        .Columns(7).Caption = "Due Date"
        .Columns(7).Width = 1500
        
        .Columns(8).DataField = "Date_Returned"
        .Columns(8).Caption = "Date Returned"
        .Columns(8).Width = 0
        
        .Columns(9).DataField = "Days_Past_Due"
        .Columns(9).Caption = "Days Past Due"
        .Columns(9).Width = 1500
               
        .Columns(10).DataField = "Fines"
        .Columns(10).Caption = "Fine"
        .Columns(10).Width = 0
        
        .Columns(11).DataField = "Status"
        .Columns(11).Caption = "Status"
        .Columns(11).Width = 0
    End With
End Sub

Private Sub optAll_Click()
If optAll.Value Then
    Call ConvertToFields
    Call SQLDB(AdoDue, "Select * from Book_Loan where Status = 0 ORDER by " & fieldName)
    
    Call Setduegrid
End If
End Sub

Private Sub optDaysPast_Click()
If optDaysPast Then
    Call ConvertToFields
    Call SQLDB(AdoDue, "Select * from Book_Loan where Date_Due < '" & Date & "' and Status = 0 ORDER by " & fieldName)
    
    Call Setduegrid
End If
End Sub

Private Sub optDue_Click()
If optDue Then
    Call ConvertToFields
    Call SQLDB(AdoDue, "Select * from Book_Loan where Date_Due = '" & Date & "' and Status = 0 ORDER by " & fieldName)
    Call Setduegrid
End If
End Sub


Private Sub tabRecs_Click(PreviousTab As Integer)
On Error Resume Next
    With tabRecs
        If .Tab = 0 Then
            
            txtSearch.SetFocus
        ElseIf .Tab = 1 Then
           
            txtBookSearch.SetFocus
        ElseIf .Tab = 2 Then
          
            cmbOrder.SetFocus
        End If
    End With
End Sub


Private Sub Timer1_Timer()
On Error Resume Next

    lblUnRet.Caption = AdoDue.Recordset.RecordCount
    lblRec.Caption = AdoBor.Recordset.RecordCount
    lblNumBooks.Caption = AdoBooks.Recordset.RecordCount
    
    If AdoBor.Recordset.RecordCount = 0 Then
        picBorHide.Visible = True
    Else
        picBorHide.Visible = False
    End If
    
    If AdoDue.Recordset.RecordCount = 0 Then
        picHide.Visible = True
    Else
        picHide.Visible = False
    End If
    
    If AdoBooks.Recordset.RecordCount = 0 Then
        picBookHide.Visible = True
    Else
        picBookHide.Visible = False
    End If
End Sub

Private Sub ConvertToFields()
   
   'for cmborder
    With cmbOrder
        If .Text = "Book ID" Then
            fieldName = "Book_Id"
            
        ElseIf .Text = "Borrower ID" Then
            fieldName = "Borrower_Id"
            
        ElseIf .Text = "Date Borrowed" Then
            fieldName = "Date_Borrowed"
        
        ElseIf .Text = "Date Due" Then
            fieldName = "Date_Due"
        
        ElseIf .Text = "Call ID" Then
            fieldName = "Call_ID"
        End If
    End With
    
End Sub

Private Sub ConvertToFieldsBor()
'for cmbBorOrder
    With cmbBorOrder
        If .Text = "Borrower ID" Then
            fieldName2 = "BorId"
        ElseIf .Text = "Course" Then
            fieldName2 = "Course"
        ElseIf .Text = "Date Registered" Then
            fieldName2 = "Date_Reg"
        ElseIf .Text = "Name" Then
            fieldName2 = "Lname"
        End If
    End With
End Sub


Private Sub AssignDueVal()

Dim tmpDaysPast As Integer
Dim counter As Integer

    lblDueInfo(0).Caption = AdoDue.Recordset.Fields("Borrower_Id")
    lblDueInfo(1).Caption = UCase(AdoDue.Recordset.Fields("Borrower_LName") & ", " & AdoDue.Recordset.Fields("Borrower_FName"))
    lblDueInfo(2).Caption = AdoDue.Recordset.Fields("Book_Id")
    lblDueInfo(3).Caption = AdoDue.Recordset.Fields("Book_Title")
    lblDueInfo(4).Caption = AdoDue.Recordset.Fields("Date_Borrowed")
    lblDueInfo(5).Caption = AdoDue.Recordset.Fields("Date_Due")
    If DateDiff("d", AdoDue.Recordset.Fields("Date_Due"), Date) <= 0 Then
            tmpDaysPast = 0
            
        Else
            tmpDaysPast = DateDiff("d", AdoDue.Recordset.Fields("Date_Due"), Date)
            
        End If
    lblDueInfo(6).Caption = tmpDaysPast

For counter = 0 To 4
    If Trim(lblDueInfo(counter).Caption) = "" Then
        lblDueInfo(counter).Caption = "--"
    End If
Next counter
End Sub

Public Sub AssignBorVal()
On Error Resume Next
    imgPic.Picture = LoadPicture("")
    lblBorInfo(1).Caption = AdoBor.Recordset.Fields("BorId")
    lblBorInfo(2).Caption = AdoBor.Recordset.Fields("Lname") & ", " & AdoBor.Recordset.Fields("Fname") & " " & AdoBor.Recordset.Fields("Mname")
    lblBorInfo(3).Caption = AdoBor.Recordset.Fields("Course")
    lblBorInfo(4).Caption = AdoBor.Recordset.Fields("Contact")
    lblBorInfo(5).Caption = AdoBor.Recordset.Fields("Date_Reg")
    imgPic.Picture = LoadPicture(AdoBor.Recordset.Fields("Pic"))
End Sub

Public Sub BorGrid()
Set dtgBorGrid.DataSource = AdoBor

    With dtgBorGrid
        .Columns(0).DataField = "Bor_Id"
        .Columns(0).Caption = "Borrower ID"
        .Columns(0).Width = 1300
        
        .Columns(1).DataField = "Fname"
        .Columns(1).Caption = "First Name"
        .Columns(1).Width = 1500
        
        .Columns(2).DataField = "Mname"
        .Columns(2).Caption = "Middle Name"
        .Columns(2).Width = 1300
        
        .Columns(3).DataField = "Lname"
        .Columns(3).Caption = "Last Name"
        .Columns(3).Width = 1300
        
        .Columns(4).DataField = "Course"
        .Columns(4).Caption = "Course"
        .Columns(4).Width = 800
        
        .Columns(5).DataField = "Contact"
        .Columns(5).Caption = "Contact Number"
        .Columns(5).Width = 1500
        
        .Columns(6).DataField = "Date_Reg"
        .Columns(6).Caption = "Date Registered"
        .Columns(6).Width = 1600
        
        .Columns(7).DataField = "Pic"
        .Columns(7).Caption = "Pic"
        .Columns(7).Width = 0
        
     End With
End Sub



Private Sub txtBookSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearchBook_Click
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
End Sub


Private Sub SetBookGrid()

Set dtgBookList.DataSource = AdoBooks
    With dtgBookList
        .Columns(0).DataField = "BookId"
        .Columns(0).Caption = "Book Number"
        .Columns(0).Width = 1280
        
        .Columns(1).DataField = "CallId"
        .Columns(1).Caption = "Call Number"
        .Columns(1).Width = 1280
        
        .Columns(2).DataField = "Title"
        .Columns(2).Caption = "Title"
        .Columns(2).Width = 2000
        
        .Columns(3).DataField = "Author"
        .Columns(3).Caption = "Author"
        .Columns(3).Width = 1500
        
        .Columns(4).DataField = "Publisher"
        .Columns(4).Caption = "Publisher"
        .Columns(4).Width = 1500
        
        .Columns(5).DataField = "DateReg"
        .Columns(5).Caption = "Date Registered"
        .Columns(5).Width = 1500
     End With
End Sub

Private Sub AssignBookVal()
'On Error Resume Next
    lblBookInfo(1).Caption = AdoBooks.Recordset.Fields("BookId")
    lblBookInfo(2).Caption = AdoBooks.Recordset.Fields("CallId")
    lblBookInfo(3).Caption = AdoBooks.Recordset.Fields("Title")
    lblBookInfo(4).Caption = AdoBooks.Recordset.Fields("Author")
    lblBookInfo(5).Caption = AdoBooks.Recordset.Fields("Publisher")
    lblBookInfo(6).Caption = AdoBooks.Recordset.Fields("DateReg")
End Sub

Private Sub ConvertToFieldBooks()
    'for cmborder
    With cmbBooks
        If .Text = "Author" Then
            fieldName3 = "Author"
            
        ElseIf .Text = "Book Number" Then
            fieldName3 = "BookId"
            
        ElseIf .Text = "Call Number" Then
            fieldName3 = "Book.CallId"
        
        ElseIf .Text = "Publisher" Then
            fieldName3 = "Publisher"
        
        ElseIf .Text = "Date Registered" Then
            fieldName3 = "DateReg"
        End If
    End With
End Sub
