VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDelBor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Borrower Profile"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   ControlBox      =   0   'False
   Icon            =   "frmDelBor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoRec 
      Height          =   330
      Left            =   1350
      Top             =   5025
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
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2595
      Top             =   4950
   End
   Begin MSAdodcLib.Adodc AdoDel 
      Height          =   330
      Left            =   120
      Top             =   5025
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "LIBRARIAN PASSWORD CONFIRMATION"
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
      Height          =   690
      Left            =   60
      TabIndex        =   4
      Top             =   4170
      Width           =   6855
      Begin VB.CheckBox chkRem 
         Caption         =   "&Remember"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5310
         TabIndex        =   9
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2220
         PasswordChar    =   "v"
         TabIndex        =   8
         Top             =   210
         Width           =   2895
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enter &Password"
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
         Left            =   465
         TabIndex        =   7
         Top             =   285
         Width           =   1635
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
      Height          =   705
      Left            =   60
      TabIndex        =   0
      Top             =   945
      Width           =   6855
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
         Left            =   2220
         TabIndex        =   16
         Top             =   225
         Width           =   2895
      End
      Begin Project1.lvButtons_H cmdSearch 
         Height          =   435
         Left            =   5340
         TabIndex        =   1
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
         Image           =   "frmDelBor.frx":0CCA
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
         TabIndex        =   2
         Top             =   300
         Width           =   1830
      End
   End
   Begin Project1.lvButtons_H cmdClose 
      Height          =   405
      Left            =   5535
      TabIndex        =   5
      Top             =   4965
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
      Left            =   4140
      TabIndex        =   6
      Top             =   4965
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
      Image           =   "frmDelBor.frx":19A4
      cBack           =   12632256
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "BORROWER PROFILE"
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
      Height          =   2415
      Left            =   60
      TabIndex        =   3
      Top             =   1710
      Width           =   6855
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00B5742D&
         Height          =   360
         Left            =   60
         ScaleHeight     =   300
         ScaleWidth      =   6675
         TabIndex        =   11
         Top             =   2010
         Width           =   6735
         Begin VB.Label lblCurrent 
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
            Left            =   5310
            TabIndex        =   15
            Top             =   45
            Width           =   120
         End
         Begin VB.Label lblAvail 
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
            Left            =   1695
            TabIndex        =   14
            Top             =   45
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current record:"
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
            Left            =   3975
            TabIndex        =   13
            Top             =   45
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available records:"
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
            Left            =   105
            TabIndex        =   12
            Top             =   45
            Width           =   1545
         End
      End
      Begin MSDataGridLib.DataGrid dtaDel 
         Bindings        =   "frmDelBor.frx":267E
         Height          =   1815
         Left            =   60
         TabIndex        =   10
         Top             =   195
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3201
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
            AllowFocus      =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   345
      Picture         =   "frmDelBor.frx":2693
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Select or Search for the Borrower Profile to delete.  Deleting requires Librarian password confirmation to continue."
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1035
      TabIndex        =   17
      Top             =   240
      Width           =   5655
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmDelBor.frx":335D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6990
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   60
      Picture         =   "frmDelBor.frx":63FD
      Stretch         =   -1  'True
      Top             =   4890
      Width           =   6855
   End
End
Attribute VB_Name = "frmDelBor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    If Val(AdoDel.Recordset.RecordCount) = 0 Then
        MsgBox "There are no more records to delete.", vbOKOnly + vbExclamation, "Library System"
    End If
    
    If Trim(txtPass.Text) = "" Then 'validates that there is a password entry
        MsgBox "Please specify the CURRENT LIBRARIAN PASSWORD to authorize deletion.", vbOKOnly + vbExclamation, "Library System"
        Exit Sub
    End If
            
    If txtPass.Text = LibPass Then
        If ChkAcount(AdoDel.Recordset.Fields("BorID")) = True Then
            MsgBox "Cannot delete record because the borrower has unsettled accounts.", vbOKOnly + vbExclamation, "Library System"
            Exit Sub
        End If
        
        If MsgBox("Password confirmed. Delete record?", vbOKCancel + vbQuestion, "Library System") = vbOK Then
            
            Call Status("Deleting record...")
                                  
            AdoDel.Recordset.Delete
            txtSearch.Text = ""
            
            If chkRem.Value = 0 Then
                txtPass.Text = ""
            End If
            
            Call Status("Ready")
            
            Exit Sub
        Else
            Exit Sub
        End If
    Else
        MsgBox "Password invalid. Deletion cancelled.", vbOKOnly + vbExclamation, "Library System"
        txtPass.SetFocus
        SendKeys HiLyt
        Exit Sub
    End If

End Sub

Private Sub cmdSearch_Click() 'search the item to delete
    Dim counter As Integer
    
        txtSearch.Text = Trim(txtSearch.Text)
        If Trim(txtSearch.Text) = "" Then
            Exit Sub
        End If
    
    AdoDel.Refresh
    
    AdoDel.Recordset.Find ("BorId = '" & UCase(txtSearch.Text) & "'")
    
    If Not AdoDel.Recordset.BOF = True And AdoDel.Recordset.EOF = True Then
        Call FormatData
        MsgBox "Record not found", vbOKOnly + vbExclamation, "Library System"
        Exit Sub
    End If
    
    Call FormatData



End Sub



Private Sub Form_Load() 'loads the SQL for the Datagrid
    On Error GoTo ErrHandler
        
        Call Status("Loading database...")
                
        Call SQLDB(AdoDel, "Select BorID, FName, LName, Course, Contact, Mname, Pic from Borrower")
        AdoDel.Refresh
        
        AdoDel.Recordset.MoveFirst 'generates an error if there are no records
        
        Call Status("Populating datagrid...")
        dtaDel.Refresh
        
        Call FormatData
        Call Status("Ready")
        Exit Sub
        
ErrHandler:
    Call NoRec(Me)

End Sub

Public Sub FormatData() 'sets the formatting of the datagrid
    Dim counter As Integer
    
With dtaDel
        .BackColor = &HE28B2C
        .Columns(0).DataField = "BorID"
        .Columns(0).Caption = "ID"
        .Columns(0).Width = 1000
                
        .Columns(1).DataField = "FName"
        .Columns(1).Caption = "First Name"
        .Columns(1).Width = 1475
        
        .Columns(2).DataField = "Lname"
        .Columns(2).Caption = "Last Name"
        .Columns(2).Width = 1400
        
        .Columns(3).DataField = "Course"
        .Columns(3).Caption = "Course"
        .Columns(3).Width = 770
        .Columns(3).Alignment = dbgCenter
        
        .Columns(4).DataField = "Contact"
        .Columns(4).Caption = "Contact Number"
        .Columns(4).Width = 1470
        .Columns(4).Alignment = dbgCenter
        
        .Columns(5).Visible = False
        .Columns(6).Visible = False
        
        .AllowUpdate = False
        .HeadFont.Bold = True
        .ScrollBars = dbgVertical
        
           
        .Splits(0).MarqueeStyle = dbgHighlightRow
        .Splits(0).Locked = True
        .Splits(0).AllowRowSizing = False
        .Splits(0).AllowFocus = False
        
        For counter = 1 To 5
            .Columns(counter).AllowSizing = False
        Next counter
    End With

End Sub

Private Sub Timer1_Timer()
    lblAvail.Caption = AdoDel.Recordset.RecordCount
    
    If AdoDel.Recordset.RecordCount = 0 Or AdoDel.Recordset.AbsolutePosition < adPosBOF Then
        lblCurrent.Caption = "--"
    Else
        lblCurrent.Caption = AdoDel.Recordset.AbsolutePosition
    End If
    
    If AdoDel.Recordset.RecordCount = 0 Then 'no record available
        cmdDel.Enabled = False
        cmdSearch.Enabled = False
        txtSearch.Enabled = False
        txtPass.Text = ""
        txtPass.Enabled = False
        chkRem.Value = 0
        chkRem.Enabled = False
    Else 'there are existing records
        cmdDel.Enabled = True
        cmdSearch.Enabled = True
        txtSearch.Enabled = True
        txtPass.Enabled = True
        chkRem.Enabled = True
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSearch_Click
    End If
End Sub
