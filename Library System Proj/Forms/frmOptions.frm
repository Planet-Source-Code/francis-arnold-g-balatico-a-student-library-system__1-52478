VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library options"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2910
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoOptions 
      Height          =   330
      Left            =   30
      Top             =   2760
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
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   23
      TabIndex        =   1
      Top             =   855
      Width           =   2865
      Begin VB.TextBox txtMaxBooks 
         Height          =   315
         Left            =   1290
         TabIndex        =   4
         Top             =   705
         Width           =   1170
      End
      Begin VB.TextBox txtFines 
         Height          =   315
         Left            =   1290
         TabIndex        =   2
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max Books:"
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
         Left            =   255
         TabIndex        =   5
         Top             =   750
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fines:"
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
         Left            =   735
         TabIndex        =   3
         Top             =   315
         Width           =   525
      End
   End
   Begin Project1.lvButtons_H cmdSave 
      Height          =   405
      Left            =   135
      TabIndex        =   6
      Top             =   2220
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Ok"
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
   Begin Project1.lvButtons_H lvButtons_H1 
      Height          =   405
      Left            =   1485
      TabIndex        =   7
      Top             =   2220
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Cancel"
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
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   555
      Left            =   45
      Picture         =   "frmOptions.frx":0000
      Stretch         =   -1  'True
      Top             =   2145
      Width           =   2835
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup library options"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   825
      TabIndex        =   0
      Top             =   300
      Width           =   1560
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   210
      Picture         =   "frmOptions.frx":2987
      Top             =   180
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmOptions.frx":3651
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2910
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
On Error Resume Next
    If Trim(txtFines.Text) = "" Then
        MsgBox "Required information missing.", vbOKOnly + vbExclamation, "Library System"
        txtFines.SetFocus
        Exit Sub
    End If
    
    If Trim(txtMaxBooks.Text) = "" Then
        MsgBox "Required information missing.", vbOKOnly + vbExclamation, "Library System"
        txtMaxBooks.SetFocus
        Exit Sub
    End If
    
    AdoOptions.Refresh
    AdoOptions.Recordset.MoveFirst
    AdoOptions.Recordset.Fields("Fine") = Trim(txtFines.Text)
    AdoOptions.Recordset.Fields("MaxBooks") = Trim(txtMaxBooks.Text)
    AdoOptions.Recordset.Update
    AdoOptions.Refresh
    
    Fines = AdoOptions.Recordset.Fields("Fine")
    MaxBooks = AdoOptions.Recordset.Fields("MaxBooks")
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call ConnectToDb(AdoOptions, "Setup")
    txtFines.Text = AdoOptions.Recordset.Fields("Fine")
    txtMaxBooks.Text = AdoOptions.Recordset.Fields("MaxBooks")
    
End Sub

Private Sub lvButtons_H1_Click()
    Unload Me
End Sub

Private Sub txtFines_Change()
    If IsNumeric(txtFines.Text) = False Then
        MsgBox "Cannot accept non-numeric input", vbOKOnly + vbExclamation, "Library System"
        SendKeys HiLyt
    Else
        Exit Sub
    End If
End Sub

Private Sub txtFines_GotFocus()
    SendKeys HiLyt
End Sub

Private Sub txtMaxBooks_Change()
    If IsNumeric(txtMaxBooks.Text) = False Then
        MsgBox "Cannot accept non-numeric input", vbOKOnly + vbExclamation, "Library System"
        SendKeys HiLyt
    Else
        Exit Sub
    End If
End Sub

Private Sub txtMaxBooks_GotFocus()
    SendKeys HiLyt
End Sub
