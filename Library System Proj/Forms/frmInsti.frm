VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInsti 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institution"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoInsti 
      Height          =   330
      Left            =   120
      Top             =   1725
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.TextBox txtInsti 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1245
      Width           =   4185
   End
   Begin Project1.lvButtons_H cmdSave 
      Height          =   435
      Left            =   2970
      TabIndex        =   2
      Top             =   1680
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   767
      Caption         =   "&OK"
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
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   435
      Picture         =   "frmInsti.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the institution/organization's name. Leave blank to use default setting."
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1110
      TabIndex        =   3
      Top             =   225
      Width           =   2985
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmInsti.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4410
   End
   Begin VB.Label lblInsti 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter &Institution/Organization Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   0
      Top             =   990
      Width           =   2955
   End
End
Attribute VB_Name = "frmInsti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    Call ReAssign
End Sub

Private Sub cmdSave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ReAssign
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Call ConnectToDb(AdoInsti, "Setup")
    
        
    AdoInsti.Refresh
    AdoInsti.Recordset.MoveFirst
    txtInsti.Text = AdoInsti.Recordset.Fields("Institution")
    Exit Sub
    
ErrHandler:
    AdoInsti.Recordset.AddNew

End Sub



Private Sub txtInsti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSave_Click
    End If
End Sub

Public Sub ReAssign()
 If Trim(txtInsti.Text) = "" Then 'if blank is specified
        If MsgBox("You did not specify an institution name. Use default?", vbYesNo + vbQuestion, "Library System") = vbYes Then
            AdoInsti.Recordset.Fields("Institution") = "Student Library System"
            AdoInsti.Recordset.Update
            AdoInsti.Refresh
            AdoInsti.Refresh
            AdoInsti.Recordset.MoveFirst
            LibInsti = Trim(txtInsti.Text)
            
            If Main_On = True Then 'if change is initiated inside MDI form
                MDIMain.lblCompany(0).Caption = "Student Library System"
                MDIMain.lblCompany(1).Caption = "Student Library System"
                
                MDIMain.cmdCompany.Width = MDIMain.lblCompany(0).Width + (285 * 2)
                MDIMain.lblCompany(0).Left = 285
                
                MDIMain.lblCompany(1).Left = MDIMain.lblCompany(0).Left + 15
                MDIMain.lblCompany(1).Top = MDIMain.lblCompany(0).Top + 15
                
                MDIMain.cmdCompany.Left = MDIMain.tlbLib.Width - (MDIMain.cmdCompany.Width + 80) 'sets the company name's position
                Main_On = True
            End If
            Unload Me
            Exit Sub
        Else
            Exit Sub
        End If
        
    End If
    
    
    If MsgBox("Institution name will be '" & Trim(txtInsti.Text) & "'. Proceed?", vbYesNo + vbQuestion, "Library System") = vbYes Then
        AdoInsti.Recordset.Fields("Institution") = Trim(txtInsti.Text)
        AdoInsti.Recordset.Update
        AdoInsti.Refresh
        AdoInsti.Refresh
        AdoInsti.Recordset.MoveFirst
        LibInsti = Trim(txtInsti.Text)
               
        If Main_On = True Then 'if change is initiated inside MDI form
            MDIMain.lblCompany(0).Caption = LibInsti
            MDIMain.lblCompany(1).Caption = LibInsti
            
            MDIMain.cmdCompany.Width = MDIMain.lblCompany(0).Width + (285 * 2)
            MDIMain.lblCompany(0).Left = 285
            
            MDIMain.lblCompany(1).Left = MDIMain.lblCompany(0).Left + 15
            MDIMain.lblCompany(1).Top = MDIMain.lblCompany(0).Top + 15
            
            MDIMain.cmdCompany.Left = MDIMain.tlbLib.Width - (MDIMain.cmdCompany.Width + 80) 'sets the company name's position
            Main_On = True
        End If
        Unload Me
        
    Else
        Exit Sub
    End If
End Sub
