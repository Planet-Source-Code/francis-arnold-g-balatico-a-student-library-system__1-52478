VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000F&
   Caption         =   "Student Library System"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   6615
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIMain.frx":0CCA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlbLib 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList"
      DisabledImageList=   "ImgList"
      HotImageList    =   "ImgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn1"
            Object.ToolTipText     =   "Transaction"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn3"
            Object.ToolTipText     =   "New"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn4"
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn5"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn7"
            Object.ToolTipText     =   "Records"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn9"
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame cmdCompany 
         Height          =   570
         Left            =   4005
         TabIndex        =   2
         ToolTipText     =   "Institution Name"
         Top             =   -45
         Visible         =   0   'False
         Width           =   2550
         Begin VB.Label lblCompany 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student Library System"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Index           =   0
            Left            =   330
            TabIndex        =   3
            ToolTipText     =   "Institution Name"
            Top             =   225
            Width           =   1950
         End
         Begin VB.Label lblCompany 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student Library System"
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
            Left            =   345
            TabIndex        =   4
            Top             =   240
            Width           =   1950
         End
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   30
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":52E2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":53B07
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":547E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":554C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5619F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":56E7B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   1305
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "3\22\2004"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "07:53 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuLib 
      Caption         =   "&Library"
      Begin VB.Menu mnuBorRet 
         Caption         =   "&Transaction..."
      End
      Begin VB.Menu mnuHyphen 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuProfiles 
         Caption         =   "&Profile Manager"
         Begin VB.Menu mnuStudNew 
            Caption         =   "&Borrower Profile"
            Begin VB.Menu mnuBorNew 
               Caption         =   "&New borrower profile..."
            End
            Begin VB.Menu mnuStudEdit 
               Caption         =   "&Edit existing borrower profile..."
            End
            Begin VB.Menu mnuBorDel 
               Caption         =   "&Delete existing borrower profile..."
            End
         End
         Begin VB.Menu mnuLibAccnt 
            Caption         =   "&Librarian Profile"
            Begin VB.Menu mnuCreate 
               Caption         =   "&New librarian profile..."
            End
            Begin VB.Menu mnuEditExist 
               Caption         =   "&Edit existing librarian profile..."
            End
            Begin VB.Menu mnuDelProf 
               Caption         =   "&Delete existing librarian profile..."
            End
         End
         Begin VB.Menu mnuBook 
            Caption         =   "&Book Profile"
            Begin VB.Menu mnuTitle 
               Caption         =   "&Title entry"
               Begin VB.Menu mnuNewTitle 
                  Caption         =   "&New title entry"
               End
               Begin VB.Menu mnuEditTitle 
                  Caption         =   "&Edit title entry"
               End
               Begin VB.Menu mnuDelTitle 
                  Caption         =   "&Delete title entry"
               End
            End
            Begin VB.Menu mnuBookEnt 
               Caption         =   "&Book entry"
               Begin VB.Menu mnuNewBook 
                  Caption         =   "&New book entry..."
               End
               Begin VB.Menu mnuBookEdit 
                  Caption         =   "&Edit book entry..."
               End
               Begin VB.Menu mnuDelBook 
                  Caption         =   "&Delete book entry..."
               End
            End
         End
      End
      Begin VB.Menu mnuRecDraw 
         Caption         =   "&Record Drawer"
         Begin VB.Menu mnuSearchStud 
            Caption         =   "&Student Record..."
         End
         Begin VB.Menu mnuBookSearch 
            Caption         =   "&Book Record..."
         End
         Begin VB.Menu mnuDue 
            Caption         =   "&Due Books..."
         End
      End
      Begin VB.Menu mnuHyphen0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReps 
         Caption         =   "&Print Reports"
      End
      Begin VB.Menu mnuHyp0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "&Log out"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuLock 
         Caption         =   "&Lock library..."
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "&Calculator..."
      End
      Begin VB.Menu mnuHyp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Hide Toolbar"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "S&etup"
      Begin VB.Menu mnuAdmin 
         Caption         =   "&Institution Name..."
      End
      Begin VB.Menu mnuHyp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuManual 
         Caption         =   "&System Req..."
      End
      Begin VB.Menu mnuHyph3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "&Credits..."
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private log_off As Boolean


Private Sub MDIForm_Load()
On Error Resume Next
    Main_On = True
    cmdCompany.Visible = False 'hides the company
    
    CenterFrm MDIMain, frmInsignia
    frmInsignia.Show
    CenterFrm MDIMain, frmInsignia 'this ensures the form is centered
    
    With stbMain 'ensures all panels are visible upon loading
        .Panels(1).Width = MDIMain.Width - (.Panels(2).Width + .Panels(3).Width + .Panels(4).Width + .Panels(5).Width + .Panels(6).Width) 'maintains the status bar's first panel width
        .Panels(1).Text = "Ready"
    End With
    
    
    MDIMain.lblCompany(0).Caption = LibInsti
    MDIMain.lblCompany(1).Caption = LibInsti
    
    cmdCompany.Width = lblCompany(0).Width + (285 * 2)
    lblCompany(0).Left = 285
    
    lblCompany(1).Left = lblCompany(0).Left + 15
    lblCompany(1).Top = lblCompany(0).Top + 15
    
    cmdCompany.Left = tlbLib.Width - (cmdCompany.Width + 80) 'sets the company name's position
    cmdCompany.Visible = True 'displays the company
    
    
        DataEnvironment1.connLib.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppDir & "\Lib_Dbase.mdb;Persist Security Info=False; Jet OLEDB:Database Password = crimson119"
   
    
    'false if the user has not logged out
    log_off = False

End Sub
Private Sub MDIForm_Resize()
On Error Resume Next
    CenterFrm MDIMain, frmInsignia
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
'if-then-else statement that determines if the unload is generated by closing the program or by logging out
If log_off = False Then 'unload is from termination of form
        If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Library System") = vbYes Then
            End
        Else
            Cancel = 1
            CenterFrm MDIMain, frmInsignia
            frmInsignia.Show
            stbMain.SimpleText = "Ready"
            CenterFrm MDIMain, frmInsignia
        End If
Else 'unload is from logging out
        If MsgBox("Are you sure you want to log out?", vbQuestion + vbYesNo, "Library System") = vbYes Then
            log_off = False
            frmLogin.Show
            Unload MDIMain
        Else
            Cancel = 1
            log_off = False
            CenterFrm MDIMain, frmInsignia
            frmInsignia.Show
            stbMain.SimpleText = "Ready"
            CenterFrm MDIMain, frmInsignia
        End If
End If
End Sub

Private Sub mnuAdmin_Click()
    frmInsti.Show vbModal
End Sub

Private Sub mnuBookEdit_Click()
On Error Resume Next
    Load frmBooks
    frmBooks.tabBooks.Tab = 1
    frmBooks.txtEdBookID.SetFocus
    frmBooks.Show vbModal
End Sub

Private Sub mnuBookSearch_Click()
    Load frmRecs
    frmRecs.tabRecs.Tab = 1
    frmRecs.Show vbModal
End Sub

Private Sub mnuBorDel_Click()
    On Error Resume Next
    frmDelBor.Show vbModal
End Sub

Private Sub mnuBorNew_Click()
    On Error Resume Next
    frmAddStud.Show vbModal
End Sub

Private Sub mnuBorRet_Click()
    On Error Resume Next
    frmStudProf.Show vbModal
End Sub

Private Sub mnuCalc_Click()
    On Error Resume Next
    Shell "Calc.exe", vbMaximizedFocus
End Sub

Private Sub mnuCreate_Click()
    On Error Resume Next
    frmAdminSetup.Show vbModal
End Sub

Private Sub mnuCredits_Click()
    On Error Resume Next
    
    frmCredits.Show vbModal
End Sub

Private Sub mnuDelBook_Click()
On Error Resume Next
    Load frmBooks
    frmBooks.tabBooks.Tab = 2
    frmBooks.txtIdDel.SetFocus
    frmBooks.Show vbModal
End Sub

Private Sub mnuDelProf_Click()
    On Error Resume Next
    frmLibDel.Show vbModal
End Sub

Private Sub mnuDelTitle_Click()
On Error Resume Next
    Load frmTitle
    frmTitle.tabTitle.Tab = 3
    frmTitle.txtIdDel.SetFocus
    frmTitle.Show vbModal
End Sub

Private Sub mnuDue_Click()
    On Error Resume Next
    Load frmRecs
    frmRecs.tabRecs.Tab = 2
    frmRecs.Show vbModal
End Sub

Private Sub mnuEditExist_Click()
On Error Resume Next
    frmLibEdit.Show vbModal
End Sub

Private Sub mnuEditTitle_Click()
On Error Resume Next
    Load frmTitle
    frmTitle.tabTitle.Tab = 1
    frmTitle.txtEdCallID.SetFocus
    frmTitle.Show vbModal
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Library System") = vbYes Then
    End
Else
'Cancel operation
    Exit Sub
End If
End Sub

Private Sub mnuLock_Click()
    frmLock.Show vbModal
End Sub

Private Sub mnuLogout_Click()
    On Error Resume Next
    log_off = True
    
    Unload MDIMain
End Sub

Private Sub mnuManual_Click()
    MsgBox "System Requirements:" & vbCrLf & "Standard Keyboard and Mouse" & vbCrLf & "32 MB RAM" & vbCrLf & "Pentium 3 processor or similar" & vbCrLf & "Win 98 or higher", vbOKOnly, "Library System"
End Sub

Private Sub mnuNewBook_Click()
On Error Resume Next
    Load frmBooks
    frmBooks.tabBooks.Tab = 0
    frmBooks.txtBookID.SetFocus
    frmBooks.Show vbModal
End Sub

Private Sub mnuNewTitle_Click()
On Error Resume Next
    Load frmTitle
    frmTitle.tabTitle.Tab = 0
    frmTitle.txtCallID.SetFocus
    frmTitle.Show vbModal
End Sub



Private Sub mnuOptions_Click()
On Error Resume Next
    frmOptions.Show vbModal
End Sub

Private Sub mnuReps_Click()
    frmPrint.Show vbModal
End Sub

Private Sub mnuSearchStud_Click()
On Error Resume Next
    Load frmRecs
    frmRecs.tabRecs.Tab = 0
    frmRecs.Show vbModal
End Sub

Private Sub mnuStudEdit_Click()
On Error Resume Next
    frmEditBor.Show vbModal
End Sub

Private Sub mnuTool_Click()
On Error Resume Next
    If tlbLib.Visible = True Then
        tlbLib.Visible = False
        mnuTool.Caption = "&Show Toolbar"
        
    Else
        tlbLib.Visible = True
        mnuTool.Caption = "&Hide Toolbar"
       
    End If

End Sub

Private Sub tlbLib_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Dim Btn As String

Btn = Button.Key

Select Case Btn
    Case "btn1"
        frmStudProf.Show vbModal
    Case "btn3"
        PopupMenu frmMenu.mnuNew, , Button.Left, (Button.Top + Button.Height)
    Case "btn4"
        PopupMenu frmMenu.mnuEdit, , Button.Left, (Button.Top + Button.Height)
    Case "btn5"
        PopupMenu frmMenu.mnuDel, , Button.Left, (Button.Top + Button.Height)
    Case "btn7"
        Load frmRecs
        frmRecs.tabRecs.Tab = 0
        frmRecs.Show vbModal
    Case "btn9"
        frmPrint.Show vbModal
End Select
End Sub


