VERSION 5.00
Begin VB.Form frmLock 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   6045
      Begin VB.TextBox txtPass 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   285
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2070
         Width           =   5565
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2835
         Top             =   3420
      End
      Begin Project1.lvButtons_H cmdUnlock 
         Height          =   420
         Left            =   1995
         TabIndex        =   2
         Top             =   2745
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   741
         Caption         =   "&Unlock"
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
      Begin VB.Shape Shape1 
         Height          =   915
         Left            =   225
         Top             =   1665
         Width           =   5685
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER PASSWORD TO UNLOCK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D3F2D&
         Height          =   240
         Left            =   1305
         TabIndex        =   4
         Top             =   1725
         Width           =   3525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOCKED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B5742D&
         Height          =   555
         Left            =   2025
         TabIndex        =   1
         Top             =   555
         Width           =   2070
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   285
         Picture         =   "frmLock.frx":0000
         Stretch         =   -1  'True
         Top             =   405
         Width           =   5565
      End
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUnlock_Click()
    If txtPass.Text = LibPass Then
        Unload Me
    Else
        MsgBox "Wrong password supplied. Attempt to unlock failed.", vbOKOnly + vbExclamation, "Library System"
        SendKeys HiLyt
        Exit Sub
    End If
End Sub


Private Sub Timer1_Timer()
    If Trim(txtPass.Text) = "" Then
        cmdUnlock.Enabled = False
    Else
        cmdUnlock.Enabled = True
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdUnlock_Click
    End If
End Sub
