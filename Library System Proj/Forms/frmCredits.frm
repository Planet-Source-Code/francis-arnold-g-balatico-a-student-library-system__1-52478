VERSION 5.00
Begin VB.Form frmCredits 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Developed by:"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   570
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   3240
   End
   Begin Project1.lvButtons_H cmdClose 
      Height          =   435
      Left            =   3720
      TabIndex        =   2
      Top             =   3255
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   767
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
      cBhover         =   11891757
      cGradient       =   11891757
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Credits"
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4890
      Begin VB.PictureBox picCredits 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   1170
         Left            =   90
         ScaleHeight     =   74
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   3
         Top             =   240
         Width           =   4725
         Begin VB.TextBox txtCredits 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            Height          =   2655
            Left            =   225
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "frmCredits.frx":0000
            Top             =   1170
            Width           =   4170
         End
      End
   End
   Begin VB.PictureBox picDev 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   135
      Picture         =   "frmCredits.frx":0162
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   90
      Width           =   1560
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "For your queries, comments and suggestions, please E-mail me at the given E-mail Address."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1845
      TabIndex        =   5
      Top             =   945
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FRANCIS ARNOLD G. BALATICO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   1965
      TabIndex        =   7
      Top             =   240
      Width           =   3045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crimson_Zenith@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   2295
      TabIndex        =   6
      Top             =   480
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FRANCIS ARNOLD G. BALATICO"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   1980
      TabIndex        =   8
      Top             =   255
      Width           =   3045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Crimson_Zenith@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2310
      TabIndex        =   9
      Top             =   495
      Width           =   2715
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If txtCredits.Top > 0 - (txtCredits.Height) Then
    txtCredits.Top = txtCredits.Top - 1
Else
    txtCredits.Visible = False
    Timer2.Enabled = True
    Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
    txtCredits.Top = 78
    txtCredits.Visible = True
    Timer1.Enabled = True
    Timer2.Enabled = False
End Sub
