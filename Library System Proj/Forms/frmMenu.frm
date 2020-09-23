VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   1110
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"frmMenu.frx":0000
      Height          =   810
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4500
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New"
      Begin VB.Menu MnuNewBor 
         Caption         =   "New Borrower Profile"
      End
      Begin VB.Menu mnuNewLib 
         Caption         =   "New Librarian Profile"
      End
      Begin VB.Menu mnuNewTitle 
         Caption         =   "New Title Entry"
      End
      Begin VB.Menu mnuNewBook 
         Caption         =   "New Book Entry"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditBor 
         Caption         =   "Edit Borrower Profile"
      End
      Begin VB.Menu mnuEditLib 
         Caption         =   "Edit Librarian Profile"
      End
      Begin VB.Menu mnuEditTitle 
         Caption         =   "Edit Title Entry"
      End
      Begin VB.Menu mnuEditBook 
         Caption         =   "Edit Book Entry"
      End
   End
   Begin VB.Menu mnuDel 
      Caption         =   "Delete"
      Begin VB.Menu mnuDelBor 
         Caption         =   "Delete Borrower Profile"
      End
      Begin VB.Menu mnuDelLib 
         Caption         =   "Delete Librarian Profile"
      End
      Begin VB.Menu mnuDelTitle 
         Caption         =   "Delete Title Entry"
      End
      Begin VB.Menu mnuDelBook 
         Caption         =   "Delete Book Entry"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuDelBook_Click()
On Error Resume Next
    Load frmBooks
    frmBooks.tabBooks.Tab = 2
    frmBooks.txtIdDel.SetFocus
    frmBooks.Show vbModal
End Sub

Private Sub mnuDelBor_Click()
On Error Resume Next
    
    frmDelBor.Show vbModal
End Sub

Private Sub mnuDelLib_Click()
On Error Resume Next
    
    frmLibDel.Show vbModal
End Sub

Private Sub mnuDelTitle_Click()
On Error Resume Next
    Load frmTitle
    frmTitle.tabTitle.Tab = 2
    frmTitle.txtIdDel.SetFocus
    frmTitle.Show vbModal
End Sub

Private Sub mnuEditBook_Click()
On Error Resume Next
    Load frmBooks
    frmBooks.tabBooks.Tab = 1
    frmBooks.txtEdBookID.SetFocus
    frmBooks.Show vbModal
End Sub

Private Sub mnuEditBor_Click()
On Error Resume Next
    
    frmEditBor.Show vbModal
End Sub



Private Sub mnuEditLib_Click()
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

Private Sub mnuNewBook_Click()
On Error Resume Next
    Load frmBooks
    frmBooks.tabBooks.Tab = 0
    frmBooks.txtBookID.SetFocus
    frmBooks.Show vbModal
End Sub

Private Sub MnuNewBor_Click()
On Error Resume Next
    
    frmAddStud.Show vbModal
End Sub

Private Sub mnuNewLib_Click()
On Error Resume Next
    
    frmAdminSetup.Show vbModal
End Sub

Private Sub mnuNewTitle_Click()
On Error Resume Next
    Load frmTitle
    frmTitle.tabTitle.Tab = 0
    frmTitle.txtCallID.SetFocus
    frmTitle.Show vbModal
End Sub
