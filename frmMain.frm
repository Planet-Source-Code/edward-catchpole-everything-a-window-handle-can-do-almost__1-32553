VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Everything Windows"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetHandle 
      Caption         =   "Get Handle"
      Height          =   615
      Left            =   1680
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame fraWin 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtNewCaption 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "Type new caption here"
         Top             =   2520
         Width           =   4335
      End
      Begin VB.CommandButton cmdFront 
         Caption         =   "Bring To Front"
         Height          =   615
         Left            =   3240
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdBottom 
         Caption         =   "Send To Back"
         Height          =   615
         Left            =   1680
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdSetActive 
         Caption         =   "Make Active"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "Move/Size Window"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdSetCaption 
         Caption         =   "Set Caption"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdGetCaption 
         Caption         =   "Get Caption"
         Height          =   615
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFlashTitle 
         Caption         =   "Flash Title Bar"
         Height          =   615
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDestroy 
         Caption         =   "Destroy"
         Height          =   615
         Left            =   3240
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chosen Window"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   4455
      Begin VB.Label lblCaption 
         Caption         =   "Caption"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Timer timSetText 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4200
      Top             =   3360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WnHnd   As Long

Private Sub cmdBottom_Click()

    Call SendToBottom(WnHnd)
    
End Sub

Private Sub cmdDestroy_Click()

    Call DestroyWindow(WnHnd)
    
End Sub

Private Sub cmdFlashTitle_Click()

    Call FlashWindow(WnHnd, 1)
    
End Sub

Private Sub cmdFront_Click()

    Call BringToFront(WnHnd)
    
End Sub

Private Sub cmdGetCaption_Click()

    lblCaption.Caption = GetWindowCaption(WnHnd)
    
End Sub

Private Sub cmdGetHandle_Click()

    MsgBox "You have 5 seconds to switch to a different window to get its handle", vbSystemModal
    
    timSetText.Enabled = True
    
End Sub

Private Sub cmdMove_Click()

    Call MoveWindow(WnHnd, 0, 0, 400, 400, 1)
    
End Sub

Private Sub cmdSetActive_Click()

    Call SetActiveWindow(WnHnd)
    
End Sub

Private Sub cmdSetCaption_Click()

    Call SetWindowText(WnHnd, txtNewCaption.Text)
    
End Sub

Private Sub timSetText_Timer()

    timSetText.Enabled = False
    WnHnd = GetActiveWindow
    Beep
    fraWin.Enabled = True
    
End Sub
