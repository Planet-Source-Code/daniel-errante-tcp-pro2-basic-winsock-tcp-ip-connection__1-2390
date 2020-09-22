VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4830
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "tcpprosplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   240
         Top             =   3240
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1560
         Picture         =   "tcpprosplash.frx":000C
         Top             =   2160
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3480
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Click to continue (or try to click the flying TCP Pro2 icon."
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "by Daniel Errante"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   360
         Picture         =   "tcpprosplash.frx":0316
         Top             =   360
         Width           =   2865
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1560
         Picture         =   "tcpprosplash.frx":1C43
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "TCP Pro2 © 1999 Daniel Errante.  All rights reserved."
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Click()
Form1.Show
Unload Me

End Sub

Private Sub Form_Load()
    Load Form1
    DoEvents
    
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "Contragulations, you clicked me, the flying TCP Pro2 icon!", vbExclamation
Form1.Show
Unload Me

End Sub

Private Sub Image3_Click()
Form1.Show
Unload Me

End Sub

Private Sub Timer1_Timer()
Image2.Left = Image2.Left + 200
Image2.Top = Image2.Top + 200
If Image2.Left >= 3120 Then Image2.Left = 0
If Image2.Top >= 4120 Then Image2.Top = 0

End Sub

Private Sub Timer2_Timer()
End Sub

Private Sub Frame1_Click()
Form1.Show
Unload Me

End Sub

Private Sub Image1_Click()
Form1.Show
Unload Me

End Sub

Private Sub Image2_Click()
Form1.Show
Unload Me

End Sub

Private Sub Label1_Click()
Form1.Show
Unload Me

End Sub

Private Sub Label3_Click()
Form1.Show
Unload Me

End Sub
