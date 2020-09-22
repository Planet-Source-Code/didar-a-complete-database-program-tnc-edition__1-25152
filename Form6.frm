VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6405
   ControlBox      =   0   'False
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4530
         Left            =   60
         Picture         =   "Form6.frx":08CA
         ScaleHeight     =   4500
         ScaleWidth      =   6000
         TabIndex        =   1
         Top             =   120
         Width           =   6030
         Begin VB.Timer Timer1 
            Left            =   240
            Top             =   120
         End
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 2000
End Sub

Private Sub Timer1_Timer()
Unload Me
Load Form1
Form1.Show
End Sub
