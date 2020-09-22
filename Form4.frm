VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About General Database"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   4560
         Left            =   120
         Picture         =   "Form4.frx":08CA
         ScaleHeight     =   4500
         ScaleWidth      =   6000
         TabIndex        =   1
         Top             =   240
         Width           =   6060
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   3885
            Left            =   720
            Picture         =   "Form4.frx":18AAC
            ScaleHeight     =   3885
            ScaleWidth      =   4500
            TabIndex        =   2
            Top             =   240
            Visible         =   0   'False
            Width           =   4500
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Picture2.Visible = True
End Sub
