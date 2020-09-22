VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   5145
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5000
      Left            =   120
      TabIndex        =   9
      Top             =   30
      Width           =   5055
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   240
         TabIndex        =   21
         Top             =   3960
         Width           =   4575
         Begin VB.CommandButton Command2 
            Caption         =   "Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   4575
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "CustomerID"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "price"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   1
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "CompanyName"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   0
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "product"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   2
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1800
            TabIndex        =   12
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Phone"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1800
            TabIndex        =   3
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1800
            TabIndex        =   4
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product"
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
            Left            =   600
            TabIndex        =   20
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID "
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
            Left            =   840
            TabIndex        =   19
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Price"
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
            Left            =   840
            TabIndex        =   18
            Top             =   1200
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customers Name"
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
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   840
            TabIndex        =   16
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Left            =   600
            TabIndex        =   15
            Top             =   2640
            Width           =   690
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            Left            =   720
            TabIndex        =   14
            Top             =   2280
            Width           =   555
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   1440
         X2              =   3360
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Top             =   120
         Width           =   1530
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Text1.Text = Text1.Text
Form1.Text7.Text = aa
Form1.Text2.Text = Text2.Text
Form1.Text6.Text = Text6.Text
Form1.Text4.Text = Text4.Text
Form1.Text8.Text = Text8.Text
Form1.Text9.Text = Text9.Text



Form1.Data1.Recordset.AddNew
Form1.Data1.Recordset("customer_name") = Left(Form1.Text1.Text, (Len(Form1.Text1.Text)) - 2)
Form1.Data1.Recordset("customerid") = aa
Form1.Data1.Recordset("price") = Left(Form1.Text6.Text, (Len(Form1.Text6.Text)) - 2)
Form1.Data1.Recordset("product") = Left(Form1.Text2.Text, (Len(Form1.Text2.Text)) - 2)
Form1.Data1.Recordset("date") = Left(Form1.Text4.Text, (Len(Form1.Text4.Text)) - 2)
Form1.Data1.Recordset("Address") = Left(Form1.Text8.Text, (Len(Form1.Text8.Text)) - 2)
Form1.Data1.Recordset("Phone") = Left(Form1.Text9.Text, (Len(Form1.Text9.Text)) - 2)



Form1.Data1.Recordset.Update
Form1.dg.Refresh

Command2.Visible = True
Command1.Visible = False
Command2_Click

End Sub

Private Sub Command2_Click()
On Error Resume Next
aa = aa + 1
Form1.Data1.Recordset.AddNew
Form1.Text1.Text = ""
Form1.Text7.Text = aa
Form1.Text2.Text = ""
Form1.Text6.Text = ""
Form1.Text4.Text = ""
Form1.Text4.Text = Date & "<>" & Time
Form1.Text8.Text = ""
Form1.Text9.Text = ""


Text1.Text = ""
Text7.Text = aa
Text2.Text = ""
Text6.Text = ""
Text4.Text = Date & "<>" & Time
Text8.Text = ""
Text9.Text = ""

Command1.Visible = True
Command2.Visible = False
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Form_Load()
Text4.Text = Date & "<>" & Time

End Sub
