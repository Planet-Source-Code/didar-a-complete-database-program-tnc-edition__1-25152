VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Database"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   6540
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   9465
      TabIndex        =   16
      Top             =   0
      Width           =   9495
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6240
         TabIndex        =   28
         ToolTipText     =   "Find Name Box"
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   8400
         TabIndex        =   27
         ToolTipText     =   "Find ID Box"
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3000
         Picture         =   "Form1.frx":157F7
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Print.."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command13 
         Height          =   495
         Left            =   1200
         Picture         =   "Form1.frx":15E61
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Load Default Database"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command12 
         Height          =   495
         Left            =   5400
         Picture         =   "Form1.frx":162A3
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Quit.."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3600
         Picture         =   "Form1.frx":166E5
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Delete Selected Row"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Height          =   495
         Left            =   4800
         Picture         =   "Form1.frx":16B27
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Help.."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         Picture         =   "Form1.frx":16F69
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Update Data"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Height          =   495
         Left            =   4200
         Picture         =   "Form1.frx":173AB
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Info.."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2400
         Picture         =   "Form1.frx":17AED
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Search.."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   600
         Picture         =   "Form1.frx":17F2F
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Open Database File"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         Picture         =   "Form1.frx":18371
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Add New Member"
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      DataField       =   "CustomerID"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text6 
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
      Left            =   6600
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      DataField       =   "Customer_Name"
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
      Left            =   6600
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      Left            =   6600
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Height          =   2415
      Left            =   -720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Text            =   "Form1.frx":187B3
      Top             =   2280
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.TextBox Text9 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      DataField       =   "Phone"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      DataField       =   "Date"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   855
      Left            =   5280
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid dg 
      Bindings        =   "Form1.frx":188B1
      Height          =   6055
      Left            =   0
      OleObjectBlob   =   "Form1.frx":188C5
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   9495
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Students Data"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\DESKTOP\General Database\didar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   6840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Corporation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7800
      TabIndex        =   14
      Top             =   6600
      Width           =   1710
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "CopyRight by Didar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "General Corporation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Printing 
         Caption         =   "Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^{F6}
      End
   End
   Begin VB.Menu Load 
      Caption         =   "Load"
      Begin VB.Menu LoadDefault 
         Caption         =   "Load Default Database"
         Shortcut        =   ^L
      End
      Begin VB.Menu Make 
         Caption         =   "Make A New Database File"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Add 
         Caption         =   "Add"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu Update 
         Caption         =   "Update"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu Printsel 
         Caption         =   "Print Selected Data"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu Find 
      Caption         =   "Find"
      Begin VB.Menu FindName 
         Caption         =   "Find Name"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu FindNumber 
         Caption         =   "FindNumber"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu FND 
         Caption         =   "Find Name List"
         Enabled         =   0   'False
      End
      Begin VB.Menu MoveFirst 
         Caption         =   "Move First"
         Enabled         =   0   'False
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu MoveLast 
         Caption         =   "Move Last"
         Enabled         =   0   'False
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu Calculate 
      Caption         =   "Calculate"
      Begin VB.Menu calmem 
         Caption         =   "Calculate Total Member"
         Enabled         =   0   'False
      End
      Begin VB.Menu calprice 
         Caption         =   "Calculate Total Price"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Date 
      Caption         =   "Sp Search"
      Begin VB.Menu Datese 
         Caption         =   "Date Search Calculation"
         Enabled         =   0   'False
      End
      Begin VB.Menu Pro 
         Caption         =   "Product Search"
         Enabled         =   0   'False
      End
      Begin VB.Menu Address 
         Caption         =   "Address Search"
         Enabled         =   0   'False
      End
      Begin VB.Menu Sstart 
         Caption         =   "Selection Search"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Help1 
      Caption         =   "Help"
      Begin VB.Menu Help 
         Caption         =   "Help"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu AboutIUBdata 
         Caption         =   "About General Database"
      End
      Begin VB.Menu SystemInfo 
         Caption         =   "SytemInfo"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AboutIUBdata_Click()
Form4.Show
End Sub

Private Sub Close_Click()
Dim a As Integer
a = MsgBox("Do You Really Want to Close.", vbYesNo, "Close..")
If a = vbNo Then
Form1.Show
Else
End
End If
End Sub


Private Sub Add_Click()
On Error Resume Next
Data1.Recordset.MoveLast
aa = Data1.Recordset("customerid")
Text7.Text = aa
Form3.Text7.Text = aa
Data1.Recordset.AddNew
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Form3.Show
Exit Sub
error:
MsgBox "Invalid Data Entry.Try again.", 16, "Data Entry Error!"
End Sub



Private Sub Address_Click()
Dim x As Integer
Dim aa As Integer
Dim gs As String
Const AposAst As String = "'*", AstApos As String = "*'"
Dim target As String
On Error Resume Next
gs = ""
x = 0
gs = "Customer ID===Customer Name===Date===Product===Price===Phone===Address" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10))
ds$ = InputBox("Please Enter The Address...Like 'Khulshi'", "Address Search", tt)
If ds = "" Then
Exit Sub
End If

Data1.Recordset.MoveFirst
target = "Address like" & AposAst & ds & AstApos
For i = 0 To (dg.ApproxCount - 1)
Data1.Recordset.FindNext target
If Data1.Recordset.NoMatch Then
Exit For
End If
gs = gs & (Chr(13) & Chr(10)) & Data1.Recordset("Customerid") & "===" & Data1.Recordset("Customer_name") & "===" & Data1.Recordset("Date") & "===" & Data1.Recordset("Product") & "===" & Data1.Recordset("price") & "===" & Data1.Recordset("Phone") & "===" & Data1.Recordset("Address") & (Chr(13) & Chr(10))
x = x + 1
aa = aa + Data1.Recordset("price")
Next i
MsgBox x & "   Data Found.   Total Price Counted.." & aa & Chr(13) & Chr(10) & Chr(13) & Chr(10) & gs, 32, "Total Address Search List"
End Sub

Private Sub calmem_Click()
On Error Resume Next
MsgBox "Total Member Is   " & (dg.ApproxCount - 1), 32, "Total Member"
End Sub

Private Sub calprice_Click()
Dim aa As Integer
Const AposAst As String = "'*", AstApos As String = "*'"
Dim target As String
On Error Resume Next

'Const apos As String = "'"
'Dim target As String
Text3.Text = ""
aa = 0
target = "customer_name like" & AposAst & Text3 & AstApos
Data1.Recordset.MoveFirst
For i = 0 To (dg.ApproxCount - 2)
Data1.Recordset.FindNext target
aa = aa + Data1.Recordset("price")
Next i
MsgBox "Total Price Is==" & aa, 32, "Total Price"
End Sub




Private Sub Command10_Click()
Help_Click
End Sub

Private Sub Command11_Click()
Delete_Click
End Sub

Private Sub Command12_Click()
Exit_Click
End Sub

Private Sub Command13_Click()
LoadDefault_Click
End Sub

Private Sub Command14_Click()
Printsel_Click
End Sub






Private Sub Command2_Click()
Add_Click
End Sub

Private Sub Command3_Click()
Open_Click
End Sub

Private Sub Command4_Click()
If Text5.Text = "" And Text3.Text <> "" Then
FindName_Click
Else
If Text5.Text <> "" And Text3.Text = "" Then
FindNumber_Click
Else
MsgBox "You Must Enter Name or ID", 16, "TNC Error"
Text3.SetFocus
End If
End If
End Sub

Private Sub Command7_Click()
Form4.Show
End Sub

Private Sub Command8_Click()
Update_Click
End Sub

Private Sub Command9_Click()
Form3.Show
End Sub

Private Sub Copy_Click()
Clipboard.SetText (dg.SelText)
End Sub


Private Sub Credits_Click()
Form5.Show
End Sub



Private Sub Datese_Click()
Dim aa As Integer
Dim gs As String
Dim dd As Integer

Const AposAst As String = "'*", AstApos As String = "*'"
Dim target As String
On Error Resume Next
aa = 0
gs = ""
gs = "Customer ID===Customer Name===Date===Product===Price===Phone===Address" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10))
ds$ = InputBox("Please Enter The Date To Find Data..Like '1/13/02'", "Date Search", tt)

If ds = "" Then
Exit Sub
End If


Data1.Recordset.MoveFirst
target = "Date like" & AposAst & ds & AstApos
For i = 0 To (dg.ApproxCount - 1)
Data1.Recordset.FindNext target
If Data1.Recordset.NoMatch Then
Exit For
End If
gs = gs & (Chr(13) & Chr(10)) & Data1.Recordset("Customerid") & "===" & Data1.Recordset("Customer_name") & "===" & Data1.Recordset("Date") & "===" & Data1.Recordset("Product") & "===" & Data1.Recordset("price") & "===" & Data1.Recordset("Phone") & "===" & Data1.Recordset("Address") & (Chr(13) & Chr(10))
aa = aa + 1
dd = dd + Data1.Recordset("price")
Next i
MsgBox aa & "   Data Found.   Total Price Counted.." & dd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & gs, 32, "Date Result"
End Sub

Private Sub Delete_Click()
On Error Resume Next
If Data1.Recordset("customerid") = 10000 Then
MsgBox "Cannot Delete First Item", 16, "Error"
Else
Data1.Recordset.Delete
Data1.Recordset.MoveFirst
End If
End Sub



Private Sub dg_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
 PopupMenu Form1!Edit, _
         vbPopupMenuLeftAlign, _
         x, Y, _
         Form1!Delete
   End If
End Sub

Private Sub Exit_Click()
Dim d As Integer
d = MsgBox("Do You Really Want to Quit?", 49, "Quit?")
If d = vbCancel Then
Form1.Show
Else
End
End If
End Sub


Private Sub FindIdNumber_Click()
Form2.Show
End Sub

Private Sub FindName_Click()

Const AposAst As String = "'*", AstApos As String = "*'"
Dim target As String

If Text3.Text = "" Then
MsgBox "You Must Enter a  Name!", 32, "General Warning!."
Text3.SetFocus
Else

target = "customer_name like" & AposAst & Text3 & AstApos
Form1.Data1.Recordset.FindNext target
If Data1.Recordset.NoMatch Then
'MsgBox "Not Found Any Name.", 32, "Invalid Name."
Data1.Recordset.FindFirst target
Text3.SetFocus
'Else
'Do Until Data1.Recordset.NoMatch
'Data1.Recordset.FindNext target
'Loop
End If
End If

End Sub

Private Sub FindNumber_Click()
Const apos As String = "'"
Dim target As String
target = "customerid=" & apos & Text5 & apos
Data1.Recordset.FindFirst target
If Form1.Data1.Recordset.NoMatch Then
MsgBox "No match with that Id number ", vbExclamation, "Membership Varification"
Text5 = ""
End If
End Sub

Private Sub FND_Click()
Dim x As Integer
Dim gs As String
Dim dd As Integer
Const AposAst As String = "'*", AstApos As String = "*'"
Dim target As String
On Error Resume Next
gs = ""
x = 0

If Text3.Text = "" Then
MsgBox "You Must Enter A Name", 16, "TNC Error"
Text3.SetFocus
Else
gs = "Customer ID===Customer Name===Date===Product===Price===Phone===Address" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10))
Data1.Recordset.MoveFirst
target = "customer_name like" & AposAst & Text3 & AstApos
For i = 0 To (dg.ApproxCount - 1)
Data1.Recordset.FindNext target
If Data1.Recordset.NoMatch Then
Exit For
End If
gs = gs & (Chr(13) & Chr(10)) & Data1.Recordset("Customerid") & "===" & Data1.Recordset("Customer_name") & "===" & Data1.Recordset("Date") & "===" & Data1.Recordset("Product") & "===" & Data1.Recordset("price") & "===" & Data1.Recordset("Phone") & "===" & Data1.Recordset("Address") & (Chr(13) & Chr(10))
x = x + 1
dd = dd + Data1.Recordset("price")
Next i
MsgBox x & "   Data Found.   Total Price Counted..." & dd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & gs, 32, "Total Find Name List"
End If
End Sub

Private Sub Form_Load()
On Error GoTo err
Data1.DatabaseName = (App.Path & "\tnc")
Exit Sub
err:
MsgBox "Default Database File 'TNC' Is Readonly Or Corrupted, Check The Files Is Exist Or Not", 16, "TNC Error"
End Sub



Private Sub Help_Click()
On Error Resume Next

msg = "                                                   HELP 3.0 " & Chr(13) & Chr(10)
msg = msg & Chr(13) & Chr(10)
msg = msg & "It's A Customized Database Software Which Is Based On OLE (Object Linking And Embadding), DDE(Dynamic Data Exchanged) And GIS (Geographical Informaion System). It Has An Strong Microsoft Database Engine." & Chr(13) & Chr(10)
msg = msg & Chr(13) & Chr(10)
msg = msg & "Customer Id Will Be Added Automatically, Which Was Not Possible Before.You Can Simply Use This Software By Using Shortcut Key.For Better Help See The Read Me File."
msg = msg & "Here The Default Database Name Is 'TNC'.So Take A Responsibility To Protect This Database." & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "To Get More Help Please Contact With     http//:www.General.com"

MsgBox msg, 64, "Help"
End Sub




Private Sub LoadDefault_Click()
On Error Resume Next


Text3.Visible = True
Text5.Visible = True
dg.Visible = True
MoveFirst.Enabled = True
MoveLast.Enabled = True
FindName.Enabled = True
FindNumber.Enabled = True
Add.Enabled = True
Delete.Enabled = True
Copy.Enabled = True
Paste.Enabled = True
Sstart.Enabled = True
FND.Enabled = True
Datese.Enabled = True
Address.Enabled = True
Pro.Enabled = True
Update.Enabled = True
calprice.Enabled = True
calmem.Enabled = True
Printing.Enabled = True
Printsel.Enabled = True
Command2.Enabled = True
Command8.Enabled = True
Command4.Enabled = True
Command14.Enabled = True
Command11.Enabled = True

End Sub

Private Sub Make_Click()
On Error GoTo err
Data1.DatabaseName = ""
Data1.Refresh
dg.Refresh
ans$ = InputBox("Please Enter A New Database FileName", "New Database", tt)
FileCopy (App.Path & "\data"), (App.Path & "\" & ans)
Exit Sub
err:
MsgBox "Cannot Make New File.Make Sure The 'Data' File Is Readonly Or Corrupted.", 16, "TNC Error"
End Sub

Private Sub MoveFirst_Click()
On Error Resume Next
Data1.Recordset.MoveFirst
End Sub

Private Sub MoveLast_Click()
On Error Resume Next
Data1.Recordset.MoveLast
End Sub

Private Sub Open_Click()
On Error GoTo err

cmdlg.Filter = "*.*|*.*"
cmdlg.ShowOpen
If cmdlg.FileName = "" Then
MsgBox "No Database file Selected.", 48, "File not Found."
Else
Data1.DatabaseName = cmdlg.FileName
Data1.Refresh
Text3.Visible = True
Text5.Visible = True
dg.Visible = True
MoveFirst.Enabled = True
MoveLast.Enabled = True
FindName.Enabled = True
FindNumber.Enabled = True
Add.Enabled = True
Delete.Enabled = True
Copy.Enabled = True
Paste.Enabled = True
FND.Enabled = True
Address.Enabled = True
Sstart.Enabled = True
Pro.Enabled = True
Datese.Enabled = True
calprice.Enabled = True
calmem.Enabled = True
Update.Enabled = True
Printing.Enabled = True
Printsel.Enabled = True
Command2.Enabled = True
Command8.Enabled = True
Command4.Enabled = True
Command14.Enabled = True
Command11.Enabled = True
dg.Refresh
End If
Exit Sub
err:
MsgBox "Readonly Or Unknown Format Database Name.", 16, "Unknown"
End Sub

Private Sub Paste_Click()
On Error Resume Next
dg.SelText = Clipboard.GetText
End Sub

Private Sub Print_Click()
On Error Resume Next
Dim counter As Integer
Printer.Print dg.Row
Printer.EndDoc
End Sub

Private Sub Save_Click()
Data1.UpdateRecord
End Sub

Private Sub Printing_Click()
On Error Resume Next

msg$ = Chr(13) & Chr(10) & Text10.Text & "             Name: " & Text1.Text
msg = msg & "                                 " & "ID: " & Text7.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "             Price: " & Text6.Text & "                                                " & " Product: " & Text2.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "             Date: " & Text4.Text & "                     Phone: " & Text8.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "             Address: " & Text9.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "             ------------------------------------------------------------------------------" & Chr(10) & Chr(13)
 msg = msg & "                                  General Corporation Bangladesh"



MsgBox msg
Printer.FontSize = 13
Printer.FontBold = True
Printer.FontItalic = False
Printer.FontName = "Times New Roman"
Printer.Print msg
Printer.Print
Printer.EndDoc

End Sub

Private Sub Printsel_Click()
On Error Resume Next

msg$ = Chr(13) & Chr(10) & Text10.Text & "             Name: " & Text1.Text
msg = msg & "                                " & "ID: " & Text7.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "             Price: " & Text6.Text & "                                                " & " Product: " & Text2.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "             Date: " & Text4.Text & "                     Phone: " & Text8.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "             Address: " & Text9.Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
msg = msg & "             -------------------------------------------------------------------------------" & Chr(10) & Chr(13)
 msg = msg & "                                  General Corporation Bangladesh"



MsgBox msg
Printer.FontSize = 13
Printer.FontBold = True
Printer.FontItalic = False
Printer.FontName = "Times New Roman"
Printer.Print msg
Printer.Print
Printer.EndDoc


End Sub

Private Sub Pro_Click()
Dim x As Integer
Dim dd As Integer
Dim gs As String
Const AposAst As String = "'*", AstApos As String = "*'"
Dim target As String
On Error Resume Next
gs = ""
x = 0

ds$ = InputBox("Please Enter The Product Name..Like 'Video CD'", "Product Search", tt)
If ds = "" Then
Exit Sub
End If

gs = "Customer ID===Customer Name===Date===Product===Price===Phone===Address" & (Chr(13) & Chr(10)) & (Chr(13) & Chr(10))
Data1.Recordset.MoveFirst
target = "Product like" & AposAst & ds & AstApos
For i = 0 To (dg.ApproxCount - 1)
Data1.Recordset.FindNext target
If Data1.Recordset.NoMatch Then
Exit For
End If
gs = gs & (Chr(13) & Chr(10)) & Data1.Recordset("Customerid") & "===" & Data1.Recordset("Customer_name") & "===" & Data1.Recordset("Date") & "===" & Data1.Recordset("Product") & "===" & Data1.Recordset("price") & "===" & Data1.Recordset("Phone") & "===" & Data1.Recordset("Address") & (Chr(13) & Chr(10))
x = x + 1
dd = dd + Data1.Recordset("price")
Next i
MsgBox x & "   Data Found.   Total Price Couned..." & dd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & gs, 32, "Total Product Name List"
End Sub

Private Sub Sstart_Click()
Dim x As Integer
Dim gs As String
Dim dd As Integer
Dim ds As Integer
Dim dh As Integer
Const AposAst As String = "'*", AstApos As String = "*'"
Dim target As String
Dim tar As String

On Error Resume Next
Data1.Recordset.MoveFirst
gs = ""
x = 0
ds = InputBox("Please Enter  Where To Start Search Data..Like '10010'", "Selection Search..", tt)
dh = InputBox("Please Enter Where Searching Will Be End...Like '10015'", "Selection Search..", tt)
If dh = 0 Then
Exit Sub
End If


tar = "customerid like" & AposAst & (ds - 1) & AstApos
Data1.Recordset.FindFirst tar


target = "customerid like" & AposAst & "" & AstApos
For i = 0 To (dg.ApproxCount - 1)
Data1.Recordset.FindNext target
If Data1.Recordset("customerid") = dh + 1 Then
Exit For
End If
gs = gs & (Chr(13) & Chr(10)) & Data1.Recordset("Customerid") & "===" & Data1.Recordset("Customer_name") & "===" & Data1.Recordset("Date") & "===" & Data1.Recordset("Product") & "===" & Data1.Recordset("price") & "===" & Data1.Recordset("Phone") & "===" & Data1.Recordset("Address") & (Chr(13) & Chr(10))
x = x + 1
dd = dd + Data1.Recordset("price")
Next i
MsgBox x & "   Data Found. Total Price Counted..." & dd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & gs, 32, "Total Selection Search List"
End Sub

Private Sub SystemInfo_Click()
On Error Resume Next

frmAbout.Show
End Sub

Private Sub Text3_GotFocus()
Text5.Text = ""
End Sub

Private Sub Text5_GotFocus()
Text3.Text = ""
End Sub

Private Sub Update_Click()
On Error Resume Next
Form1.Data1.Recordset.Update
Form1.Data1.Refresh
Form1.dg.Refresh
Data1.Refresh
End Sub
