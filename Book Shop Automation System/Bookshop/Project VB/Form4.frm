VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H8000000D&
      Caption         =   "User Guide"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H8000000D&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000D&
      Caption         =   "Add Books to Library"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "Edit Library"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "Delete Book from Library"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "Make Invoice"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   3135
      Left            =   12840
      Picture         =   "Form4.frx":300C42
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   3135
      Left            =   9120
      Picture         =   "Form4.frx":307812
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   3135
      Left            =   5520
      Picture         =   "Form4.frx":313EEF
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   1800
      Picture         =   "Form4.frx":31BF18
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9480
      TabIndex        =   6
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents kell As Timer
Attribute kell.VB_VarHelpID = -1
Private Sub Command1_Click()
Dim b As Integer
b = MsgBox("Do you want to logout", vbYesNo, "Logout confirmation")
If b = vbYes Then
Form2.Show
Me.Hide
Else
MsgBox "Logout failed", vbInformation, "Continue"
End If
End Sub

Private Sub Command2_Click()
Form7.Command1.Visible = True
Form7.Show
Me.Hide
End Sub

Private Sub Command3_Click()
Form6.Command1.Visible = True
Form6.Show
Me.Hide
End Sub

Private Sub Command4_Click()
Form6.Command1.Visible = False
Form6.Show
Me.Hide
End Sub

Private Sub Command5_Click()
Form5.Show
Me.Hide
End Sub

Private Sub Command6_Click()
Form9.Show
End Sub

Private Sub Command7_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()
Label1.Caption = "Welcome, " + Form2.Text1.Text
Set kell = Form4.Controls.Add("vb.timer", "kell", Form4)
With kell: .Interval = 200: .Enabled = True: End With
End Sub

Private Sub kell_Timer()
Label2.Caption = Format$(Time, "hh:mm:ss AM/PM")
Label3.Caption = Format$(Now, "dddd, mmmm dd, yyyy")
End Sub

