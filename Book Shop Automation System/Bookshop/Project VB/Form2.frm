VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "I wanna Go Home!"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   16920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log Me In!"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   13440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5760
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   0
      Top             =   4440
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Login as Buyer / Make Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   13440
      TabIndex        =   5
      Top             =   9480
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid Username/Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13440
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Integer
If (Text1.Text = "Snehil" And Text2.Text = "Santhalia") Then
Form4.Show
Unload Me
ElseIf (Text1.Text = "Subham" And Text2.Text = "Sah") Then
Form4.Show
Unload Me
Else
Label1.Visible = True
a = MsgBox("Invalid Username/Password", vbAbortRetryIgnore, "Error")
If a = 3 Then
Form3.Show
Unload Me
ElseIf a = 4 Then
Label1.Visible = False
Text1.Text = ""
Text2.Text = ""
End If
End If
End Sub

Private Sub Command2_Click()
Form3.Show
Form2.Hide
End Sub

Private Sub Label2_Click()
frmLogin.Show
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then  ' The ENTER key.
Dim a As Integer
If (Text1.Text = "Snehil" And Text2.Text = "Santhalia") Then
Form4.Show
Unload Me
ElseIf (Text1.Text = "Subham" And Text2.Text = "Sah") Then
Form4.Show
Unload Me
Else
Label1.Visible = True
a = MsgBox("Invalid Username/Password", vbAbortRetryIgnore, "Error")
If a = 3 Then
Form3.Show
Unload Me
ElseIf a = 4 Then
Label1.Visible = False
Text1.Text = ""
Text2.Text = ""
End If
End If
End If
End Sub
