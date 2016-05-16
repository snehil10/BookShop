VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-- Move back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8280
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   9
      Top             =   6720
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   8
      Top             =   5760
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   7
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   6
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   5
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7080
      X2              =   8400
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7080
      X2              =   8280
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   13440
      X2              =   13440
      Y1              =   1320
      Y2              =   9480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7080
      X2              =   13440
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7080
      X2              =   13440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7080
      X2              =   7080
      Y1              =   1320
      Y2              =   9480
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Must know all this about a Book!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   14280
      Picture         =   "Form5.frx":300C42
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Image Image2 
      Height          =   5655
      Left            =   2160
      Picture         =   "Form5.frx":315366
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   5880
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
DataEnvironment1.AddBook Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text
a = MsgBox("Sucessfully Added", vbOKOnly, "Adding Data")
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command2_Click()
Form4.Show
Unload Me
End Sub

