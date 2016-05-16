VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1320
      Top             =   2280
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   2280
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   2
      Top             =   6840
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   8640
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Timer1_Timer()

If ProgressBar1.Value = ProgressBar1.Max Then
Form2.Show
Unload Me
Timer1.Enabled = False
End If

i = i + 1
ProgressBar1.Value = ProgressBar1.Value + 1

If (i Mod 2 = 0) Then
Label2.Caption = Label2.Caption + " . "
If (i Mod 34 = 0) Then
Label2.Caption = ""
End If
End If

Select Case i
Case 2
Label1.Caption = "Loading Forms.."
Case 18
Label1.Caption = "Connecting Database.."
Case 28
Label1.Caption = "Preparing User Interface..."
Case 48
Label1.Caption = "Checking Connectivity."
Case 52
Label1.Caption = "Checking Connectivity. ."
Case 56
Label1.Caption = "Checking Connectivity. . ."
Case 60
Label1.Caption = "Checking Connectivity."
Case 64
Label1.Caption = "Checking Connectivity. ."
Case 68
Label1.Caption = "Checking Connectivity. . ."
Case 72
Label1.Caption = "Checking Connectivity. . . . ."
Case 78
Label1.Caption = "Preparing Accounts Info.."
Case 94
Label1.Caption = "Welcome!"

End Select


End Sub
