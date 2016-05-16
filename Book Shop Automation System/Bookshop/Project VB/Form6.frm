VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Save Changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9000
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000080FF&
      Caption         =   "<-- Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Delete Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form6.frx":3AA08
      Height          =   5175
      Left            =   4440
      TabIndex        =   0
      Top             =   3600
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16384
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   24
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Aparajita"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adogrid 
      Height          =   855
      Left            =   14400
      Top             =   10080
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Your Library"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   8400
      TabIndex        =   5
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   9000
      Picture         =   "Form6.frx":3AA1E
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Book Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim a As Integer
a = MsgBox("Do you want to delete record", vbYesNo + vbExclamation, "Warning message")
If a = vbYes Then
adogrid.Recordset.Delete
MsgBox "Record deleted", vbInformation, "Delete record"
Else
MsgBox "Record not deleted", vbInformation, "Record saved"
End If
End Sub

Private Sub Command2_Click()
MsgBox "Saved Successfully", vbOKOnly, "Saved!"
End Sub

Private Sub Command5_Click()
Form4.Show
Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next
adogrid.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Book Shop Automation System\Bookshop\Database3.mdb;Persist Security Info=False"
adogrid.RecordSource = "select * From Books"
Set DataGrid1.DataSource = adogrid

End Sub

Private Sub Text1_Change()
Dim rs As ADODB.Recordset

Set rs = adogrid.Recordset
'Title
With rs
        .Close
        .Source = "select * from Books where TITLE like '%" & Text1.Text & "%'"
        .Open
End With
    DataGrid1.ReBind

End Sub
