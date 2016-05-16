VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15885
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cart"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12720
      TabIndex        =   9
      Top             =   9360
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   7
      Top             =   2880
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc adogrid 
      Height          =   975
      Left            =   17280
      Top             =   9840
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1720
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":300C42
      Height          =   4575
      Left            =   4920
      TabIndex        =   6
      Top             =   3720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   8421440
      ForeColor       =   -2147483637
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add to Cart"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      TabIndex        =   5
      Top             =   9360
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   4
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<-- Move Back"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7440
      TabIndex        =   0
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   9000
      Picture         =   "Form7.frx":300C58
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   5295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Book Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10560
      TabIndex        =   2
      Top             =   8520
      Width           =   1455
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As String



Private Sub Command1_Click()
Form4.Show
Unload Me
End Sub

''Private Sub Command2_Click()
'adogrid.Recordset.MoveFirst

'While Not adogrid.Recordset.EOF
'Dim gcol, gcol2, gcol1, gcol3 As MSDataGridLib.Column
'Set gcol1 = Form8.DataGrid1.Columns("Quantity")
'Set gcol = Form8.DataGrid1.Columns("ID")
'Set gcol2 = Form8.DataGrid1.Columns("StockQuantity")
'DataEnvironment1.Edit Val(gcol2) - Val(gcol1), gcol
'adogrid.Recordset.MoveNext
'Wend

'DataReport1.Show

'If DataEnvironment2.rsCommand1.State = adStateOpen Then
'DataEnvironment2.rsCommand1.Close
'End If
'
'End Sub

Private Sub Command3_Click()

Dim gcol, gcol2, gcol1, gcol3 As MSDataGridLib.Column
Set gcol = DataGrid1.Columns("Quantity")
Set gcol2 = DataGrid1.Columns("Price")
Set gcol1 = DataGrid1.Columns("Title")
Set gcol3 = DataGrid1.Columns("ID")

If (gcol < Val(Text1.Text)) Then
MsgBox "Sorry! Your Expectation is way high than our Stock right now", vbOKOnly, "Error"
Else
DataEnvironment3.Add gcol1, Val(Text2.Text), Val(Text1.Text), gcol2, gcol3, gcol
total = Val(total) + 1
Command4.Caption = "Cart" + "(" + total + ")"
End If


End Sub

Private Sub Command4_Click()
Form8.Show
End Sub

Private Sub Form_Load()

On Error Resume Next
adogrid.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Book Shop Automation System\Bookshop\Database3.mdb;Persist Security Info=False"
adogrid.RecordSource = "select * From Books"
Set DataGrid1.DataSource = adogrid

DataEnvironment3.Delete
total = "0"
End Sub

Private Sub Text1_LostFocus()

Dim gcol As MSDataGridLib.Column
Set gcol = DataGrid1.Columns("Price")
If (gcol > Val(Text1.Text)) Then
Text2.Text = Val(Text1.Text) * gcol
End If
End Sub

Private Sub Text3_Change()
Dim rs As ADODB.Recordset

Set rs = adogrid.Recordset
'Title
With rs
        .Close
        .Source = "select * from Books where TITLE like '%" & Text3.Text & "%'"
        .Open
End With
    DataGrid1.ReBind

End Sub
