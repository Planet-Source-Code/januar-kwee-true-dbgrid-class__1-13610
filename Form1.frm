VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\VbJob\clsTDBGrid\demo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1860
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid2 
      Bindings        =   "Form1.frx":0000
      Height          =   3015
      Left            =   5760
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "Form1.frx":2702
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "Form1.frx":2716
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Default TdbGrid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "USE CLASS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myClass As clsTDBGrid

Private Sub Form_Load()
   Data1.DatabaseName = App.Path & "\demo.mdb"
   Data1.RecordSource = "select [Order ID],[Customer ID],[Order Date],[Freight],[Ship Country] from Orders"
   Data1.Refresh
   
   Set myClass = New clsTDBGrid
   myClass.SetGridName = TDBGrid1
   myClass.setFormName = Me
   myClass.setDataControl = Data1
   
   Data2.DatabaseName = App.Path & "\demo.mdb"
   Data2.RecordSource = "select [Order ID],[Customer ID],[Order Date],[Freight],[Ship Country] from Orders"
   Data2.Refresh
   
End Sub
