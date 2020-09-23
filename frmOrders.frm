VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmOrders 
   Caption         =   "Orders"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Orders"
      Connect         =   "Access"
      DatabaseName    =   "E:\Moji programi\video.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1980
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmOrders.frx":0000
      Height          =   2295
      Left            =   240
      OleObjectBlob   =   "frmOrders.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label lblRecords 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   6255
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strText
    strText = "Number of registered orders in database is: "
    Data1.RecordSource = "SELECT * FROM tblOrder"
    Data1.Refresh
    If Data1.Recordset.RecordCount <> 0 Then
      Data1.Recordset.MoveLast
      Data1.Recordset.MoveFirst
      lblRecords.Caption = strText & Data1.Recordset.RecordCount
    Else
      lblRecords.Caption = strText & "0"
    End If
    
End Sub
