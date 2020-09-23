VERSION 5.00
Begin VB.Form frmRegVideo 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtType 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtVideoID 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtVideoName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Register"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblVideoID 
      Alignment       =   1  'Right Justify
      Caption         =   "VideoID"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblVideoName 
      Alignment       =   1  'Right Justify
      Caption         =   "Video Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblPrice 
      Alignment       =   1  'Right Justify
      Caption         =   "Price"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmRegVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As Connection
Dim adoRS As ADODB.Recordset
Dim update As Boolean

Function closeConnection()

    If (cn <> Null) Then
        cn.Close
        Set cn = Nothing
    End If
    
End Function

Function connDB()
   
    Dim strConnect As String
    Dim strProvider As String
    Dim strDataSource As String
    Dim strDataBaseName As String
    
    strProvider = "Provider= Microsoft.Jet.OLEDB.3.51;"
    strDataSource = App.Path
    strDataBaseName = "\video.mdb;"
    strDataSource = "Data Source=" & strDataSource & strDataBaseName
    strConnect = strProvider & strDataSource
    
    Set cn = New ADODB.Connection
    cn.Open strConnect
    
    Set adoRS = New ADODB.Recordset
    
    adoRS.CursorType = adOpenStatic
    adoRS.CursorLocation = adUseClient
    adoRS.LockType = adLockPessimistic
    
End Function

Private Sub cmdCancel_Click()
    Call clearFields
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call closeConnection
    Call connDB
    Dim intF As Integer
    'test if all fields all corect
    'we will only test if the necessary fields are field
    If (txtVideoName.Text <> "" And txtPrice <> "") Then
        If Not (update) Then
            'update DB
            adoRS.ActiveConnection = cn
            adoRS.Source = "INSERT INTO tblVideo (VideoID, VideoName, Description, Price, Type, RentedTo) Values('" _
                        & Val(txtVideoID.Text) & "','" & txtVideoName.Text & "','" _
                        & txtDescription.Text & "','" & Val(txtPrice.Text) & "','" _
                        & txtType.Text & "','-1')"
        adoRS.Open
        Call clearFields
    
        Else
            adoRS.ActiveConnection = cn
            adoRS.Source = "UPDATE tblVideo SET VideoName ='" & txtVideoName.Text _
                        & "', Description ='" & txtDescription.Text _
                        & "', Price ='" & Val(txtPrice.Text) _
                        & "', Type ='" & txtType.Text & "', RentedTo ='-1'" _
                        & " WHERE VideoID =" & Val(txtVideoID.Text)
            adoRS.Open
            Call clearFields
    
        End If
    Else
        intF = MsgBox("Video Name and Price MUST be filed", vbCritical, "Fil all necesarry fields")
    End If

End Sub

Private Sub Form_Load()
    Call connDB
    Call clearFields
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPrice.SetFocus
    End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtType.SetFocus
    End If
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOK.SetFocus
    End If
End Sub

Private Sub txtVideoName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtDescription.SetFocus
    End If
End Sub

Private Sub txtVideoID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If (txtVideoID.Text <> "") Then
            Call getVideo
            txtVideoID.Enabled = False
            txtVideoName.SetFocus
        Else
            Call closeConnection
            Unload Me
        End If
    End If
End Sub

Private Sub txtVideoID_LostFocus()
        If (txtVideoID.Text <> "") Then
            Call getVideo
            txtVideoID.Enabled = False
            txtVideoName.SetFocus
        Else
            Call closeConnection
            Unload Me
        End If
End Sub

Function getVideo()
    
    'get the video data
    Call closeConnection
    Call connDB
    
    adoRS.ActiveConnection = cn
    adoRS.Source = "SELECT * FROM tblVideo WHERE VideoID=" _
                    & Val(txtVideoID.Text)
    adoRS.Open
    If Not (adoRS.EOF) Then
        txtVideoName.Text = adoRS.Fields("VideoName").Value
        txtDescription.Text = adoRS.Fields("Description").Value
        txtPrice.Text = adoRS.Fields("Price").Value
        txtType.Text = adoRS.Fields("Type").Value
        update = True
    Else
        update = False
    End If
    
End Function

Function clearFields()
    
    txtVideoID.Text = ""
    txtVideoID.Enabled = True
    txtVideoName.Text = ""
    txtDescription.Text = ""
    txtPrice.Text = ""
    txtType.Text = ""
    
End Function
