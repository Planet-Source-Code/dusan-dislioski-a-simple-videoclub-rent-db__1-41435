VERSION 5.00
Begin VB.Form frmRegProduct 
   Caption         =   "Register Product"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Register"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtProductName 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtProductID 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1215
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
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblProductName 
      Alignment       =   1  'Right Justify
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblProductID 
      Alignment       =   1  'Right Justify
      Caption         =   "ProductID"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmRegProduct"
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
    ' clears all fields and returns to main menu
    Call clearForm
    Call showCID
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call closeConnection
    Call connDB
    
    'test if all fields all corect
    'we will only test if the necessary fields are field
    If (txtProductName.Text <> "" And txtPrice <> "") Then
        If Not (update) Then
            'update DB
            adoRS.ActiveConnection = cn
            adoRS.Source = "INSERT INTO tblProduct (ProductID, ProductName, Description, Price) Values('" _
                        & Val(txtProductID.Text) & "','" & txtProductName.Text & "','" _
                        & txtDescription.Text & "','" & Val(txtPrice.Text) & "')"
        adoRS.Open
        Call clearFields
    
        Else
            adoRS.ActiveConnection = cn
            adoRS.Source = "UPDATE tblProduct SET ProductName ='" & txtProductName.Text _
                        & "', Description ='" & txtDescription.Text _
                        & "', Price ='" & Val(txtPrice.Text) _
                        & "' WHERE ProductID =" & Val(txtProductID.Text)
            adoRS.Open
            Call clearFields
    
        End If
    Else
        intF = MsgBox("Product Name and Price MUST be filed", vbCritical, "Fil all necesarry fields")
    End If
End Sub

Private Sub Form_Load()
    Call connDB
End Sub

Function clearFields()
    
    txtProductID.Text = ""
    txtProductID.Enabled = True
    txtProductName.Text = ""
    txtDescription.Text = ""
    txtPrice.Text = ""
    txtProductID.SetFocus
    
End Function

Function getProduct()
    'get the product data
    Call closeConnection
    Call connDB
    
    adoRS.ActiveConnection = cn
    adoRS.Source = "SELECT * FROM tblProduct WHERE ProductID=" _
                    & Val(txtProductID.Text)
    adoRS.Open
    If Not (adoRS.EOF) Then
        txtProductName.Text = adoRS.Fields("ProductName").Value
        txtDescription.Text = adoRS.Fields("Description").Value
        txtPrice.Text = adoRS.Fields("Price").Value
        update = True
    Else
        update = False
    End If
    
End Function

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPrice.SetFocus
    End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOK.SetFocus
    End If
End Sub

Private Sub txtProductID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If (txtProductID.Text <> "") Then
            Call getProduct
            txtProductID.Enabled = False
            txtProductName.SetFocus
        Else
            Call closeConnection
            Unload Me
        End If
    End If
End Sub

Private Sub txtProductID_LostFocus()
        If (txtProductID.Text <> "") Then
            Call getProduct
            txtProductID.Enabled = False
            txtProductName.SetFocus
        Else
            Call closeConnection
            Unload Me
        End If
End Sub

Private Sub txtProductName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtDescription.SetFocus
    End If
End Sub
