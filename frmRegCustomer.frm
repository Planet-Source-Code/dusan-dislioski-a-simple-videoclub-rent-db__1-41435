VERSION 5.00
Begin VB.Form frmRegCustomer 
   Caption         =   "Register Customer"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTelefon 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Register Customer"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtID 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblTlf 
      Alignment       =   1  'Right Justify
      Caption         =   "Telefon"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblZip 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblCity 
      Alignment       =   1  'Right Justify
      Caption         =   "City"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblLastName 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      Caption         =   "CustomerID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmRegCustomer"
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

Private Sub cmdCancel_Click()
    
    ' clears all fields and returns to main menu
    Call clearFields
    Call clearForm
    Call showCID
    Unload Me
    
End Sub

Private Sub cmdUpdate_Click()
    
    Call closeConnection
    Call connDB
    
    
    ' test if all fields all corect
    'update DB
    If Not (update) Then
        adoRS.ActiveConnection = cn
        adoRS.Source = "INSERT INTO tblCustomer (CustomerID, Name, LastName, Address, Telefon, City, Zip) Values('" _
                    & Val(txtID.Text) & "','" & txtName.Text & "','" _
                    & txtLastName.Text & "','" & txtAddress.Text & "','" _
                    & txtTelefon.Text & "','" _
                    & txtCity.Text & "','" & txtZip.Text & "')"
    adoRS.Open
    Call clearFields
    Unload Me
    Else
        adoRS.ActiveConnection = cn
        adoRS.Source = "UPDATE tblCustomer SET Name ='" & txtName.Text _
                    & "', LastName ='" & txtLastName.Text & "', Address ='" & txtAddress.Text _
                    & "', Telefon ='" & txtTelefon.Text & "', City ='" & txtCity.Text _
                    & "', Zip ='" & txtZip.Text & "' WHERE CustomerID =" & Val(txtID.Text)
                        
                    

        adoRS.Open
    Call clearFields
    Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call connDB
End Sub

Function clearFields()

    txtName.Text = ""
    txtLastName.Text = ""
    txtAddress.Text = ""
    txtTelefon.Text = ""
    txtCity.Text = ""
    txtZip.Text = ""

End Function

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtTelefon.SetFocus
    End If
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtZip.SetFocus
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If (txtID <> "") Then
            Call txtID_LostFocus
            txtName.SetFocus
        Else
            Call closeConnection
            Call clearFields
            Unload Me
        End If
    End If
End Sub

Private Sub txtID_LostFocus()
    Call closeConnection
    Call connDB
    
    adoRS.ActiveConnection = cn
    adoRS.Source = "Select * From tblCustomer WHERE CustomerID=" _
                    & Val(txtID.Text)
    adoRS.Open
    If Not (adoRS.EOF And adoRS.BOF) Then
        txtName.Text = adoRS.Fields("Name").Value
        txtLastName.Text = adoRS.Fields("LastName").Value
        txtAddress.Text = adoRS.Fields("Address").Value
        txtTelefon.Text = adoRS.Fields("Telefon").Value
        txtCity.Text = adoRS.Fields("City").Value
        txtZip.Text = adoRS.Fields("Zip").Value
        update = True
    Else
        update = False
    End If
End Sub


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

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtAddress.SetFocus
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLastName.SetFocus
    End If
End Sub

Private Sub txtTelefon_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCity.SetFocus
    End If
End Sub

Private Sub txtZip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdUpdate.SetFocus
    End If
End Sub
