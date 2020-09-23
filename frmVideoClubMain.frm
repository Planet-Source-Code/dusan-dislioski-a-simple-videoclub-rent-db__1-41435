VERSION 5.00
Begin VB.Form frmVideoClub 
   Caption         =   "VIDEO CLUB"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowOrders 
      Caption         =   "Show Orders"
      Height          =   495
      Left            =   360
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Timer tmrAlfa 
      Interval        =   10
      Left            =   10320
      Top             =   120
   End
   Begin VB.TextBox txtReturnMID 
      Height          =   285
      Left            =   3720
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Retun Video"
      Height          =   495
      Left            =   360
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdShowRented 
      Caption         =   "Has Movies"
      Height          =   255
      Left            =   8160
      TabIndex        =   38
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox lstProductPrice 
      Height          =   450
      Left            =   3720
      TabIndex        =   37
      Top             =   6480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstMoviePrice 
      Height          =   645
      Left            =   8520
      TabIndex        =   36
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstMovieID 
      Height          =   645
      Left            =   6960
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdNextVideo 
      Caption         =   "Next Video"
      Height          =   255
      Left            =   7080
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdNextCID 
      Caption         =   "Next"
      Height          =   255
      Left            =   7080
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtSVideo 
      Height          =   285
      Left            =   3720
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdBuyP 
      Caption         =   "&Buy Product"
      Height          =   495
      Left            =   360
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdCFind 
      Caption         =   "Find &Customer"
      Height          =   495
      Left            =   360
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartUp 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.ListBox lstProducts 
      Height          =   645
      Left            =   3720
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstMovies 
      Height          =   645
      Left            =   3720
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdFinishOrder 
      Caption         =   "&Finish"
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   9480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame frmAdm 
      Height          =   3375
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   2055
      Begin VB.CommandButton cmdRegProdukt 
         Caption         =   "Register (Edit) New Product"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdRegVideo 
         Caption         =   "Register (Edit) New Video"
         Height          =   495
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdRegCustomer 
         Caption         =   "Register (Edit) New Customer"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblSubMenu 
         Alignment       =   2  'Center
         Caption         =   "ADMINISTRATION TOOLS"
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frmRent 
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   2055
      Begin VB.CommandButton cmdRent 
         Caption         =   "&Rent Video or DVD"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdFindVideo 
         Caption         =   "Find &Video"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H80000004&
      DataField       =   "Address"
      DataMember      =   "findCustomer"
      DataSource      =   "VC"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H80000004&
      DataField       =   "Name"
      DataMember      =   "findCustomer"
      DataSource      =   "VC"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtLastName 
      BackColor       =   &H80000004&
      DataField       =   "LastName"
      DataMember      =   "findCustomer"
      DataSource      =   "VC"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtMovieRented 
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtProduct 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label lblReturnID 
      Alignment       =   1  'Right Justify
      Caption         =   "Movie Returned"
      Height          =   255
      Left            =   2400
      TabIndex        =   41
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblVideoName 
      Alignment       =   1  'Right Justify
      Caption         =   "Video Name"
      Height          =   255
      Left            =   2400
      TabIndex        =   32
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPOrder 
      Alignment       =   1  'Right Justify
      Caption         =   "Products Selected"
      Height          =   255
      Left            =   2160
      TabIndex        =   28
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblVideos 
      Alignment       =   1  'Right Justify
      Caption         =   "Movies Selected"
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4785
      TabIndex        =   24
      Top             =   6720
      Width           =   435
   End
   Begin VB.Label lblCustomerID 
      Alignment       =   1  'Right Justify
      Caption         =   "CustomerID"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLastName 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblMovie 
      Alignment       =   1  'Right Justify
      Caption         =   "Movie Rented"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblProduct 
      Alignment       =   1  'Right Justify
      Caption         =   "Product"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VIDEO CLUB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmVideoClub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intQuestion As Integer
Dim WithEvents cn As Connection
Attribute cn.VB_VarHelpID = -1
Dim WithEvents adoRS As ADODB.Recordset
Attribute adoRS.VB_VarHelpID = -1
Dim WithEvents conn As Connection
Attribute conn.VB_VarHelpID = -1
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
'Dim timer1 As Integer
Function closeConnection()

    If (cn <> Null) Then
        cn.Close
        Set cn = Nothing
    End If
    
End Function

Public Function connDB()
       
    ' connects to a DB
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

Private Sub cmdBuyP_Click()
        
    Call clearForm
    Call showProduct
        
End Sub

Private Sub cmdCFind_Click()
    
    Call clearForm
    
    lblName.Visible = True
    txtName.Visible = True
    txtName.BackColor = &H80000005 ' white
    txtName.Enabled = True
    lblLastName.Visible = True
    txtLastName.Visible = True
    txtLastName.BackColor = &H80000005 ' white
    txtLastName.Enabled = True
    txtName.SetFocus
    
End Sub

Private Sub cmdCVideo_Click()
    txtVideoName.Text = ""
End Sub

Private Sub cmdExit_Click()
    
    Call closeConnection
    
    Unload Me
    End
    
End Sub

Private Sub cmdFindVideo_Click()
    
    Call clearForm
    
    lblVideoName.Visible = True
    txtSVideo.Visible = True
    txtSVideo.SetFocus
    
End Sub

Private Sub cmdMoreVideos_Click()
    
    Dim i As Integer
    Call closeConnection
    Call connDB

    ' if somebody understands this mess ....
    adoRS.ActiveConnection = cn
               
    If (IsNumeric(txtMovieRented.Text) = True) Then
        adoRS.Source = "Select * From tblVideo WHERE VideoID =" _
                        & Val(txtMovieRented.Text)
        adoRS.Open
        If Not (adoRS.EOF) Then
            If (adoRS.Fields("RentedTo").Value = -1) Then
                lstMovies.AddItem adoRS.Fields("VideoName").Value
                lstMovieID.AddItem adoRS.Fields("VideoID").Value
                lstMoviePrice.AddItem adoRS.Fields("Price").Value
        
                txtMovieRented.Text = ""
                txtMovieRented.SetFocus
            Else
                lblMSG.Caption = "Selected movie is rented to" & adoRS.Fields("RentedTo").Value
                txtMovieRented.Text = ""
                txtMovieRented.SetFocus
            End If
        Else
            i = MsgBox("Video does not exist", vbCritical, "No Video Found")
            txtMovieRented.Text = ""
            txtMovieRented.SetFocus
        End If
            
    Else

        adoRS.Source = "SELECT * FROM tblVideo WHERE VideoName ='" _
                        & txtMovieRented.Text & "'"
        adoRS.Open
        If Not (adoRS.EOF) Then
            'here can you add a loop to check if there are
            'more movies with the same name/title
            'we assume that there is just one copy
            If (adoRS.Fields("RentedTo").Value = -1) Then
                lstMovies.AddItem adoRS.Fields("VideoName").Value
                lstMovieID.AddItem adoRS.Fields("VideoID").Value
                lstMoviePrice.AddItem adoRS.Fields("Price").Value
        
                txtMovieRented.Text = ""
                txtMovieRented.SetFocus
            Else
                lblMSG.Caption = "Selected movie is rented to" & adoRS.Fields("RentedTo").Value
                txtMovieRented.Text = ""
                txtMovieRented.SetFocus
            End If
        Else
            i = MsgBox("Video does not exist", vbCritical, "No Video Found")
            txtMovieRented.Text = ""
            txtMovieRented.SetFocus
        End If
                            
    End If
    
    
End Sub

Private Sub cmdNCancel_Click()
    txtName.Text = ""
    txtLastName.Text = ""
End Sub

Private Sub cmdOKVideo_Click()
    ' serch if there is a video with that name
    ' and check if there is a free copy to rent
    Call closeConnection
    Call connDB
    
    adoRS.ActiveConnection = cn
    adoRS.Source = "SELECT * FROM tblVideo WHERE VideoName = '" _
                    & txtSVideo.Text & "'"
    adoRS.Open
                     
    If Not (adoRS.EOF Or adoRS.BOF) Then
        If (adoRS.Fields("RentedTo").Value = -1) Then
            lblMSG.Caption = "Video ID: '" & adoRS.Fields("VideoID").Value _
                        & "' Rented: NO"
        Else
            lblMSG.Caption = "Video ID: '" & adoRS.Fields("VideoID").Value _
                        & "' Rented to customer: " & adoRS.Fields("RentedTo").Value
        End If
    Else
        lblMSG.Caption = "The movie does not exist in the database"
    End If
    
End Sub

Private Sub cmdNextCID_Click()
    If Not (adoRS.EOF) Then
        adoRS.MoveNext
    End If
    If Not (adoRS.EOF) Then
        lblMSG.Caption = "Your ID is: " & adoRS.Fields("CustomerID").Value
        lblAddress.Visible = True
        txtAddress.Visible = True
        txtAddress.Text = adoRS.Fields("Address").Value
        
    End If
End Sub

Private Sub cmdOKName_Click()
    closeConnection
    connDB
 
    If (txtName <> "") Then
        adoRS.ActiveConnection = cn
        adoRS.Source = "Select * From tblCustomer WHERE Name ='" _
                        & txtName.Text & "' AND LastName ='" _
                        & txtLastName.Text & "'"
        adoRS.Open
    
    Else
        adoRS.ActiveConnection = cn
        adoRS.Source = "Select * From tblCustomer WHERE LastName ='" _
                        & txtLastName.Text & "'"
        adoRS.Open
    End If
    
    If Not (adoRS.EOF) Then
    
        lblMSG.Caption = "Your ID is: " & adoRS.Fields("CustomerID").Value
        lblAddress.Visible = True
        txtAddress.Visible = True
        txtAddress.Text = adoRS.Fields("Address").Value
        
        If (adoRS.RecordCount <> 1) Then
            cmdNextCID.Visible = True
        End If
    Else
            lblMSG.Caption = "Customer does not exist"
    End If
    
End Sub

Private Sub cmdNextVideo_Click()
    If Not (adoRS.EOF) Then
        adoRS.MoveNext
    End If
    If Not (adoRS.EOF) Then
            lblMSG.Caption = "Video ID: " & adoRS.Fields("VideoID").Value _
                            & ", Rented: " & adoRS.Fields("Rented").Value
    End If
    
End Sub

Private Sub cmdPMore_Click()
    
    Dim i As Integer
    Call closeConnection
    Call connDB

    ' add more products
    If (Str(Val(txtCustomerID.Text)) = " " & Trim(txtCustomerID.Text)) Then
        If (txtProduct.Text <> "") Then
            adoRS.ActiveConnection = cn
            adoRS.Source = "Select * From tblProduct WHERE ProductID =" _
                            & Val(txtProduct.Text)
            adoRS.Open
            If (adoRS.EOF = False And adoRS.BOF = False) Then
                lstProducts.AddItem adoRS.Fields("ProductName").Value
                lstProductPrice.AddItem adoRS.Fields("Price").Value
            Else
                i = MsgBox("Product does not exist", vbCritical, "No Product Found")
                txtProduct.Text = ""
                txtProduct.SetFocus
            End If
        
        End If
    Else
        intCreate = MsgBox("Product should be a number here", vbExclamation, _
                    "ProductID is not a number")
        txtProduct.Text = ""
        txtProduct.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub cmdRegCustomer_Click()
    
    Call closeConnection
    Call connDB
    Dim dblNextID As Double
    frmRegCustomer.Left = 100
    frmRegCustomer.Top = 100
    frmRegCustomer.Show
    
    adoRS.ActiveConnection = cn
    adoRS.Source = "SELECT * FROM tblCustomer"
    adoRS.Open
    
    adoRS.MoveLast
    
    dblNextID = adoRS.Fields("CustomerID").Value + 1
    frmRegCustomer.txtID = dblNextID
    
End Sub

Private Sub cmdRegProdukt_Click()
    frmRegProduct.Show
End Sub

Private Sub cmdRegVideo_Click()
    frmRegVideo.Show
End Sub

Private Sub cmdRent_Click()
    Call clearForm
    showCID
    txtCustomerID.SetFocus
End Sub

Private Sub cmdReturn_Click()
    Call clearForm
    Call showReturn
    txtReturnMID.SetFocus
End Sub

Private Sub cmdShowOrders_Click()
    frmOrders.Show
End Sub

Private Sub cmdShowRented_Click()
    Dim strVideos As String
    Dim VID As Integer
    Dim intE
    Call closeConnection
    Call connDB

    
    adoRS.ActiveConnection = cn
    adoRS.Source = "Select * From tblOrder WHERE CustomerID =" _
                    & Val(txtCustomerID.Text)
    adoRS.Open
    Do While Not (adoRS.EOF)
        If Not (adoRS.EOF) Then
            VID = adoRS.Fields("VideoID").Value
            strVideos = strVideos & " " & findVName(VID) & ","
            
            adoRS.MoveNext
        End If
        
    Loop
        intE = MsgBox(strVideos, vbInformation, "Customer has NOT returned these videos")
End Sub

Private Sub cmdStartUp_Click()
    Call clearForm
    Call showCID
End Sub

Private Sub Form_Load()
    ' start up calls
    Call clearForm
    Call showCID
    Call connDB
End Sub

Private Sub cmdCID_Click()
    Call closeConnection
    Call connDB
    Dim intCreate As Integer
    ' get data

    'test if CustomerID is a number
    'didn't use isNumeric because we want a simple nice number here
    If (Str(Val(txtCustomerID.Text)) = " " & Trim(txtCustomerID.Text)) Then
        adoRS.ActiveConnection = cn
        adoRS.Source = "Select * From tblCustomer WHERE CustomerID =" _
                        & Val(txtCustomerID.Text)
        adoRS.Open
    Else
        intCreate = MsgBox("CustomerID should be a number here", vbExclamation, _
                    "CustomerID is not a number")
        txtCustomerID.Text = ""
        txtCustomerID.SetFocus
        Exit Sub
    End If
    
    If (adoRS.EOF = False And adoRS.BOF = False) Then
        txtName.Visible = True
        txtName.Text = adoRS.Fields("Name").Value
        txtName.Enabled = False
        txtName.BackColor = &H80000004
        txtLastName.Visible = True
        txtLastName.Text = adoRS.Fields("LastName").Value
        txtLastName.Enabled = False
        txtLastName.BackColor = &H80000004
        txtAddress.Visible = True
        txtAddress.Text = adoRS.Fields("Address").Value
        txtAddress.Enabled = False
        txtAddress.BackColor = &H80000004
        lblName.Visible = True
        lblLastName.Visible = True
        lblAddress.Visible = True
        
        Call showMovies
        txtMovieRented.SetFocus
        txtCustomerID.Enabled = False
        txtCustomerID.BackColor = &H80000004
        
    Else
        ' register new user
        intCreate = MsgBox("ID does not exist. Do you want to create a new user?", _
                            vbQuestion + vbYesNo + vbDefaultButton2, "Create New User")
        If (intCreate = vbYes) Then
            Call cmdRegCustomer_Click
        Else
            txtCustomerID.Text = ""
            txtCustomerID.SetFocus
                        
        End If
    End If
End Sub

Private Sub cmdFinishOrder_Click()
    Dim curPrice As Currency
    curPrice = 0

    For i = 0 To lstMoviePrice.ListCount - 1
        curPrice = curPrice + Val(lstMoviePrice.List(i))
    Next
    For i = 0 To lstProductPrice.ListCount - 1
        curPrice = curPrice + Val(lstProductPrice.List(i))
    Next
    
    intQuestion = MsgBox("The Price is: " & curPrice _
    & ". Is this correct", vbQuestion + vbYesNo + vbDefaultButton1, _
    "Rent Request")
    ' test
    If (intQuestion = vbYes) Then
        ' register order
        Call regOrder
        Call clearForm
        Call showCID
        txtCustomerID.SetFocus
        
    Else
        txtMovieRented.SetFocus
    End If
End Sub

Private Sub lstMovies_DblClick()
    Dim i As Integer
    i = lstMovies.ListIndex
    lstMovies.RemoveItem (i)
    lstMovieID.RemoveItem (i)
    lstMoviePrice.RemoveItem (i)
    txtMovieRented.SetFocus
End Sub

Private Sub lstProducts_DblClick()
    Dim i As Integer
    i = (lstProducts.ListIndex)
    lstProducts.RemoveItem (i)
    lstProductPrice.RemoveItem (i)
    txtProduct.SetFocus
End Sub

Private Sub txtCustomerID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdCID_Click
        Call checkRented
    End If
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdOKName_Click
    End If
End Sub

Private Sub txtMovieRented_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If (txtMovieRented <> "") Then
            Call cmdMoreVideos_Click
        Else
            txtMovieRented.Text = ""
            Call showProduct
            txtProduct.SetFocus
        End If
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLastName.SetFocus
    End If
End Sub

Private Sub txtProduct_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If (txtProduct <> "") Then
            Call cmdPMore_Click
            txtProduct.Text = ""
            txtProduct.SetFocus
        Else
            Call cmdFinishOrder_Click
        End If
    End If
End Sub

Private Sub txtSVideo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdOKVideo_Click
    End If
End Sub

Private Sub txtReturnMID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If (txtReturnMID <> "") Then
            Call returnVideo
        Else
            Call clearForm
            Call showCID
            txtCustomerID.SetFocus
        End If
    End If
End Sub

Function regOrder()
    Dim i As Integer

    'Register order
    For i = 0 To lstMovieID.ListCount - 1
        
        Call closeConnection
        Call connDB
        
        adoRS.ActiveConnection = cn
        adoRS.Source = "INSERT INTO tblOrder (CustomerID, VideoID, orderDate) VALUES ('" _
                        & Val(txtCustomerID.Text) & "','" & Val(lstMovieID.List(i)) & "','" _
                        & Date & "')"
        adoRS.Open
        
        adoRS.ActiveConnection = cn
        adoRS.Source = "UPDATE tblVideo SET RentedTo =" & Val(txtCustomerID.Text) _
                        & " WHERE VideoID =" & Val(lstMovieID.List(i))
        adoRS.Open
    Next
    
End Function

Function returnVideo()
    
    Dim CID As Integer
    CID = getCID()
    
    Call closeConnection
    Call connDB

    If (CID <> 0) Then
        adoRS.ActiveConnection = cn
        adoRS.Source = "DELETE FROM tblOrder WHERE VideoID =" & Val(txtReturnMID.Text)
        adoRS.Open
        
        lblMSG.Caption = "Movie " & txtReturnMID.Text & " Returned"
        
        adoRS.ActiveConnection = cn
        adoRS.Source = "UPDATE tblVideo SET RentedTo = -1 WHERE VideoID =" & Val(txtReturnMID.Text)
        adoRS.Open
        
        txtReturnMID.Text = ""
        txtReturnMID.SetFocus
    Else
        txtReturnMID.Text = ""
        txtReturnMID.SetFocus
    End If
    
End Function

Function getCID() As Integer
    
    Call closeConnection
    Call connDB
    adoRS.ActiveConnection = cn
    adoRS.Source = "SELECT * FROM tblOrder WHERE VideoID=" & Val(txtReturnMID.Text)
    adoRS.Open

    If Not (adoRS.EOF) Then
        getCID = adoRS.Fields("CustomerID").Value
    Else
        lblMSG.Caption = "Movie isn't rented"
    End If
    
End Function

Function findVName(VID As Integer) As String
    
    ' connects to a DB
    Dim strConnect As String
    Dim strProvider As String
    Dim strDataSource As String
    Dim strDataBaseName As String
    
    strProvider = "Provider= Microsoft.Jet.OLEDB.3.51;"
    strDataSource = App.Path
    strDataBaseName = "\video.mdb;"
    strDataSource = "Data Source=" & strDataSource & strDataBaseName
    strConnect = strProvider & strDataSource
    
    Set conn = New ADODB.Connection
    conn.Open strConnect
    
    Set rs = New ADODB.Recordset
    
    rs.CursorType = adOpenStatic
    rs.CursorLocation = adUseClient
    rs.LockType = adLockPessimistic
    
    rs.ActiveConnection = cn
    rs.Source = "SELECT * FROM tblVideo WHERE VideoID =" & VID
    rs.Open
    
    findVName = rs.Fields("VideoName").Value
    
End Function

Function checkRented()
    
    'check if the customer has returned all movies he/she has rented
    Call closeConnection
    Call connDB

    adoRS.ActiveConnection = cn
    adoRS.Source = "SELECT * FROM tblOrder WHERE CustomerID =" _
                    & Val(txtCustomerID.Text)
    adoRS.Open
    If (adoRS.RecordCount > 0) Then
        cmdShowRented.Visible = True
    End If
    
End Function

