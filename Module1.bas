Attribute VB_Name = "Module1"
Public Function clearForm()
    
    With frmVideoClub
        ' clears the form and gets ready to take the next order
        clearMControl
        clearPControl
        clearCSControl
        clearMSControl
        clearCID
        clearReturn
        
        .lblMSG.Caption = ""
        
    End With
    
End Function

Public Function clearCID()
    With frmVideoClub
        
        .lblCustomerID.Visible = False
        .txtCustomerID.Visible = False
        .cmdShowRented.Visible = False
    
    End With
    
End Function

Public Function clearMControl()
    With frmVideoClub
        
        'clears "movies-control"
        .lblMovie.Visible = False
        .txtMovieRented.Visible = False
        .txtMovieRented.Text = ""
        .lstMovies.Visible = False
        .lstMovies.Clear
        .lstMoviePrice.Clear
        .lstMovieID.Clear
        
    End With
    
End Function

Public Function clearPControl()
    With frmVideoClub
        
        'clears the "product-controls"
        .lblProduct.Visible = False
        .txtProduct.Visible = False
        .txtProduct.Text = ""
        .cmdFinishOrder.Visible = False
        .lstProducts.Visible = False
        .lstProducts.Clear
        .lstProductPrice.Clear
        
    End With
    
End Function

Public Function clearMSControl()
    With frmVideoClub
        
        'clears the "moviesearch-control"
        .txtSVideo.Text = ""
        .lblVideoName.Visible = False
        .txtSVideo.Visible = False
        
    End With
    
End Function

Public Function clearCSControl()
    With frmVideoClub
        
        'clears the "customersearch-control"
        .lblName.Visible = False
        .txtName.Visible = False
        .txtName.Text = ""
        .lblLastName.Visible = False
        .txtLastName.Visible = False
        .txtLastName.Text = ""
        .lblAddress.Visible = False
        .txtAddress.Visible = False
        .txtAddress.Text = ""
        .cmdNextCID.Visible = False
               
    End With
    
End Function

Public Function clearReturn()

    With frmVideoClub
        
        .lblReturnID.Visible = False
        .txtReturnMID.Visible = False
                      
    End With
    
End Function
Public Function showCID()
    
    With frmVideoClub
        
        .txtCustomerID.Text = ""
        .txtCustomerID.BackColor = &H80000005 ' white
        .txtCustomerID.Enabled = True
        .txtCustomerID.Visible = True
        .lblCustomerID.Visible = True
        
    End With

End Function

Public Function showProduct()
    With frmVideoClub
        
        .lblProduct.Visible = True
        .txtProduct.Visible = True
        .cmdFinishOrder.Visible = True
        .lstProducts.Visible = True
        .txtProduct.SetFocus
    
    End With
        
End Function

Public Function showMovies()

    With frmVideoClub
        
        .lblMovie.Visible = True
        .txtMovieRented.Visible = True
        
        .lstMovies.Visible = True
        .txtMovieRented.SetFocus
                        
    End With
    
End Function

Public Function showReturn()

    With frmVideoClub
        
        .txtReturnMID.Visible = True
        .lblReturnID.Visible = True
        
    End With
    
End Function

