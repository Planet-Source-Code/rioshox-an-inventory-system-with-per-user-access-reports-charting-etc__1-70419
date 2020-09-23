Attribute VB_Name = "modSystemSpecifics"
Option Explicit
Sub Main()
    On Error GoTo ERR_Main
    
    If App.PrevInstance Then End
    
    connRFM.Open "Provider=SQLOLEDB;Data Source=.;Initial Catalog=RFM;User Id=sa;Password="
    
    
    frmLogIn.Show 1
    MainForm.Show
    
        
ERR_Main:
    If Err.Number <> 0 Then
        MsgBox "Errors occurred, cannot launch the system!" & vbCrLf & Err.Description, vbCritical
        End
    End If

End Sub


Public Sub HideSubMenus()
    frmSubMenuProducts.Hide
    frmSubMenuSalesOrder.Hide
    frmSubMenuCustomers.Hide
    frmSubMenuPackagingTypes.Hide
End Sub

Public Sub DeauthorizeNotify()
    frmTopBar.Show
    frmTopBar.lblBigNo.Visible = True
    frmTopBar.lblSmallNo.Visible = True
    frmTopBar.Timer1.Enabled = True
End Sub

Public Sub SetPriviledges()
    Dim rsSystemUsers As New ADODB.Recordset
    
    rsSystemUsers.Open "SELECT * FROM tblSystemUSers WHERE EmployeeID = " & g_CurrentUser.EmployeeID, connRFM, adOpenStatic
    
    With g_CurrentUser
        .Products = rsSystemUsers("Products")
        .Customer = rsSystemUsers("Customer")
        .UserAccess = rsSystemUsers("UserAccess")
        .ApproveSO = rsSystemUsers("ApproveSO")
        .ApproveShip = rsSystemUsers("ApproveShip")
        .BackupData = rsSystemUsers("BackupData")
        .ReturnItem = rsSystemUsers("ReturnItem")
        .ManualSO = rsSystemUsers("ManualSO")
    End With
    
    ClearRS rsSystemUsers
End Sub

Public Sub Generate4YrData(dfrm As Form)
    Dim rsTransactions As New ADODB.Recordset
    Dim rsTransactionDetails As New ADODB.Recordset
    Dim rsCustomers As New ADODB.Recordset
    Dim rsProducts As New ADODB.Recordset
    Dim CurrentDate As Date
    Dim CurrentCustomer As String, CurrentProduct
    
    CurrentDate = #1/1/2001#
    
    
    rsTransactions.CursorLocation = adUseServer
    rsTransactionDetails.CursorLocation = adUseServer
    
    rsTransactions.Open "SELECT * FROM tblTransactions", connRFM, adOpenDynamic, adLockOptimistic
    rsTransactionDetails.Open "SELECT * FROM tblTransactionDetails", connRFM, adOpenDynamic, adLockOptimistic
    rsCustomers.Open "SELECT CustomerID FROM tblCustomers", connRFM, adOpenStatic
    rsProducts.Open "SELECT ProductID, UnitPrice FROM tblProducts", connRFM, adOpenStatic
    
    
    
    
    
    
    While CurrentDate < #1/20/2007#
    
        If rsCustomers.EOF Then rsCustomers.MoveFirst
        
    
        If rsProducts.EOF Then rsProducts.MoveFirst
        
        
        dfrm.Caption = CurrentDate
        
        rsTransactions.AddNew
            rsTransactions("CustomerID") = rsCustomers("CustomerID")
            rsTransactions("OrderDate") = CurrentDate
            rsTransactions("RequiredDate") = CurrentDate + Format(Time, "ss")
            rsTransactions("EmployeeID") = GenerateEndorser
            rsTransactions("ShippedDate") = CurrentDate + Day(Date)
        rsTransactions.Update
        
        
        
        
        rsTransactionDetails.AddNew
            rsTransactionDetails("TransactionID") = 1 'GetLastTransactionID
            rsTransactionDetails("ProductID") = rsProducts("ProductID")
            rsTransactionDetails("UnitPrice") = rsProducts("UnitPrice")
            rsTransactionDetails("Quantity") = Format(Time, "ss")
            rsTransactionDetails("Discount") = 0
        rsTransactionDetails.Update
        rsProducts.MoveNext
        If rsProducts.EOF Then rsProducts.MoveFirst
        
        rsTransactionDetails.AddNew
            rsTransactionDetails("TransactionID") = 1 'GetLastTransactionID
            rsTransactionDetails("ProductID") = rsProducts("ProductID")
            rsTransactionDetails("UnitPrice") = rsProducts("UnitPrice")
            rsTransactionDetails("Quantity") = Format(Time, "ss")
            rsTransactionDetails("Discount") = 0
        rsTransactionDetails.Update
        rsProducts.MoveNext
        If rsProducts.EOF Then rsProducts.MoveFirst
        
        rsTransactionDetails.AddNew
            rsTransactionDetails("TransactionID") = 1 ' GetLastTransactionID
            rsTransactionDetails("ProductID") = rsProducts("ProductID")
            rsTransactionDetails("UnitPrice") = rsProducts("UnitPrice")
            rsTransactionDetails("Quantity") = Format(Time, "ss")
            rsTransactionDetails("Discount") = 0
        rsTransactionDetails.Update
        rsProducts.MoveNext
        If rsProducts.EOF Then rsProducts.MoveFirst
        
        rsTransactionDetails.AddNew
            rsTransactionDetails("TransactionID") = 1 ' GetLastTransactionID
            rsTransactionDetails("ProductID") = rsProducts("ProductID")
            rsTransactionDetails("UnitPrice") = rsProducts("UnitPrice")
            rsTransactionDetails("Quantity") = Format(Time, "ss")
            rsTransactionDetails("Discount") = 0
        rsTransactionDetails.Update
        
        rsProducts.MoveNext
        If rsProducts.EOF Then rsProducts.MoveFirst
        
        
        rsTransactionDetails.AddNew
            rsTransactionDetails("TransactionID") = 1 'GetLastTransactionID
            rsTransactionDetails("ProductID") = rsProducts("ProductID")
            rsTransactionDetails("UnitPrice") = rsProducts("UnitPrice")
            rsTransactionDetails("Quantity") = Format(Time, "ss")
            rsTransactionDetails("Discount") = 0
        rsTransactionDetails.Update
        
        
        
        CurrentDate = CurrentDate + 1
    Wend
    
End Sub

Private Function GenerateEndorser() As Integer
    Select Case Format(Time, "ss")
        Case 1 To 15
            GenerateEndorser = 1
        Case 3 To 30
            GenerateEndorser = 2
        Case 31 To 45
            GenerateEndorser = 3
        Case 46 To 60
            GenerateEndorser = 4
    End Select
End Function



Public Sub Reprice(dfrm As Form)
    Dim rs As New ADODB.Recordset
    
    rs.Open "SELECT dbo.tblTransactionDetails.UnitPrice, dbo.tblTransactions.OrderDate FROM dbo.tblTransactions INNER JOIN dbo.tblTransactionDetails ON dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID WHERE ORDERDATE<='12/31/2006' ORDER BY OrderDate", connRFM, adOpenDynamic, adLockOptimistic
    
    rs.MoveFirst
    While Not rs.EOF
        dfrm.Caption = rs("ORderdate")
        rs("UnitPrice") = rs("UnitPrice") / (Year(Date) - Year(rs("ORderdate")))
        rs.Update
        rs.MoveNext
    Wend
    
    MsgBox "done", vbInformation
End Sub


Public Sub UpdateBuyer(dfrm As Form)
    Dim rsCustomers As New ADODB.Recordset
    Dim rsTransactions As New ADODB.Recordset
    
    rsCustomers.Open "SELECT CustomerID FROM tblCustomers", connRFM
    rsTransactions.Open "SELECT TransactionID,CustomerID FROM tblTransactions", connRFM, adOpenDynamic, adLockOptimistic
    
    
        While Not rsCustomers.EOF
            If rsTransactions.EOF Then GoTo EndNow
            rsTransactions("CustomerID") = rsCustomers("CustomerID")
            rsCustomers.MoveNext
            rsTransactions.MoveNext
            If rsCustomers.EOF Then rsCustomers.MoveFirst
            dfrm.Caption = rsTransactions("TransactionID")
        Wend
    
    
EndNow:
    MsgBox "Done", vbInformation
    
End Sub
