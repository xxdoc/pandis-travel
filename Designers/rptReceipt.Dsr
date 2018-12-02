VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptReceipt 
   Caption         =   "Απόδειξη"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   16080
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   28363
   _ExtentY        =   18759
   SectionData     =   "rptReceipt.dsx":0000
End
Attribute VB_Name = "rptReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intReceiptCount As Integer
Dim strCustomersOrSuppliers As String

Private Sub ActiveReport_DataInitialize()

    Fields.RemoveAll
    
    Fields.Add "CompanyData"
    Fields.Add "Date"
    Fields.Add "Batch"
    Fields.Add "ReceiptDescription"
    Fields.Add "ReceiptNo"
    Fields.Add "Amount"
    Fields.Add "Description"
    Fields.Add "Profession"
    Fields.Add "Address"
    Fields.Add "TaxNo"
    Fields.Add "TaxOfficeDescription"
    Fields.Add "Reason"
    Fields.Add "PaymentWayDescription"
    Fields.Add "BankDescription"
    Fields.Add "FullNumber"

End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    On Error GoTo ErrTrap
    
    'Local variables
    intReceiptCount = intReceiptCount + 1
    
    If intReceiptCount > 2 Then
        EOF = True
        Exit Sub
    End If
    
    With PersonsTransactions
        'Στοιχεία εταιρίας
        Fields("CompanyData") = arrCompanyData(1) & Chr(13) & arrCompanyData(2) & Chr(13) & arrCompanyData(3) & Chr(13) & arrCompanyData(4) & Chr(13) & arrCompanyData(5) & Chr(13) & arrCompanyData(6)
        'Παραστατικό
        Fields("Date") = .mskDateIssue.text
        Fields("Batch") = .lblCodeBatch.Caption
        Fields("ReceiptDescription") = .lblCodeDescription.Caption
        Fields("ReceiptNo") = "Νο " & .txtInvoiceNo.text
        'Ποσό
        Fields("Amount") = .mskAmount.text
        
        'Βρίσκω τα στοιχεία του πελάτη
        Dim strSQL As String
        Dim rstRecordset As Recordset
        
        strSQL = "SELECT Description, Profession, Address, TaxNo, TaxOfficeDescription " _
        & "FROM " & .txtCustomersOrSuppliers.text & " " _
        & "INNER JOIN TaxOffices ON " & .txtCustomersOrSuppliers.text & ".TaxOfficeID = TaxOffices.TaxOfficeID " _
        & "WHERE " & .txtCustomersOrSuppliers.text & ".ID = " & PersonsTransactions.txtInvoicePersonID.text
        
        Set TempQuery = CommonDB.CreateQueryDef(""): TempQuery.SQL = strSQL
        Set rstRecordset = TempQuery.OpenRecordset()
        
        'Εισπραξη ή πληρωμή
        lblPaymentInOrPaymentOut.Caption = IIf(.txtCustomersOrSuppliers.text = "Customers", "ΕΙΣΠΡΑΞΑΜΕ ΑΠΟ", "ΠΛΗΡΩΣΑΜΕ ΣΕ")
        
        'Πελάτης
        Fields("Description") = rstRecordset!Description
        Fields("Profession") = rstRecordset!Profession
        Fields("Address") = rstRecordset!Address
        Fields("TaxNo") = rstRecordset!TaxNo
        Fields("TaxOfficeDescription") = rstRecordset!TaxOfficeDescription
        
        'Λεπτομέρειες κίνησης
        Fields("Reason") = .txtReason.text
        Fields("PaymentWayDescription") = .txtPaymentWayDescription.text
        Fields("BankDescription") = .txtBankDescription.text
        Fields("FullNumber") = .lblFullNumber.Caption
        
        EOF = False
        
    End With
    
    Exit Sub
    
ErrTrap:
    If Err.Number = 6 Then
        Resume Next
    Else
        DisplayErrorMessage True, Err.Description
    End If
    
End Sub

