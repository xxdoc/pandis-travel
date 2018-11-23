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

    fields.RemoveAll
    
    fields.Add "CompanyData"
    fields.Add "Date"
    fields.Add "Batch"
    fields.Add "ReceiptDescription"
    fields.Add "ReceiptNo"
    fields.Add "Amount"
    fields.Add "Description"
    fields.Add "Profession"
    fields.Add "Address"
    fields.Add "TaxNo"
    fields.Add "TaxOfficeDescription"
    fields.Add "Reason"
    fields.Add "PaymentWayDescription"
    fields.Add "BankDescription"
    fields.Add "FullNumber"

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
        fields("CompanyData") = arrCompanyData(1) & Chr(13) & arrCompanyData(2) & Chr(13) & arrCompanyData(3) & Chr(13) & arrCompanyData(4) & Chr(13) & arrCompanyData(5) & Chr(13) & arrCompanyData(6)
        'Παραστατικό
        fields("Date") = .mskDateIssue.text
        fields("Batch") = .lblCodeBatch.Caption
        fields("ReceiptDescription") = .lblCodeDescription.Caption
        fields("ReceiptNo") = "Νο " & .txtInvoiceNo.text
        'Ποσό
        fields("Amount") = .mskAmount.text
        
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
        fields("Description") = rstRecordset!description
        fields("Profession") = rstRecordset!Profession
        fields("Address") = rstRecordset!Address
        fields("TaxNo") = rstRecordset!taxNo
        fields("TaxOfficeDescription") = rstRecordset!TaxOfficeDescription
        
        'Λεπτομέρειες κίνησης
        fields("Reason") = .txtReason.text
        fields("PaymentWayDescription") = .txtPaymentWayDescription.text
        fields("BankDescription") = .txtBankDescription.text
        fields("FullNumber") = .lblFullNumber.Caption
        
        EOF = False
        
    End With
    
    Exit Sub
    
ErrTrap:
    If Err.Number = 6 Then
        Resume Next
    Else
        DisplayErrorMessage True, Err.description
    End If
    
End Sub

