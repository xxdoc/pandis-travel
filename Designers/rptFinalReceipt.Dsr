VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptFinalReceipt 
   Caption         =   "Απόδειξη"
   ClientHeight    =   13935
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   14220
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   25083
   _ExtentY        =   24580
   SectionData     =   "rptFinalReceipt.dsx":0000
End
Attribute VB_Name = "rptFinalReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intReceiptCount As Integer

Private Sub ActiveReport_DataInitialize()

    fields.RemoveAll
    
    fields.Add "CompanyData"
    fields.Add "Date"
    fields.Add "Batch"
    fields.Add "ReceiptDescription"
    fields.Add "ReceiptNo"
    fields.Add "Amount"
    fields.Add "CompanyDescription"
    fields.Add "CompanyProfession"
    fields.Add "CompanyAddress"
    fields.Add "CompanyTaxNo"
    fields.Add "TaxOfficeDescription"
    fields.Add "Reason"
    fields.Add "PaymentWayDescription"
    fields.Add "BankDescription"

End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    On Error GoTo ErrTrap
    
    'Local variables
    intReceiptCount = intReceiptCount + 1
    
    If intReceiptCount > 2 Then
        EOF = True
        Exit Sub
    End If
    
    'If EOF Then Exit Sub
        
    With PersonsTransactions
        'Στοιχεία εταιρίας
        fields("CompanyData") = arrCompanyData(1) & Chr(13) & arrCompanyData(2) & Chr(13) & arrCompanyData(3) & Chr(13) & arrCompanyData(4) & Chr(13) & arrCompanyData(5) & Chr(13) & arrCompanyData(6)
        'Παραστατικό
        fields("Date") = " ΗΜΕΡΟΜΗΝΙΑ ΕΚΔΟΣΗΣ: " & .mskDate.text
        fields("Batch") = .lblCodeBatch.Caption
        fields("ReceiptDescription") = .lblCodeDescription.Caption
        fields("ReceiptNo") = "Νο " & Right("00000" & .txtInvoiceNo.text, 5)
        'Ποσό
        fields("Amount") = .mskAmount.text
        
        'Βρίσκω τα στοιχεία του πελάτη
        Dim strSQL As String
        Dim rstRecordset As Recordset
        
        strSQL = "SELECT CompanyDescription, CompanyProfession, CompanyAddress, CompanyTaxNo, TaxOfficeDescription " _
        & "FROM Companies " _
        & "INNER JOIN TaxOffices ON Companies.CompanyTaxOfficeID = TaxOffices.TaxOfficeID " _
        & "WHERE Companies.CompanyID = " & PersonsTransactions.txtCompanyID.text
        
        Set TempQuery = CommonDB.CreateQueryDef(""): TempQuery.SQL = strSQL
        Set rstRecordset = TempQuery.OpenRecordset()
        
        'Πελάτης
        fields("CompanyDescription") = rstRecordset!CompanyDescription
        fields("CompanyProfession") = rstRecordset!CompanyProfession
        fields("CompanyAddress") = rstRecordset!CompanyAddress
        fields("CompanyTaxNo") = rstRecordset!CompanyTaxNo
        fields("TaxOfficeDescription") = rstRecordset!TaxOfficeDescription
        
        'Λεπτομέρειες κίνησης
        fields("Reason") = .txtReason.text
        fields("PaymentWayDescription") = .txtPaymentWayDescription.text
        fields("BankDescription") = .txtBankDescription.text
        
        'Ολογράφως
        '
        'Ολογράφως
        
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

