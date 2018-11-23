VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCoachesReport 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11895
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   16080
   Icon            =   "rptCoachesReport.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   28363
   _ExtentY        =   20981
   SectionData     =   "rptCoachesReport.dsx":1CFA
End
Attribute VB_Name = "rptCoachesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRow As Long
Private blnLastPage As Boolean

Private Sub ActiveReport_DataInitialize()

    Fields.RemoveAll
    
    Fields.Add "TransferDate"
    Fields.Add "RouteDescription"
    Fields.Add "PickupPointHotelDescription"
    Fields.Add "CompanyDescription"
    Fields.Add "TransferAdults"
    Fields.Add "TransferKids"
    Fields.Add "TransferFree"
    Fields.Add "TransferTotal"
    Fields.Add "PickUpPointTime"
    Fields.Add "PickUpPointExactPoint"
    Fields.Add "TransferRemarks"
    Fields.Add "TransferDestination"
    
    lngRow = 0
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    lngRow = lngRow + 1
    
    With CoachesReport.grdCoachesReport
        
        If lngRow > .RowCount Then
            blnLastPage = True
            EOF = True
            Exit Sub
        End If
        
        Fields("TransferDate") = .CellValue(lngRow, "TransferDate")
        Fields("RouteDescription") = .CellValue(lngRow, "RouteDescription")
        Fields("PickupPointHotelDescription") = .CellValue(lngRow, "PickupPointHotelDescription")
        Fields("CompanyDescription") = .CellValue(lngRow, "CompanyDescription")
        Fields("TransferAdults") = .CellValue(lngRow, "TransferAdults")
        Fields("TransferKids") = .CellValue(lngRow, "TransferKids")
        Fields("TransferFree") = .CellValue(lngRow, "TransferFree")
        Fields("TransferTotal") = .CellValue(lngRow, "TransferTotal")
        Fields("PickUpPointTime") = .CellValue(lngRow, "PickUpPointTime")
        Fields("PickUpPointExactPoint") = .CellValue(lngRow, "PickUpPointExactPoint")
        Fields("TransferRemarks") = .CellValue(lngRow, "TransferRemarks")
        Fields("TransferDestination") = .CellValue(lngRow, "DestinationDescription")
    
        EOF = False
        blnLastPage = False
    
    End With

End Sub

Private Sub Detail_Format()
    
    LayoutAction = 7
    
End Sub

Private Sub PageFooter_Format()

    If Not blnLastPage Then
       lblContinue.Caption = "г ейтупысг сумевифетаи..."
    Else
        lblContinue.Caption = "текос ейтупысгс"
    End If

End Sub

Private Sub PageHeader_Format()

    rptHeaderL1.text = arrCompanyData(7)
    rptHeaderL2.text = arrCompanyData(8)
    rptHeaderL3.text = arrCompanyData(9)
    rptHeaderL4.text = arrCompanyData(10)
    
    rptTitle.text = "глеяокоцио летажояым апо " & CoachesReport.mskFrom.text & " еыс " & CoachesReport.mskTo.text
    
End Sub


