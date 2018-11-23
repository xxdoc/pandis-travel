VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptShipsRouteReport 
   Caption         =   "Λίστα Επιβαινόντων"
   ClientHeight    =   12450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14985
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26432
   _ExtentY        =   21960
   SectionData     =   "rptShipsRouteReport.dsx":0000
End
Attribute VB_Name = "rptShipsRouteReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRow As Long
Dim blnError As Boolean
Private blnLastPage As Boolean

Private Sub ActiveReport_DataInitialize()

    lngRow = 0
    blnError = False
    
    Fields.RemoveAll
    
    Fields.Add "AA"
    Fields.Add "LastName"
    Fields.Add "FirstName"
    Fields.Add "GenderDescription"
    Fields.Add "AgeDescription"
    Fields.Add "OccupantDescription"
    Fields.Add "Care"
    Fields.Add "Remarks"
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If blnError Then Exit Sub
    
    lngRow = lngRow + 1
    
    With ShipsRouteReport
        
        If lngRow > .grdShipsRouteReport.RowCount Then
            blnLastPage = True
            EOF = True
            Exit Sub
        End If
        
        Fields("AA") = .grdShipsRouteReport.CellValue(lngRow, "AA")
        Fields("LastName") = .grdShipsRouteReport.CellValue(lngRow, "LastName")
        Fields("FirstName") = .grdShipsRouteReport.CellValue(lngRow, "FirstName")
        Fields("GenderDescription") = .grdShipsRouteReport.CellText(lngRow, "Gender")
        Fields("AgeDescription") = .grdShipsRouteReport.CellText(lngRow, "Age")
        Fields("OccupantDescription") = .grdShipsRouteReport.CellText(lngRow, "OccupantDescription")
        Fields("Care") = .grdShipsRouteReport.CellValue(lngRow, "Care")
        Fields("Remarks") = .grdShipsRouteReport.CellValue(lngRow, "Remarks")
        
        EOF = False
        blnLastPage = False
        
    End With
        
End Sub

Private Sub ReportHeader_Format()

    On Error GoTo ErrTrap
    
    With ShipsRouteReport
        
        'Φορολογικά στοιχεία εταιρίας
        lblCompanyData.Caption = arrCompanyData(1) & Chr(13) & arrCompanyData(2) & Chr(13) & arrCompanyData(3) & Chr(13) & arrCompanyData(4) & Chr(13) & arrCompanyData(5) & Chr(13) & arrCompanyData(6)
        
        'Ημερομηνία
        lblDate.Caption = "ΔΡΟΜΟΛΟΓΙΟ ΤΗΣ " + .mskDate.text
        
        'Βρίσκω τα στοιχεία του πλοίου και των υπεύθυνων καταγραφής επιβατών
        Dim strSQL As String
        Dim rstRecordset As Recordset
        
        strSQL = "SELECT ShipDescription, ShipFlag, ShipRegistryNo, ShipIMO, ShipManagerName, ShipManagerNameInGreece, ShipManagerAgent, " _
            & "ShipManagerAPersonName, ShipManagerAPersonPhones, ShipManagerAPersonEmail, ShipManagerAPersonFax, ShipManagerAPersonAddress, " _
            & "ShipManagerBPersonName, ShipManagerBPersonPhones, ShipManagerBPersonEmail, ShipManagerBPersonFax, ShipManagerBPersonAddress " _
        & "FROM Ships " _
        & "INNER JOIN ShipsManagers ON Ships.ShipID = ShipsManagers.ShipManagerShipID " _
        & "WHERE Ships.ShipID = " & ShipsRouteReport.txtShipID.text
        
        Set TempQuery = CommonDB.CreateQueryDef(""): TempQuery.SQL = strSQL
        Set rstRecordset = TempQuery.OpenRecordset()
        
        'Πλοίο
        lblShipDescription.Caption = rstRecordset!ShipDescription
        lblShipFlag.Caption = rstRecordset!ShipFlag
        lblShipRegistryNo.Caption = rstRecordset!ShipRegistryNo
        lblShipIMO.Caption = rstRecordset!ShipIMO
        
        'Δρομολόγιο
        lblRouteFrom.Caption = ShipsRouteReport.txtFrom.text
        lblRouteVia.Caption = ShipsRouteReport.txtVia.text
        lblRouteTo.Caption = ShipsRouteReport.txtTo.text
        
        'Εταιρία
        lblManager.Caption = rstRecordset!ShipManagerName
        lblManagerInGreece.Caption = rstRecordset!ShipManagerNameInGreece
        lblAgent.Caption = rstRecordset!ShipManagerAgent
        
        'Υπεύθυνος καταγραφής
        lblAPersonName.Caption = rstRecordset!ShipManagerAPersonName
        lblAPersonPhones.Caption = rstRecordset!ShipManagerAPersonPhones
        lblAPersonEmail.Caption = rstRecordset!ShipManagerAPersonEmail
        lblAPersonFax.Caption = rstRecordset!ShipManagerAPersonFax
        lblAPersonAddress.Caption = rstRecordset!ShipManagerAPersonAddress
        
        'Αντικαταστάτης υπεύθυνου καταγραφής
        lblBPersonName.Caption = rstRecordset!ShipManagerBPersonName
        lblBPersonPhones.Caption = rstRecordset!ShipManagerBPersonPhones
        lblBPersonEmail.Caption = rstRecordset!ShipManagerBPersonEmail
        lblBPersonFax.Caption = rstRecordset!ShipManagerBPersonFax
        lblBPersonAddress.Caption = rstRecordset!ShipManagerBPersonAddress
        
        'Υπεύθυνο άτομο για τη διαβίβαση
        lblPerson.Caption = rstRecordset!ShipManagerName
        
        'Ημερομηνίες
        lblPlace.Caption = "ΚΕΡΚΥΡΑ"
        lblFooterDate.Caption = Format(Date, "dd/mm/yyyy")
        lblTime.Caption = Format(Time, "hh:mm")
        
    End With
    
    Exit Sub
    
ErrTrap:
    blnError = True
    DisplayErrorMessage True, Err.Description

End Sub


