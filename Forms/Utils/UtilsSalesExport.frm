VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form UtilsSalesExport 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12300
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   5700
      TabIndex        =   16
      Top             =   6000
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "UtilsSalesExport.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "UtilsSalesExport.frx":001C
         BarPictureMode  =   0
         BackPictureMode =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMaster 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Τίτλος"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   18
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   75
      TabIndex        =   11
      Top             =   7725
      Width           =   6090
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Συνέχεια"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   3
         Left            =   4500
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   8421631
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Κλείσιμο"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1650
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Αποθήκευση"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   3075
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Νέα αναζήτηση"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   525
      TabIndex        =   10
      Top             =   4350
      Width           =   4740
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   75
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "UtilsSalesExport.frx":0038
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin VB.Frame frmCriteria 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   525
      TabIndex        =   1
      Top             =   5100
      Width           =   5115
      Begin UserControls.newDate mskFrom 
         Height          =   465
         Left            =   1575
         TabIndex        =   2
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         ForeColor       =   0
         Text            =   "01/01/2017"
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UserControls.newDate mskTo 
         Height          =   465
         Left            =   3150
         TabIndex        =   3
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
         ForeColor       =   0
         Text            =   "01/01/2017"
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   2
         Left            =   4650
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   1
         Left            =   1125
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   0
         Left            =   0
         Top             =   675
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Περίοδος"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   450
         TabIndex        =   8
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   1575
         Width           =   5115
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "Κριτήρια αναζήτησης"
         BeginProperty Font 
            Name            =   "Aka-Acid-Steelfish"
            Size            =   14.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   75
         Width           =   1665
      End
      Begin VB.Label lblToday 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         Caption         =   "01/05/2017"
         BeginProperty Font 
            Name            =   "Aka-Acid-Steelfish"
            Size            =   14.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   2250
         TabIndex        =   4
         Top             =   75
         Width           =   2715
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5115
      End
   End
   Begin iGrid300_10Tec.iGrid grdSalesExport 
      Height          =   5715
      Left            =   450
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1500
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   10081
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483631
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   6000
      Top             =   7200
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   10950
      Top             =   4275
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   3000
      Top             =   8550
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   12
      Left            =   0
      Top             =   2250
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Δημιουργία αρχείου πωλήσεων"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   30
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   720
      Left            =   225
      TabIndex        =   9
      Top             =   75
      Width           =   7320
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   0
      Top             =   0
      Width           =   840
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "UtilsSalesExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnProcessing As Boolean

Private Function AddRecordsToGrid(rstRecordset As Recordset)

    'Variables
    Dim lngRow As Long
    Dim intLine As Integer
    Dim strSmallInvoice As String
    Dim strFullInvoice As String
    Dim strGrossAmount As String
    Dim curNetAmount As Currency
    Dim strVAT As String
    Dim curGrossAmount As Currency
    Dim curTotalGrossAmount As Currency
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 3, 0, 0, 1, 0
    cmdButton(2).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            
            'Στοιχείο - Νο
            strFullInvoice = !CodeShortDescriptionB
            strFullInvoice = strFullInvoice & Right("000000" & !InvoiceNo, 6)
            strSmallInvoice = !CodeShortDescriptionB
            strSmallInvoice = strSmallInvoice & Right("00000" & !InvoiceNo, 5)
            
            If !InvoiceMasterRefersTo = "4" Then
                '38.00 Ταμείο ή 38.03 Τράπεζα
                intLine = intLine + 1
                grdSalesExport.AddRow
                lngRow = grdSalesExport.RowCount
                grdSalesExport.CellValue(lngRow, "Line") = intLine
                grdSalesExport.CellValue(lngRow, "AccountCode") = IIf(IsNull(!BankAccountsCode), strCashAccountsCode, !BankAccountsCode)
                grdSalesExport.CellValue(lngRow, "InvoiceDateIssue") = Format(!InvoiceDateIssue, "dd/mm/yy")
                grdSalesExport.CellValue(lngRow, "ShowInList") = "ΧΡ"
                grdSalesExport.CellValue(lngRow, "CustomerDescription") = !Description
                grdSalesExport.CellValue(lngRow, "CodeShortDescriptionB") = strFullInvoice
                grdSalesExport.CellValue(lngRow, "CodeShortDescriptionBSmaller") = strSmallInvoice
                strGrossAmount = !Amount
                strGrossAmount = (!CodeCustomers + strGrossAmount) * -1
                curGrossAmount = CCur(strGrossAmount)
                curNetAmount = curGrossAmount
                grdSalesExport.CellValue(lngRow, "Debit") = curGrossAmount
                grdSalesExport.CellValue(lngRow, "D/C") = "1"
                grdSalesExport.CellValue(lngRow, "TaxNo") = !TaxNo
                '30 Πελάτης
                intLine = intLine + 1
                grdSalesExport.AddRow
                lngRow = grdSalesExport.RowCount
                grdSalesExport.CellValue(lngRow, "Line") = intLine
                grdSalesExport.CellValue(lngRow, "AccountCode") = !AccountCode
                grdSalesExport.CellValue(lngRow, "InvoiceDateIssue") = Format(!InvoiceDateIssue, "dd/mm/yy")
                grdSalesExport.CellValue(lngRow, "ShowInList") = "ΧΡ"
                grdSalesExport.CellValue(lngRow, "CustomerDescription") = !Description
                grdSalesExport.CellValue(lngRow, "CodeShortDescriptionB") = strFullInvoice
                grdSalesExport.CellValue(lngRow, "CodeShortDescriptionBSmaller") = strSmallInvoice
                strGrossAmount = !Amount
                strGrossAmount = (!CodeCustomers + strGrossAmount) * -1
                curGrossAmount = CCur(strGrossAmount)
                curNetAmount = curGrossAmount
                grdSalesExport.CellValue(lngRow, "Debit") = curGrossAmount
                grdSalesExport.CellValue(lngRow, "D/C") = "0"
                grdSalesExport.CellValue(lngRow, "TaxNo") = !TaxNo
            End If
            
            If !InvoiceMasterRefersTo = "2" Then
                '30 Πελάτης
                intLine = intLine + 1
                grdSalesExport.AddRow
                lngRow = grdSalesExport.RowCount
                grdSalesExport.CellValue(lngRow, "Line") = intLine
                grdSalesExport.CellValue(lngRow, "AccountCode") = !AccountCode
                grdSalesExport.CellValue(lngRow, "InvoiceDateIssue") = Format(!InvoiceDateIssue, "dd/mm/yy")
                grdSalesExport.CellValue(lngRow, "ShowInList") = "ΠΩ"
                grdSalesExport.CellValue(lngRow, "CustomerDescription") = !Description
                grdSalesExport.CellValue(lngRow, "CodeShortDescriptionB") = strFullInvoice
                grdSalesExport.CellValue(lngRow, "CodeShortDescriptionBSmaller") = strSmallInvoice
                strGrossAmount = CCur(!InvoiceOutAdultsAmountWithTransfer + !InvoiceOutKidsAmountWithTransfer + !InvoiceOutAdultsAmountWithoutTransfer + !InvoiceOutKidsAmountWithoutTransfer + !InvoiceOutDirectAmount)
                strGrossAmount = !CodeCustomers + strGrossAmount
                curGrossAmount = CCur(strGrossAmount)
                curNetAmount = curGrossAmount
                grdSalesExport.CellValue(lngRow, "Debit") = Format(curGrossAmount)
                grdSalesExport.CellValue(lngRow, "D/C") = "1"
                grdSalesExport.CellValue(lngRow, "TaxNo") = !TaxNo
                '54 ΦΠΑ
                If !VATStateID = 1 Then
                    intLine = intLine + 1
                    grdSalesExport.AddRow
                    lngRow = grdSalesExport.RowCount
                    grdSalesExport.CellValue(lngRow, "Line") = intLine
                    grdSalesExport.CellValue(lngRow, "AccountCode") = strVATAccountsCode
                    grdSalesExport.CellValue(lngRow, "InvoiceDateIssue") = Format(!InvoiceDateIssue, "dd/mm/yy")
                    grdSalesExport.CellValue(lngRow, "ShowInList") = "ΠΩ"
                    grdSalesExport.CellValue(lngRow, "CustomerDescription") = !Description
                    grdSalesExport.CellValue(lngRow, "CodeShortDescriptionB") = strFullInvoice
                    grdSalesExport.CellValue(lngRow, "CodeShortDescriptionBSmaller") = strSmallInvoice
                    strVAT = "1." & intVAT
                    curNetAmount = curGrossAmount / Val(strVAT)
                    grdSalesExport.CellValue(lngRow, "Debit") = Format(curGrossAmount - curNetAmount, "#,##0.00")
                    grdSalesExport.CellValue(lngRow, "D/C") = "0"
                    grdSalesExport.CellValue(lngRow, "TaxNo") = !TaxNo
                End If
                '73 Πωλήσεις
                intLine = intLine + 1
                grdSalesExport.AddRow
                lngRow = grdSalesExport.RowCount
                grdSalesExport.CellValue(lngRow, "Line") = intLine
                grdSalesExport.CellValue(lngRow, "AccountCode") = IIf(!ShipSalesCode <> "", !ShipSalesCode, strSalesAccountsCode)
                grdSalesExport.CellValue(lngRow, "InvoiceDateIssue") = Format(!InvoiceDateIssue, "dd/mm/yy")
                grdSalesExport.CellValue(lngRow, "ShowInList") = "ΠΩ"
                grdSalesExport.CellValue(lngRow, "CustomerDescription") = !Description
                grdSalesExport.CellValue(lngRow, "CodeShortDescriptionB") = strFullInvoice
                grdSalesExport.CellValue(lngRow, "CodeShortDescriptionBSmaller") = strSmallInvoice
                grdSalesExport.CellValue(lngRow, "Debit") = Format(curNetAmount, "#,##.00")
                grdSalesExport.CellValue(lngRow, "D/C") = "0"
                grdSalesExport.CellValue(lngRow, "TaxNo") = !TaxNo
            End If
            intLine = 0
            'Σύνολα
            curTotalGrossAmount = curTotalGrossAmount + strGrossAmount
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
    End With
    
    'Still here
    If blnProcessing Then
        grdSalesExport.AddRow , , , , , , , 2
        lngRow = grdSalesExport.RowCount
        grdSalesExport.CellValue(lngRow, "Debit") = curTotalGrossAmount
        blnProcessing = False
        AddRecordsToGrid = True
    Else
        AddRecordsToGrid = False
    End If
    
End Function

Private Function CreateFile()

    On Error GoTo ErrTrap
        
    'Local μεταβλητες
    Dim lngRow As Long
    Dim IsGridValid As Boolean
    Dim intLoop As Integer
    Dim strNewDebit As String
    Dim strTemp As String
    Dim strCompany As String
    
    'Αρχικές τιμές
    IsGridValid = True
        
    'Αρχείο λαθών
    Open strReportsPathName & "Errors.txt" For Append As #1
        
    'Ελέγχω το πλέγμα για λάθη
    With grdSalesExport
        For lngRow = 1 To .RowCount
            If .CellValue(lngRow, "Line") <> "" And (.CellValue(lngRow, "AccountCode") = "" Or IsNull(.CellValue(lngRow, "AccountCode"))) Then
                Print #1, .CellValue(lngRow, "CompanyDescription"); " " & .CellValue(lngRow, "CodeShortDescriptionBSmaller") & " - Δεν υπάρχει κωδικός πελάτη ή/και Α.Φ.Μ."
                IsGridValid = False
            End If
        Next lngRow
    End With
    
    'Αρχείο λαθών
    Close #1
    
    'Ελεγχος
    If Not IsGridValid Then
        If MyMsgBox(4, strAppTitle, strAppMessages(6), 1) Then
        End If
        CreateFile = False
        Exit Function
    End If
    
    'Δημιουργία αρχείου
    Open strReportsPathName & strAccountsFileName For Output As #1
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, grdSalesExport.RowCount
    
    'Πλέγμα
    With grdSalesExport
        For lngRow = 1 To .RowCount
            UpdateProgressBar Me
            If .CellValue(lngRow, "Line") <> "" Then
                strNewDebit = ""
                'Εγγραφές άρθρου
                strTemp = .CellValue(lngRow, "Line")
                strTemp = Right("  " & .CellValue(lngRow, "Line"), 2)
                'Πελάτης
                strCompany = Trim(.CellValue(lngRow, "CustomerDescription"))
                If Len(strCompany) < 20 Then
                    strCompany = strCompany & Space(20 - Len(strCompany))
                Else
                    strCompany = Left(strCompany, 20)
                End If
                'Γραμμή
                For intLoop = 1 To Len(.CellText(lngRow, "Debit"))
                    If Mid(.CellText(lngRow, "Debit"), intLoop, 1) = "," Then
                        strNewDebit = strNewDebit + "."
                    Else
                        If Mid(.CellText(lngRow, "Debit"), intLoop, 1) <> "." Then
                            strNewDebit = strNewDebit + Mid(.CellText(lngRow, "Debit"), intLoop, 1)
                        End If
                    End If
                Next intLoop
                Print #1, strTemp & " " _
                    & .CellText(lngRow, "AccountCode") _
                    & Space(9) _
                    & .CellText(lngRow, "InvoiceDateIssue") _
                    & Space(1) _
                    & .CellText(lngRow, "ShowInList") _
                    & Space(1) _
                    & strCompany _
                    & Space(21) _
                    & .CellText(lngRow, "CodeShortDescriptionBSmaller") _
                    & Space(21 - Len(strNewDebit)) & strNewDebit _
                    & Space(1) _
                    & .CellText(lngRow, "D/C") _
                    & Space(1) _
                    & .CellText(lngRow, "TaxNo")
            End If
        Next lngRow
    End With
    
    'Τελικές ενέργειες
    Close #1
    frmProgress.Visible = False
    CreateFile = True

    Exit Function
    
ErrTrap:
    grdSalesExport.Redraw = True
    CreateFile = False
    frmProgress.Visible = False
    DisplayErrorMessage True, Err.Description
    Close #1

End Function

Private Function ExportGridToFile()

    If CreateFile Then MyMsgBox 1, strAppTitle, strStandardMessages(8), 1

End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            DoJobs
        Case 1
            ExportGridToFile
        Case 2
            AbortProcedure False
        Case 3
            AbortProcedure True
    End Select
    
End Sub

Private Function ValidateFields()

    ValidateFields = False
    
    'Από
    If Not IsDate(mskFrom.text) Then MyMsgBox 4, strAppTitle, strStandardMessages(1), 1: mskFrom.SetFocus: Exit Function
    
    'Εως
    If Not IsDate(mskTo.text) Then MyMsgBox 4, strAppTitle, strStandardMessages(1), 1: mskTo.SetFocus: Exit Function
    
    'Σωστό διάστημα
    If CDate(mskFrom.text) > CDate(mskTo.text) Then MyMsgBox 4, strAppTitle, strStandardMessages(10), 1: mskFrom.SetFocus: Exit Function
    
    ValidateFields = True
    
End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grdSalesExport
        frmCriteria.Visible = True
        mskFrom.SetFocus
        UpdateButtons Me, 3, 1, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function GetRecords()

    'SQL
    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String

    'Local variables
    Dim intLine As Integer
    Dim strFullInvoice As String
    Dim strSmallInvoice As String
    Dim strGrossAmount As String
    Dim curGrossAmount As Currency
    Dim curNetAmount As Currency
    Dim strVAT As String
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset

    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    frmCriteria.Visible = False

    'Πλέγμα
    With grdSalesExport
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT " _
        & "InvoiceMasterRefersTo, InvoiceDateIssue, InvoiceNo, " _
        & "InvoiceOutAdultsAmountWithTransfer, InvoiceOutKidsAmountWithTransfer, InvoiceOutAdultsAmountWithoutTransfer, InvoiceOutKidsAmountWithoutTransfer, InvoiceOutDirectAmount, " _
        & "Amount, " _
        & "BankAccountsCode, " _
        & "CodeShortDescriptionB, CodeCustomers, " _
        & "Description, VATStateID, TaxNo, AccountCode, " _
        & "ShipSalesCode " _
        & "FROM (((((((Invoices  " _
        & "LEFT JOIN InvoicesOut ON InvoicesOut.InvoiceOutTrnID = Invoices.InvoiceTrnID) " _
        & "LEFT JOIN PaymentIn ON PaymentIn.TrnID = Invoices.InvoiceTrnID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN Customers ON Invoices.InvoicePersonID = Customers.ID) " _
        & "LEFT JOIN Ships ON InvoicesOut.InvoiceOutShipID = Ships.ShipID) " _
        & "LEFT JOIN Banks ON PaymentIn.BankID = Banks.BankID) " _
        & "LEFT JOIN PaymentWays ON PaymentIn.PaymentWayID = PaymentWays.PaymentWayID) "

    'Από
    If IsDate(mskFrom.text) Then
        strThisParameter = "datFrom Date"
        strThisQuery = "Invoices.InvoiceDateIssue >= datFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskFrom.text)
    End If

    'Εως
    If IsDate(mskTo.text) Then
        strThisParameter = "datTo Date"
        strThisQuery = "Invoices.InvoiceDateIssue <= datTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskTo.text)
    End If
    
    'Πελάτες κανονικοί
    strThisParameter = "strNotNormalCustomers String"
    strThisQuery = "AccountCode <> strNotNormalCustomers"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = "99.99.99.999"
    
    'Κινήσεις πωλήσεων
    strThisParameter = "strInvoiceMasterRefersToShips String"
    strThisQuery = "(Invoices.InvoiceMasterRefersTo = strInvoiceMasterRefersToShips"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = "2"
    
    'Κινήσεις πελατών
    strThisParameter = "strInvoiceMasterRefersToCoaches String"
    strThisQuery = "Invoices.InvoiceMasterRefersTo = strInvoiceMasterRefersToCoaches)"
    strLogic = " OR "
    GoSub UpdateSQLString
    arrQuery(intIndex) = "4"
    
    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceDateIssue, CodeShortDescriptionB, InvoiceNo"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Επιστροφή
    Set GetRecords = rstRecordset
    
    Exit Function
    
UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return
    
ErrTrap:
    ClearFields grdSalesExport, frmProgress
    DisplayErrorMessage True, Err.Description
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)

    Dim CtrlDown
    
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyS And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyEscape
            If cmdButton(2).Enabled Then cmdButton_Click 2: Exit Function
            If cmdButton(3).Enabled Then cmdButton_Click 3
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()
    
    AddColumns
    UpdateColors Me, False, grdSalesExport
    SetUpGrid lstIconList, grdSalesExport
    ClearFields mskFrom, mskTo
    EnableFields mskFrom, mskTo
    UpdateButtons Me, 3, 1, 0, 0, 1
    frmCriteria.Visible = True
    
End Sub

Private Sub grdSalesExport_ColHeaderMouseEnter(ByVal lCol As Long)

    grdSalesExport.Header.Buttons = True

End Sub

Private Function DoJobs()

    'On Error GoTo ErrTrap
    
    Dim rstRecordset As Recordset
    
    If ValidateFields Then
        Set rstRecordset = GetRecords
        If rstRecordset.RecordCount > 0 Then
            If AddRecordsToGrid(rstRecordset) Then
                EnableGrid grdSalesExport, False
                HighlightRow grdSalesExport, 1, 1, "", True
                frmProgress.Visible = False
                cmdButton(2).Caption = "Νέα αναζήτηση"
                UpdateButtons Me, 3, 0, 1, 1, 0
            Else
                MyMsgBox 4, strAppTitle, strStandardMessages(27), 1
                frmProgress.Visible = False
                frmCriteria.Visible = True
                cmdButton(2).Caption = "Νέα αναζήτηση"
                UpdateButtons Me, 3, 1, 0, 0, 1
                mskFrom.SetFocus
            End If
        Else
            UpdateButtons Me, 3, 1, 0, 0, 1
            MyMsgBox 1, strAppTitle, strStandardMessages(7), 1
            frmCriteria.Visible = True
            UpdateButtons Me, 3, 1, 0, 0, 1
            mskFrom.SetFocus
        End If
    End If
    
    Exit Function
    
ErrTrap:
    ClearFields grdSalesExport, frmProgress
    DisplayErrorMessage True, Err.Description
    frmCriteria.Visible = True
    UpdateButtons Me, 3, 1, 0, 0, 1
    mskFrom.SetFocus

End Function

Private Sub grdSalesExport_ColHeaderMouseLeave(ByVal lCol As Long)

    grdSalesExport.Header.Buttons = False
    
End Sub

Private Sub grdSalesExport_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdSalesExport", grdSalesExport.LayoutCol

End Sub

Private Function AddColumns()

    AddColumnsToGrid grdSalesExport, 44, GetSetting(strAppTitle, "Layout Strings", "grdSalesExport"), "04NCILine,15NCNXAccountCode,10NCDInvoiceDateIssue,15NCNShowInList,15NLNCustomerDescription,05NLNCodeShortDescriptionB,05NLNCodeShortDescriptionBSmaller,10NRFDebit,10NCND/C,15NLNTaxNo", "Γ,Κωδ. Γεν. Λογιστικής,Ημερομηνία,Τ,Επωνυμία,Παραστατικό,Παραστατικό,Αξία,Χ/Π,Α.Φ.Μ."
    
End Function

Private Function AddDummyLinesToGrid()

    AddDummyLines grdSalesExport, "99", "123456789012345", "A99/99/99A", "AA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAA", "AAAAAAAAAA", "99999.99", "A"
    
End Function
