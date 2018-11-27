VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Begin VB.Form CommonMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6210
   ClientLeft      =   2025
   ClientTop       =   120
   ClientWidth     =   7545
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "CommonMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7545
   WindowState     =   2  'Maximized
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl vbExplorerBar 
      Height          =   481
      Left            =   624
      TabIndex        =   1
      Top             =   2106
      Width           =   481
      _ExtentX        =   847
      _ExtentY        =   847
      BackColorEnd    =   0
      BackColorStart  =   0
   End
   Begin vbalIml6.vbalImageList imgImageList 
      Left            =   9975
      Top             =   450
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   2296
      Images          =   "CommonMain.frx":0ECA
      Version         =   131072
      KeyCount        =   2
      Keys            =   "ÿ"
   End
   Begin vbalIml6.vbalImageList ilsExplorerIcons 
      Left            =   600
      Top             =   2625
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   48
      IconSizeY       =   48
      ColourDepth     =   24
      Size            =   57960
      Images          =   "CommonMain.frx":17E2
      Version         =   131072
      KeyCount        =   6
      Keys            =   "ÿÿÿÿÿ"
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   525
      Width           =   3690
   End
   Begin VB.Image imgImage 
      Appearance      =   0  'Flat
      Height          =   2400
      Left            =   4275
      Picture         =   "CommonMain.frx":FA6A
      Top             =   3000
      Width           =   1995
   End
End
Attribute VB_Name = "CommonMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function AddCreditorsSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = vbExplorerBar.Bars.Add(, "Πιστωτές", Space(5) & "Πιστωτές")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "ΠιστωτέςΑρχείο", " - Αρχείο")
        Set cItem = cBar.Items.Add(, "ΠιστωτέςΚινήσεις", " - Κινήσεις")
        Set cItem = cBar.Items.Add(, "ΠιστωτέςΚαρτέλα", " - Καρτέλα")
        Set cItem = cBar.Items.Add(, "ΠιστωτέςΙσοζύγιο", " - Ισοζύγιο")
        Set cItem = cBar.Items.Add(, "ΠιστωτέςΤύποιΠαραστατικών", " - Τύποι παραστατικών")

End Function

Private Function AddDebitorsSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = vbExplorerBar.Bars.Add(, "Χρεώστες", Space(5) & "Χρεώστες")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "ΧρεώστεςΑρχείο", " - Αρχείο")
        Set cItem = cBar.Items.Add(, "ΧρεώστεςΚινήσειςΕκδρομέςΠλοίων", " - Κινήσεις εκδρομών πλοίων")
        Set cItem = cBar.Items.Add(, "ΧρεώστεςΚινήσειςΕκδρομέςΛεωφορείων", " - Κινήσεις εκδρομών λεωφορείων")
        Set cItem = cBar.Items.Add(, "ΧρεώστεςΚαρτέλα", " - Καρτέλα")
        Set cItem = cBar.Items.Add(, "ΧρεώστεςΙσοζύγιο", " - Ισοζύγιο")
        Set cItem = cBar.Items.Add(, "ΧρεώστεςΤύποιΠαραστατικώνΕκδρομώνΠλοίων", " - Τύποι παραστατικών εκδρομών πλοίων")
        Set cItem = cBar.Items.Add(, "ΧρεώστεςΤύποιΠαραστατικώνΕκδρομώνΛεωφορείων", " - Τύποι παραστατικών εκδρομών λεωφορείων")

End Function

Private Function AddExitSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = vbExplorerBar.Bars.Add(, "Εξοδος", Space(5) & "Εξοδος")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "ΕξοδοςΑλλαγήΕταιρίας", "- Αλλαγή εταιρίας")
        Set cItem = cBar.Items.Add(, "ΕξοδοςΤερματισμόςΕφαρμογής", "- Τερματισμός εφαρμογής")

End Function

Private Function AddIncomeMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = vbExplorerBar.Bars.Add(, "Εσοδα", Space(5) & "Εσοδα")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "ΕσοδαΕκδρομέςΠλοίων", "Εκδρομές πλοίων")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "ΕσοδαΕκδρομέςΠλοίωνΚινήσεις", " - Κινήσεις")
            Set cItem = cBar.Items.Add(, "ΕσοδαΕκδρομέςΠλοίωνΗμερολόγιο", " - Ημερολόγιο")
        Set cItem = cBar.Items.Add(, "ΕσοδαΕκδρομέςΛεωφορείων", "Εκδρομές λεωφορείων")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "ΕσοδαΕκδρομέςΛεωφορείωνΚινήσεις", " - Κινήσεις")
            Set cItem = cBar.Items.Add(, "ΕσοδαΕκδρομέςΛεωφορείωνΗμερολόγιο", " - Ημερολόγιο")
        Set cItem = cBar.Items.Add(, "ΕσοδαΤύποιΠαραστατικών", "Τύποι παραστατικών")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "ΕσοδαΤύποιΠαραστατικώνΔιαχείρηση", " - Διαχείρηση")

End Function

Private Function AddShipsPassengers()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
        
    Set cBar = vbExplorerBar.Bars.Add(, "ΕπιβαίνοντεςΠλοίων", Space(5) & "Επιβαίνοντες πλοίων")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "ΕπιβαίνοντεςΠλοίωνΚινήσεις", " - Αρχείο")
        Set cItem = cBar.Items.Add(, "ΕπιβαίνοντεςΠλοίωνΚατάσταση", " - Κατάσταση επιβατών")
        Set cItem = cBar.Items.Add(, "ΕπιβαίνοντεςΠλοίωνΣτατιστικάΠωλήσεις", " - Στατιστικά με βάση τις πωλήσεις")
        Set cItem = cBar.Items.Add(, "ΕπιβαίνοντεςΠλοίωνΣτατιστικάΛιμεναρχείο", " - Στατιστικά με βάση τους επιβαίνοντες")

End Function

Private Function AddTransfersSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
        
    Set cBar = vbExplorerBar.Bars.Add(, "ΜεταφορέςΛεωφορείων", Space(5) & "Επιβαίνοντες λεωφορείων")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "ΜεταφορέςΛεωφορείωνΓρήγορηΚαταχώρηση", " - Γρήγορη καταχώρηση")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "ΜεταφορέςΛεωφορείωνΓρήγορηΚαταχώρησηΚινήσεις", Space(5) & " - Κινήσεις")
            Set cItem = cBar.Items.Add(, "ΜεταφορέςΛεωφορείωνΓρήγορηΚαταχώρησηΗμερολόγιο", Space(5) & " - Κατάσταση επιβατών")
        Set cItem = cBar.Items.Add(, "ΜεταφορέςΛεωφορείωνΚανονικήΚαταχώρηση", " - Κανονική καταχώρηση")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "ΜεταφορέςΛεωφορείωνΚανονικήΚαταχώρησηΚινήσεις", Space(5) & " - Κινήσεις")
            Set cItem = cBar.Items.Add(, "ΜεταφορέςΛεωφορείωνΚανονικήΚαταχώρησηΗμερολόγιο", Space(5) & " - Κατάσταση επιβατών")

End Function

Private Function AddUtilsSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
        
    Set cBar = vbExplorerBar.Bars.Add(, "Βοηθητικά", Space(5) & "Βοηθητικά")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "Παραμετροποίηση", "Παραμετροποίηση")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "ΒοηθητικάΠαράμετροι", Space(5) & " - Γενικές παράμετροι")
            Set cItem = cBar.Items.Add(, "TablesPrinters", Space(5) & " - Εκτυπωτές")
        Set cItem = cBar.Items.Add(, "Πίνακες", "Πίνακες")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "TablesShipRoutes", Space(5) & " - Δρομολόγια πλοίων")
            Set cItem = cBar.Items.Add(, "TablesCoachRoutes", Space(5) & " - Δρομολόγια λεωφορείων")
            Set cItem = cBar.Items.Add(, "TablesExpenseCategories", Space(5) & " - Κατηγορίες εξόδων")
            Set cItem = cBar.Items.Add(, "TablesTaxOffices", Space(5) & " - Οικονομικές υπηρεσίες")
            Set cItem = cBar.Items.Add(, "TablesPaymentTerms", Space(5) & " - Όροι πληρωμής")
            Set cItem = cBar.Items.Add(, "TablesShips", Space(5) & " - Πλοία")
            Set cItem = cBar.Items.Add(, "TablesVATPercents", Space(5) & " - Ποσοστά Φ.Π.Α.")
            Set cItem = cBar.Items.Add(, "TablesDestinationsΠλοίων", Space(5) & " - Προορισμοί πλοίων")
            Set cItem = cBar.Items.Add(, "TablesDestinationsΛεωφορείων", Space(5) & " - Προορισμοί λεωφορείων")
            Set cItem = cBar.Items.Add(, "TablesPickupPoints", Space(5) & " - Σημεία παραλαβής επιβατών")
            Set cItem = cBar.Items.Add(, "UtilsJoinDestinationsWithRoutes", Space(5) & " - Σύνδεση προορισμών με δρομολόγια λεωφορείων")
            Set cItem = cBar.Items.Add(, "TablesPriceListsΠλοίων", Space(5) & " - Τιμοκατάλογοι εκδρομών πλοίων")
            Set cItem = cBar.Items.Add(, "TablesPriceListsΛεωφορείων", Space(5) & " - Τιμοκατάλογοι εκδρομών λεωφορείων")
            Set cItem = cBar.Items.Add(, "TablesBanks", Space(5) & " - Τράπεζες")
            Set cItem = cBar.Items.Add(, "TablesPaymentWays", Space(5) & " - Τρόποι πληρωμής")
            Set cItem = cBar.Items.Add(, "TablesOccupantsDescriptions", Space(5) & " - Χαρακτηρισμοί επιβαινόντων")
            Set cItem = cBar.Items.Add(, "TablesUsers", Space(5) & " - Χρήστες")
        Set cItem = cBar.Items.Add(, "Εργασίες", "Εργασίες")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "UtilsSalesExport", Space(5) & " - Δημιουργία αρχείου πωλήσεων")
            Set cItem = cBar.Items.Add(, "UtilsCheckFiles", Space(5) & " - Ελεγχος αρχείων")

End Function

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
End Function



Private Function UpdateCompanyLabel()
            
    With lblCompany
        .BackColor = CommonMain.BackColor
        .Top = vbExplorerBar.Top - 300
        .Left = vbExplorerBar.Left
        .Width = vbExplorerBar.Width + 10
    End With

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Sub Form_Load()

    BuildMenu
    
    strImageDirectory = GetSetting(strApplicationName, "Path Names", "Image Directory")

    With CommonMain
        .ScaleHeight = .Height
        .ScaleWidth = .Width
        .imgImage.Top = .Height - .imgImage.Height - 1000
        .imgImage.Left = Screen.Width - .imgImage.Width - 500
        .BackColor = vbBlack
        .Refresh
    End With
    
    UpdateCompanyLabel
    
    strReportsPathName = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Reports Path Name")
    strUnicodeFile = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Reports Path Name") & "UnicodeFile.txt"
    strAsciiFile = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Reports Path Name") & "AsciiFile.txt"
    
    blnAppIsRunning = True
    
End Sub

Private Function BuildMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
    
    Dim intLoop As Integer
    Dim intItem As Integer
    Dim strMenuCategory As String
    Dim strMenuCategories As String
    
    With CommonMain
        .Tag = "True"
        .Height = Screen.Height
        .ScaleHeight = .Height
    End With
    
    strMenuCategories = GetSetting(strApplicationName, "Settings", "Menu Categories")
    For intLoop = 1 To Len(strMenuCategories)
        While Mid(strMenuCategories, intLoop, 1) <> "," And intLoop <= Len(strMenuCategories)
            strMenuCategory = strMenuCategory & Mid(strMenuCategories, intLoop, 1)
            intLoop = intLoop + 1
        Wend
        intItem = intItem + 1
        ReDim Preserve arrMenu(intItem)
        arrMenu(intItem) = Int(strMenuCategory)
        strMenuCategory = ""
    Next intLoop
    
    With CommonMain.vbExplorerBar
        
        .Height = GetSetting(strApplicationName, "Settings", "Menu Height")
        .Left = ((Screen.Width / Screen.TwipsPerPixelX) / 3)
        .Redraw = False
        .Top = (CommonMain.Height / 2) - (.Height / 2) - 200
        .UseExplorerStyle = False
        .Width = GetSetting(strApplicationName, "Settings", "Menu Width")
        
        AddIncomeMenu
        AddExpensesMenu
        AddDebitorsSubMenu
        AddCreditorsSubMenu
        AddShipsPassengers
        AddTransfersSubMenu
        AddUtilsSubMenu
        AddExitSubMenu
        
        .Redraw = True
        
    End With

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim obj As Object
    
    'Επιλογή κλεισίματος απο το μενού συστήματος, κλικ στο Χ ή ALT-F4
    If UnloadMode = 0 Then
        If CloseApp Then
            For Each obj In Forms
                Unload obj
            Next
            KillProcess strApplicationEXEName: End
        Else
            Cancel = 1
            Exit Sub
        End If
    End If
    
    'Επιλογή κλεισίματος από την επιλογή Εξοδος > Τερματισμός
    If UnloadMode = 1 Then
        KillProcess strApplicationEXEName
    End If

End Sub

Private Function CloseApp()

    CloseApp = False
    
    If MyMsgBox(2, strApplicationName, strStandardMessages(16), 2) Then
        CloseApp = True
    End If

End Function

Private Sub vbExplorerBar_BarClick(bar As vbalExplorerBarLib6.cExplorerBar)

    ResizeBar bar.index, bar.State, vbExplorerBar, arrMenu(1), arrMenu(2), arrMenu(3), arrMenu(4), arrMenu(5), arrMenu(6), arrMenu(7), arrMenu(8)
    
    UpdateCompanyLabel

End Sub

Private Sub ResizeBar(intKey, blnState As Boolean, ExplorerBar As vbalExplorerBarCtl, ParamArray Buttons() As Variant)

    'Κάθετο κεντράρισμα
    With ExplorerBar
        .Height = GetSetting(strApplicationName, "Settings", "Menu Height")
        If Not blnState Then .Top = (Me.Height / 2) - (.Height / 2): Exit Sub
        .Redraw = False
        .Height = Buttons(intKey - 1)
        .Top = (Me.Height / 2) - (.Height / 2) - 50
        .Redraw = True
    End With
        
End Sub

Private Sub vbExplorerBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)

    Dim obj As Object

    Select Case itm.Key
        'Εσοδα
        Case "ΕσοδαΕκδρομέςΠλοίωνΚινήσεις"
            With InvoicesOut
                .lblTitle.Caption = "Εκδρομές πλοίων"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕσοδαΕκδρομέςΠλοίωνΗμερολόγιο"
            With InvoicesOutIndex
                .lblTitle.Caption = "Ημερολόγιο εκδρομών πλοίων"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕσοδαΕκδρομέςΛεωφορείωνΚινήσεις"
            With InvoicesOut
                .lblTitle.Caption = "Εκδρομές λεωφορείων"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .lblLabel(5).Visible = False
                .txtShipDescription.Visible = False
                .cmdIndex(3).Visible = False
                .cmdIndex(8).Visible = False
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕσοδαΕκδρομέςΛεωφορείωνΗμερολόγιο"
            With InvoicesOutIndex
                .lblTitle.Caption = "Ημερολόγιο εκδρομών λεωφορείων"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕσοδαΤύποιΠαραστατικώνΔιαχείρηση"
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών εσόδων"
                .txtCodeMasterRefersTo.text = "2"
                .txtCodeSecondaryRefersTo.text = "0"
                .Tag = "True"
                .Show 1, Me
            End With
        'Τέλος έσοδα
        
        'Εξοδα
        Case "ΕξοδαΚινήσεις"
            With InvoicesIn
                .Tag = "True"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtInvoiceSecondaryRefersTo.text = ""
                .Show 1, Me
            End With
        Case "ΕξοδαΗμερολόγιο"
            With InvoicesInIndex
                .txtInvoiceMasterRefersTo.text = "1"
                .txtInvoiceSecondaryRefersTo.text = ""
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕξοδαΤύποιΠαραστατικών"
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών εξόδων"
                .txtCodeMasterRefersTo.text = "1"
                .txtCodeSecondaryRefersTo.text = "0"
                .Tag = "True"
                .Show 1, Me
            End With
        'Τέλος έξοδα
        
        'Χρεώστες
        Case "ΧρεώστεςΑρχείο"
            With persons
                .Tag = "True"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtInvoiceMasterRefersTo.text = "2"
                .lblTitle.Caption = "Πελάτες"
                .Show 1, Me
            End With
        Case "ΧρεώστεςΚινήσειςΕκδρομέςΠλοίων"
            With PersonsTransactions
                .lblTitle.Caption = "Κινήσεις εκδρομών πλοίων"
                .txtInvoiceMasterRefersTo.text = "4"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΧρεώστεςΚινήσειςΕκδρομέςΛεωφορείων"
            With PersonsTransactions
                .lblTitle.Caption = "Κινήσεις εκδρομών λεωφορείων"
                .txtInvoiceMasterRefersTo.text = "4"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΧρεώστεςΚαρτέλα"
            With PersonsLedger
                .lblTitle.Caption = "Καρτέλα χρεώστη"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΧρεώστεςΙσοζύγιο"
            With PersonsBalanceSheet
                .lblTitle.Caption = "Ισοζύγιο χρεωστών"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΧρεώστεςΤύποιΠαραστατικώνΕκδρομώνΠλοίων"
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών εκδρομών πλοίων"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "4"
                .txtCodeSecondaryRefersTo.text = "1"
                .Show 1, Me
            End With
        Case "ΧρεώστεςΤύποιΠαραστατικώνΕκδρομώνΛεωφορείων"
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών εκδρομών λεωφορείων"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "4"
                .txtCodeSecondaryRefersTo.text = "2"
                .Show 1, Me
            End With
        'Τέλος χρεώστες
        
        'ΠΙστωτές
        Case "ΠιστωτέςΑρχείο"
            With persons
                .Tag = "True"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtInvoiceMasterRefersTo.text = "1"
                .lblTitle.Caption = "Πιστωτές"
                .Show 1, Me
            End With
        Case "ΠιστωτέςΚινήσεις"
            With PersonsTransactions
                .lblTitle.Caption = "Κινήσεις πιστωτών"
                .txtInvoiceMasterRefersTo.text = "3"
                .txtInvoiceSecondaryRefersTo.text = ""
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠιστωτέςΚαρτέλα"
            With PersonsLedger
                .lblTitle.Caption = "Καρτέλα πιστωτή"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠιστωτέςΙσοζύγιο"
            With PersonsBalanceSheet
                .lblTitle.Caption = "Ισοζύγιο πιστωτών"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΠιστωτέςΤύποιΠαραστατικών"
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών πιστωτών"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "3"
                .Show 1, Me
            End With
        'Τέλος πιστωτές
            
        'Επιβαίνοντες πλοίων
        Case "ΕπιβαίνοντεςΠλοίωνΚινήσεις"
            With ShipsTransactions
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕπιβαίνοντεςΠλοίωνΚατάσταση"
            With ShipsRouteReport
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕπιβαίνοντεςΠλοίωνΣτατιστικάΠωλήσεις"
            With ShipsStatistics
                .lblTitle.Caption = "Στατιστικά με βάση τις πωλήσεις"
                .txtTable.text = "Sales"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΕπιβαίνοντεςΠλοίωνΣτατιστικάΛιμεναρχείο"
            With ShipsStatistics
                .lblTitle.Caption = "Στατιστικά με βάση τους επιβαίνοντες"
                .txtTable.text = "Manifest"
                .Tag = "True"
                .Show 1, Me
            End With
        'Τέλος επιβαίνοντες πλοίων
        
        'Μεταφορές λεωφορείων
        Case "ΜεταφορέςΛεωφορείωνΓρήγορηΚαταχώρησηΚινήσεις"
            With CoachesPickupsBrief
                .txtRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΜεταφορέςΛεωφορείωνΓρήγορηΚαταχώρησηΗμερολόγιο"
            With CoachesReport
                .Tag = "True"
                .txtRefersTo.text = "2"
                .txtCallingForm = "MainMenu"
                .grdCoachesReport.Tag = "grdCoachesReportBrief"
                .Show 1, Me
            End With
        Case "ΜεταφορέςΛεωφορείωνΚανονικήΚαταχώρησηΚινήσεις"
            With CoachesPickupsStandard
                .txtRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΜεταφορέςΛεωφορείωνΚανονικήΚαταχώρησηΗμερολόγιο"
            With CoachesReport
                .Tag = "True"
                .txtRefersTo.text = "1"
                .txtCallingForm = "MainMenu"
                .grdCoachesReport.Tag = "grdCoachesReportStandard"
                .Show 1, Me
            End With
        'Τέλος μεταφορές λεωφορείων
        
        'Βοηθητικά
        Case "TablesCoachRoutes"
            With TablesCoachRoutes
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesShipRoutes"
            With TablesShipRoutes
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsSalesExport"
            With UtilsSalesExport
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesPrinters"
            With TablesPrinters
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesTaxOffices"
            With TablesTaxOffices
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΒοηθητικάΠαράμετροι"
            With TablesSettings
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesShips"
            With TablesShips
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesDestinationsΠλοίων"
            With TablesDestinations
                .Tag = "True"
                .lblTitle.Caption = "Προορισμοί πλοίων"
                .txtShowInList.text = "1"
                .Show 1, Me
            End With
        Case "TablesDestinationsΛεωφορείων"
            With TablesDestinations
                .Tag = "True"
                .lblTitle.Caption = "Προορισμοί λεωφορείων"
                .txtShowInList.text = "2"
                .Show 1, Me
            End With
        Case "TablesPickupPoints"
            With TablesPickupPoints
                .Tag = "True"
                .Show 1, Me
            End With
        Case "ΒοηθητικάΠροορισμοίΔρομολόγια"
            With UtilsJoinDestinationsWithRoutes
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesPriceListsΠλοίων"
            With TablesPrices
                .Tag = "True"
                .lblTitle.Caption = "Τιμοκατάλογοι εκδρομών πλοίων"
                .txtShowInList.text = "1"
                .Show 1, Me
            End With
        Case "TablesPriceListsΛεωφορείων"
            With TablesPrices
                .Tag = "True"
                .lblTitle.Caption = "Τιμοκατάλογοι εκδρομών λεωφορείων"
                .txtShowInList.text = "2"
                .Show 1, Me
            End With
        Case "TablesPaymentTerms"
            With TablesPaymentTerms
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesPaymentWays"
            With TablesPaymentWays
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesExpenseCategories"
            With TablesExpenseCategories
                .Tag = "True"
                .Show 1, Me
            End With
        Case "UtilsCheckFiles"
            With UtilsCheckFiles
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesOccupantsDescriptions"
            With TablesOccupantsDescriptions
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesUsers"
            With TablesUsers
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesBanks"
            With TablesBanks
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesVATPercents"
            With TablesVATPercents
                .Tag = "True"
                .Show 1, Me
            End With
        'Τέλος βοηθητικά
        
        'Εξοδος
        Case "ΕξοδοςΑλλαγήΕταιρίας"
            With CommonLogin
                .Tag = "True"
                .Visible = True
            End With
        Case "ΕξοδοςΤερματισμόςΕφαρμογής"
            If CloseApp Then
                For Each obj In Forms
                    Unload obj
                Next
                End
            End If
        'Τέλος έξοδος
        
    End Select

End Sub

Private Function AddExpensesMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = CommonMain.vbExplorerBar.Bars.Add(, "Εξοδα", Space(5) & "Εξοδα")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "ΕξοδαΚινήσεις", " - Κινήσεις")
        Set cItem = cBar.Items.Add(, "ΕξοδαΗμερολόγιο", " - Ημερολόγιο")
        Set cItem = cBar.Items.Add(, "ΕξοδαΤύποιΠαραστατικών", " - Τύποι παραστατικών")

End Function




