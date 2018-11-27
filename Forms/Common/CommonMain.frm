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
      Keys            =   "�"
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
      Keys            =   "�����"
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

    Set cBar = vbExplorerBar.Bars.Add(, "��������", Space(5) & "��������")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "��������������", " - ������")
        Set cItem = cBar.Items.Add(, "����������������", " - ��������")
        Set cItem = cBar.Items.Add(, "���������������", " - �������")
        Set cItem = cBar.Items.Add(, "����������������", " - ��������")
        Set cItem = cBar.Items.Add(, "�������������������������", " - ����� ������������")

End Function

Private Function AddDebitorsSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = vbExplorerBar.Bars.Add(, "��������", Space(5) & "��������")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "��������������", " - ������")
        Set cItem = cBar.Items.Add(, "������������������������������", " - �������� �������� ������")
        Set cItem = cBar.Items.Add(, "����������������������������������", " - �������� �������� ����������")
        Set cItem = cBar.Items.Add(, "���������������", " - �������")
        Set cItem = cBar.Items.Add(, "����������������", " - ��������")
        Set cItem = cBar.Items.Add(, "���������������������������������������", " - ����� ������������ �������� ������")
        Set cItem = cBar.Items.Add(, "�������������������������������������������", " - ����� ������������ �������� ����������")

End Function

Private Function AddExitSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = vbExplorerBar.Bars.Add(, "������", Space(5) & "������")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "��������������������", "- ������ ��������")
        Set cItem = cBar.Items.Add(, "��������������������������", "- ����������� ���������")

End Function

Private Function AddIncomeMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = vbExplorerBar.Bars.Add(, "�����", Space(5) & "�����")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "�������������������", "�������� ������")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "���������������������������", " - ��������")
            Set cItem = cBar.Items.Add(, "�����������������������������", " - ����������")
        Set cItem = cBar.Items.Add(, "�����������������������", "�������� ����������")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "�������������������������������", " - ��������")
            Set cItem = cBar.Items.Add(, "���������������������������������", " - ����������")
        Set cItem = cBar.Items.Add(, "����������������������", "����� ������������")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "��������������������������������", " - ����������")

End Function

Private Function AddShipsPassengers()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
        
    Set cBar = vbExplorerBar.Bars.Add(, "������������������", Space(5) & "������������ ������")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "��������������������������", " - ������")
        Set cItem = cBar.Items.Add(, "���������������������������", " - ��������� ��������")
        Set cItem = cBar.Items.Add(, "������������������������������������", " - ���������� �� ���� ��� ��������")
        Set cItem = cBar.Items.Add(, "���������������������������������������", " - ���������� �� ���� ���� ������������")

End Function

Private Function AddTransfersSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
        
    Set cBar = vbExplorerBar.Bars.Add(, "�������������������", Space(5) & "������������ ����������")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "������������������������������������", " - ������� ����������")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "��������������������������������������������", Space(5) & " - ��������")
            Set cItem = cBar.Items.Add(, "����������������������������������������������", Space(5) & " - ��������� ��������")
        Set cItem = cBar.Items.Add(, "�������������������������������������", " - �������� ����������")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "���������������������������������������������", Space(5) & " - ��������")
            Set cItem = cBar.Items.Add(, "�����������������������������������������������", Space(5) & " - ��������� ��������")

End Function

Private Function AddUtilsSubMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem
        
    Set cBar = vbExplorerBar.Bars.Add(, "���������", Space(5) & "���������")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "���������������", "���������������")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "�������������������", Space(5) & " - ������� ����������")
            Set cItem = cBar.Items.Add(, "TablesPrinters", Space(5) & " - ���������")
        Set cItem = cBar.Items.Add(, "�������", "�������")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "TablesShipRoutes", Space(5) & " - ���������� ������")
            Set cItem = cBar.Items.Add(, "TablesCoachRoutes", Space(5) & " - ���������� ����������")
            Set cItem = cBar.Items.Add(, "TablesExpenseCategories", Space(5) & " - ���������� ������")
            Set cItem = cBar.Items.Add(, "TablesTaxOffices", Space(5) & " - ����������� ���������")
            Set cItem = cBar.Items.Add(, "TablesPaymentTerms", Space(5) & " - ���� ��������")
            Set cItem = cBar.Items.Add(, "TablesShips", Space(5) & " - �����")
            Set cItem = cBar.Items.Add(, "TablesVATPercents", Space(5) & " - ������� �.�.�.")
            Set cItem = cBar.Items.Add(, "TablesDestinations������", Space(5) & " - ���������� ������")
            Set cItem = cBar.Items.Add(, "TablesDestinations����������", Space(5) & " - ���������� ����������")
            Set cItem = cBar.Items.Add(, "TablesPickupPoints", Space(5) & " - ������ ��������� ��������")
            Set cItem = cBar.Items.Add(, "UtilsJoinDestinationsWithRoutes", Space(5) & " - ������� ���������� �� ���������� ����������")
            Set cItem = cBar.Items.Add(, "TablesPriceLists������", Space(5) & " - ������������� �������� ������")
            Set cItem = cBar.Items.Add(, "TablesPriceLists����������", Space(5) & " - ������������� �������� ����������")
            Set cItem = cBar.Items.Add(, "TablesBanks", Space(5) & " - ��������")
            Set cItem = cBar.Items.Add(, "TablesPaymentWays", Space(5) & " - ������ ��������")
            Set cItem = cBar.Items.Add(, "TablesOccupantsDescriptions", Space(5) & " - ������������� ������������")
            Set cItem = cBar.Items.Add(, "TablesUsers", Space(5) & " - �������")
        Set cItem = cBar.Items.Add(, "��������", "��������")
            cItem.ItemType = eItemText
            cItem.Bold = True
            cItem.TextColor = RGB(96, 150, 207)
            Set cItem = cBar.Items.Add(, "UtilsSalesExport", Space(5) & " - ���������� ������� ��������")
            Set cItem = cBar.Items.Add(, "UtilsCheckFiles", Space(5) & " - ������� �������")

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
    
    '������� ����������� ��� �� ����� ����������, ���� ��� � � ALT-F4
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
    
    '������� ����������� ��� ��� ������� ������ > �����������
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

    '������ �����������
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
        '�����
        Case "���������������������������"
            With InvoicesOut
                .lblTitle.Caption = "�������� ������"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "�����������������������������"
            With InvoicesOutIndex
                .lblTitle.Caption = "���������� �������� ������"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "�������������������������������"
            With InvoicesOut
                .lblTitle.Caption = "�������� ����������"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .lblLabel(5).Visible = False
                .txtShipDescription.Visible = False
                .cmdIndex(3).Visible = False
                .cmdIndex(8).Visible = False
                .Tag = "True"
                .Show 1, Me
            End With
        Case "���������������������������������"
            With InvoicesOutIndex
                .lblTitle.Caption = "���������� �������� ����������"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "��������������������������������"
            With TablesCodes
                .lblTitle.Caption = "����� ������������ ������"
                .txtCodeMasterRefersTo.text = "2"
                .txtCodeSecondaryRefersTo.text = "0"
                .Tag = "True"
                .Show 1, Me
            End With
        '����� �����
        
        '�����
        Case "�������������"
            With InvoicesIn
                .Tag = "True"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtInvoiceSecondaryRefersTo.text = ""
                .Show 1, Me
            End With
        Case "���������������"
            With InvoicesInIndex
                .txtInvoiceMasterRefersTo.text = "1"
                .txtInvoiceSecondaryRefersTo.text = ""
                .Tag = "True"
                .Show 1, Me
            End With
        Case "����������������������"
            With TablesCodes
                .lblTitle.Caption = "����� ������������ ������"
                .txtCodeMasterRefersTo.text = "1"
                .txtCodeSecondaryRefersTo.text = "0"
                .Tag = "True"
                .Show 1, Me
            End With
        '����� �����
        
        '��������
        Case "��������������"
            With persons
                .Tag = "True"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtInvoiceMasterRefersTo.text = "2"
                .lblTitle.Caption = "�������"
                .Show 1, Me
            End With
        Case "������������������������������"
            With PersonsTransactions
                .lblTitle.Caption = "�������� �������� ������"
                .txtInvoiceMasterRefersTo.text = "4"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "����������������������������������"
            With PersonsTransactions
                .lblTitle.Caption = "�������� �������� ����������"
                .txtInvoiceMasterRefersTo.text = "4"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "���������������"
            With PersonsLedger
                .lblTitle.Caption = "������� �������"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "����������������"
            With PersonsBalanceSheet
                .lblTitle.Caption = "�������� ��������"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "���������������������������������������"
            With TablesCodes
                .lblTitle.Caption = "����� ������������ �������� ������"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "4"
                .txtCodeSecondaryRefersTo.text = "1"
                .Show 1, Me
            End With
        Case "�������������������������������������������"
            With TablesCodes
                .lblTitle.Caption = "����� ������������ �������� ����������"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "4"
                .txtCodeSecondaryRefersTo.text = "2"
                .Show 1, Me
            End With
        '����� ��������
        
        '��������
        Case "��������������"
            With persons
                .Tag = "True"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtInvoiceMasterRefersTo.text = "1"
                .lblTitle.Caption = "��������"
                .Show 1, Me
            End With
        Case "����������������"
            With PersonsTransactions
                .lblTitle.Caption = "�������� ��������"
                .txtInvoiceMasterRefersTo.text = "3"
                .txtInvoiceSecondaryRefersTo.text = ""
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "���������������"
            With PersonsLedger
                .lblTitle.Caption = "������� �������"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "����������������"
            With PersonsBalanceSheet
                .lblTitle.Caption = "�������� ��������"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "�������������������������"
            With TablesCodes
                .lblTitle.Caption = "����� ������������ ��������"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "3"
                .Show 1, Me
            End With
        '����� ��������
            
        '������������ ������
        Case "��������������������������"
            With ShipsTransactions
                .Tag = "True"
                .Show 1, Me
            End With
        Case "���������������������������"
            With ShipsRouteReport
                .Tag = "True"
                .Show 1, Me
            End With
        Case "������������������������������������"
            With ShipsStatistics
                .lblTitle.Caption = "���������� �� ���� ��� ��������"
                .txtTable.text = "Sales"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "���������������������������������������"
            With ShipsStatistics
                .lblTitle.Caption = "���������� �� ���� ���� ������������"
                .txtTable.text = "Manifest"
                .Tag = "True"
                .Show 1, Me
            End With
        '����� ������������ ������
        
        '��������� ����������
        Case "��������������������������������������������"
            With CoachesPickupsBrief
                .txtRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "����������������������������������������������"
            With CoachesReport
                .Tag = "True"
                .txtRefersTo.text = "2"
                .txtCallingForm = "MainMenu"
                .grdCoachesReport.Tag = "grdCoachesReportBrief"
                .Show 1, Me
            End With
        Case "���������������������������������������������"
            With CoachesPickupsStandard
                .txtRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case "�����������������������������������������������"
            With CoachesReport
                .Tag = "True"
                .txtRefersTo.text = "1"
                .txtCallingForm = "MainMenu"
                .grdCoachesReport.Tag = "grdCoachesReportStandard"
                .Show 1, Me
            End With
        '����� ��������� ����������
        
        '���������
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
        Case "�������������������"
            With TablesSettings
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesShips"
            With TablesShips
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesDestinations������"
            With TablesDestinations
                .Tag = "True"
                .lblTitle.Caption = "���������� ������"
                .txtShowInList.text = "1"
                .Show 1, Me
            End With
        Case "TablesDestinations����������"
            With TablesDestinations
                .Tag = "True"
                .lblTitle.Caption = "���������� ����������"
                .txtShowInList.text = "2"
                .Show 1, Me
            End With
        Case "TablesPickupPoints"
            With TablesPickupPoints
                .Tag = "True"
                .Show 1, Me
            End With
        Case "�����������������������������"
            With UtilsJoinDestinationsWithRoutes
                .Tag = "True"
                .Show 1, Me
            End With
        Case "TablesPriceLists������"
            With TablesPrices
                .Tag = "True"
                .lblTitle.Caption = "������������� �������� ������"
                .txtShowInList.text = "1"
                .Show 1, Me
            End With
        Case "TablesPriceLists����������"
            With TablesPrices
                .Tag = "True"
                .lblTitle.Caption = "������������� �������� ����������"
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
        '����� ���������
        
        '������
        Case "��������������������"
            With CommonLogin
                .Tag = "True"
                .Visible = True
            End With
        Case "��������������������������"
            If CloseApp Then
                For Each obj In Forms
                    Unload obj
                Next
                End
            End If
        '����� ������
        
    End Select

End Sub

Private Function AddExpensesMenu()

    Dim cBar As cExplorerBar
    Dim cItem As cExplorerBarItem

    Set cBar = CommonMain.vbExplorerBar.Bars.Add(, "�����", Space(5) & "�����")
        cBar.IsSpecial = True
        cBar.State = eBarCollapsed
        Set cItem = cBar.Items.Add(, "�������������", " - ��������")
        Set cItem = cBar.Items.Add(, "���������������", " - ����������")
        Set cItem = cBar.Items.Add(, "����������������������", " - ����� ������������")

End Function




