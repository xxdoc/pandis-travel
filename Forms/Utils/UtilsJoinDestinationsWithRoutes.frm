VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form UtilsJoinDestinationsWithRoutes 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19170
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1140
      Left            =   225
      TabIndex        =   24
      Top             =   5175
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "UtilsJoinDestinationsWithRoutes.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "UtilsJoinDestinationsWithRoutes.frx":001C
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
         TabIndex        =   26
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9765
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   17
         Top             =   9000
         Width           =   8940
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            Index           =   5
            Left            =   7350
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
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
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            Index           =   4
            Left            =   5925
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Ακυρο"
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
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Εκτύπωση"
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
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Δημιουργία αρχείου PDF"
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
         Height          =   1065
         Left            =   8250
         TabIndex        =   11
         Top             =   7800
         Width           =   4515
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
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
            Height          =   315
            Left            =   75
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Destinations.DestinationID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtDestinationID 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
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
            Height          =   315
            Left            =   3675
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   75
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   450
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   4592
            Images          =   "UtilsJoinDestinationsWithRoutes.frx":0038
            Version         =   131072
            KeyCount        =   4
            Keys            =   "ÿÿÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2565
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   6300
         Width           =   8040
         Begin UserControls.newText txtDestinationShortDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   1
            Top             =   825
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   4194304
            MaxLength       =   4
            Text            =   "AAAA"
            BackColor       =   4210688
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
         Begin UserControls.newText txtDestinationDescription 
            Height          =   465
            Left            =   2175
            TabIndex        =   2
            Top             =   1350
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   4194304
            MaxLength       =   40
            Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
            BackColor       =   4210688
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   3075
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   825
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   820
            BackColor       =   16777215
            ButtonShape     =   3
            ButtonStyle     =   2
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            PicNormal       =   "UtilsJoinDestinationsWithRoutes.frx":1248
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   7200
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1350
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   820
            BackColor       =   16777215
            ButtonShape     =   3
            ButtonStyle     =   2
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            PicNormal       =   "UtilsJoinDestinationsWithRoutes.frx":17E2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   1
            Left            =   7575
            Top             =   1125
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
            Left            =   1725
            Top             =   900
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   2
            Left            =   0
            Top             =   900
            Visible         =   0   'False
            Width           =   465
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
            Height          =   390
            Left            =   2775
            TabIndex        =   14
            Top             =   75
            Width           =   5115
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Συντ. προορισμού"
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
            Index           =   1
            Left            =   450
            TabIndex        =   9
            Top             =   900
            Width           =   1290
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Περιγραφή"
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
            Index           =   6
            Left            =   450
            TabIndex        =   8
            Top             =   1425
            Width           =   1290
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
            TabIndex        =   7
            Top             =   75
            Width           =   1665
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
            TabIndex        =   6
            Top             =   2100
            Width           =   8040
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
            TabIndex        =   10
            Top             =   0
            Width           =   8040
         End
      End
      Begin iGrid300_10Tec.iGrid grdDestinationsJoinRoutes 
         Height          =   7440
         Left            =   75
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1500
         Width           =   18840
         _ExtentX        =   33232
         _ExtentY        =   13123
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνδεση προορισμών με δρομολόγια λεωφορείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   30
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   720
         Left            =   75
         TabIndex        =   4
         Top             =   75
         Width           =   11085
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
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "UtilsJoinDestinationsWithRoutes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean


Private Function AddNextOrLastPageFooter(intProcessedDetailLines, intReportDetailLines, lngTotalDetailLines, intMessage)

    Dim intBlankLines As Integer
    
    For intBlankLines = intProcessedDetailLines To Val(intReportDetailLines)
        Print #1, ""
    Next intBlankLines
    
    If intProcessedDetailLines < lngTotalDetailLines Then
        Print #1, ""
        Print #1, strStandardMessages(intMessage)
    End If

End Function

Private Function AddPageHeader(intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2)

    PrintHeadings 107, intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
    PrintColumnHeadings 1, "ΠΕΡΙΓΡΑΦΗ", 52, "ΣΗΜΕΙΟ", 103, "ΩΡΑ"
    Print #1, ""

End Function

Private Function PrintRecords()

    If SelectPrinter("PrinterPrintsReports") Then
        CreateUnicodeFile "ΣΗΜΕΙΑ ΠΑΡΑΛΑΒΗΣ ΠΡΟΟΡΙΣΜΟΥ", txtDestinationDescription.text, "", intPrinterReportDetailLines
        With rptOneLiner
            If intPreviewReports = 1 Then
                .Restart
                .Zoom = -2
                .WindowState = vbMaximized
                .Show 1
            Else
                .PrintReport False
                .Run True
            End If
        End With
    End If

End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Προορισμός
    If txtDestinationID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtDestinationDescription.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)
    
    If blnProcessing Then blnProcessing = False: Exit Function

    If Not blnStatus Then
        'ClearFields txtDestinationID, txtDestinationShortDescription, txtDestinationDescription, grdDestinationsJoinRoutes
        ClearFields grdDestinationsJoinRoutes
        frmCriteria(0).Visible = True
        txtDestinationShortDescription.SetFocus
        UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function RefreshList()

    On Error GoTo ErrTrap
    
    'SQL
    Dim strSQL As String
    
    'Local variables
    Dim lngRow As Long
    
    'Areas
    Dim lngPickupRouteID As Long
    Dim strPickupRouteDescription As String
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'Αρχικές τιμές
    lngRow = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    
    'Πλέγμα
    With grdDestinationsJoinRoutes
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT PickupRouteID, PickupRouteShortDescription, PickupRouteDescription, PickUpPointID, PickUpPointHotelDescription, PickUpPointExactPoint, PickUpPointTime " _
        & "FROM PickupRoutes " _
        & "LEFT JOIN PickupPoints ON PickupRoutes.PickupRouteID = PickupPoints.PickUpPointRouteID " _
        & "ORDER BY PickupRouteDescription, PickUpPointTime "
        
    'Ανοίγω το recordset
    Set rstRecordset = CommonDB.OpenRecordset(strSQL)
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnErrors = False: RefreshList = False: Exit Function
    
    'Control-Break
    GoSub UpdateVariables
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 5, 0, 0, 0, 0, 1, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        lngRowCount = rstRecordset.RecordCount
        GoSub AddRouteLine
        Do While Not .EOF
            If lngPickupRouteID = !PickupRouteID Then
                grdDestinationsJoinRoutes.AddRow
                lngRow = lngRow + 1
                UpdateProgressBar Me
                grdDestinationsJoinRoutes.CellValue(lngRow, "PickupRouteID") = lngPickupRouteID
                grdDestinationsJoinRoutes.CellValue(lngRow, "PickupPointID") = !PickupPointID
                grdDestinationsJoinRoutes.CellValue(lngRow, "PickupRouteDescription") = strPickupRouteDescription
                grdDestinationsJoinRoutes.CellValue(lngRow, "PickupPointDescription") = !PickupPointHotelDescription
                grdDestinationsJoinRoutes.CellIndent(lngRow, "PickupPointDescription") = 30
                grdDestinationsJoinRoutes.CellValue(lngRow, "PickupPointExactPoint") = !PickupPointExactPoint
                grdDestinationsJoinRoutes.CellValue(lngRow, "PickupPointTime") = !PickupPointTime
                rstRecordset.MoveNext
                DoEvents
                If Not blnProcessing Then Exit Do
            Else
                GoSub UpdateVariables
                GoSub AddRouteLine
            End If
        Loop
        .Close
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdDestinationsJoinRoutes
        RefreshList = 0
    Else
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Τελικές ενέργειες
    cmdButton(4).Caption = "Νέα αναζήτηση"
    frmProgress.Visible = False
    
    Exit Function
    
UpdateVariables:
    lngPickupRouteID = rstRecordset!PickupRouteID
    strPickupRouteDescription = rstRecordset!PickupRouteDescription
    
    Return
    
AddRouteLine:
    grdDestinationsJoinRoutes.AddRow
    lngRow = grdDestinationsJoinRoutes.RowCount
    grdDestinationsJoinRoutes.CellValue(lngRow, "PickupRouteID") = lngPickupRouteID
    grdDestinationsJoinRoutes.CellValue(lngRow, "PickupRouteDescription") = strPickupRouteDescription
    grdDestinationsJoinRoutes.CellValue(lngRow, "PickupPointDescription") = strPickupRouteDescription
    grdDestinationsJoinRoutes.CellFont(lngRow, "PickupPointDescription").Bold = True
    
    Return
    
ErrTrap:
    blnErrors = True
    ClearFields grdDestinationsJoinRoutes, frmProgress
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function CreateUnicodeFile(strReportTitle, strReportSubTitle1, strReportSubTitle2, intReportDetailLines)

    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    Dim lngTotalDetailLines As Long
    Dim blnHeaderPrinted As Boolean
    Dim blnFirstPage As Boolean
    Dim strPickupRouteDescription As String
    
    intPageNo = 0
    
    With grdDestinationsJoinRoutes
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "SelectedPickupPointID") > 1 Then
                lngTotalDetailLines = lngTotalDetailLines + 1
            End If
        Next lngRow
    End With
    
    Open strUnicodeFile For Output As #1
    
    blnFirstPage = True
    
    With grdDestinationsJoinRoutes
        
        For lngRow = 1 To .RowCount
            
            'Αν το σημείο παραλαβής είναι επιλεγμένο
            If .CellIcon(lngRow, "SelectedPickupPointID") = lstIconList.ItemIndex(4) Then
                'Αν η περιγραφή της διαδρομής είναι διαφορετική από αυτήν που έχω κρατήσει
                If strPickupRouteDescription <> .CellText(lngRow, "PickupRouteDescription") Then
                    'Αν είμαι στην πρώτη σελίδα
                    If blnFirstPage Then
                        'Προσθέτω νέα σελίδα
                        intPageNo = intPageNo + 1
                        AddPageHeader intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
                        intProcessedDetailLines = 4
                        'Δεν είμαι πλέον στην πρώτη σελίδα
                        blnFirstPage = False
                    Else
                        'Αν δεν είμαι στην πρώτη σελίδα
                        If Not blnFirstPage Then
                            'Προσθέτω υποσέλιδο
                            AddNextOrLastPageFooter intProcessedDetailLines, intReportDetailLines, lngTotalDetailLines, 24
                            'Προσθέτω νέα σελίδα
                            intPageNo = intPageNo + 1
                            AddPageHeader intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
                            intProcessedDetailLines = 4
                        End If
                    End If
                    'Κρατάω την περιγραφή της διαδρομής
                    strPickupRouteDescription = .CellText(lngRow, "PickupRouteDescription")
                    AddPickupRouteDescription strPickupRouteDescription
                    'Αναλυτική γραμμή
                    intProcessedDetailLines = intProcessedDetailLines + 4
                End If
            End If
            
            'Αν το σημείο παραλαβής είναι επιλεγμένο
            If .CellIcon(lngRow, "SelectedPickupPointID") = lstIconList.ItemIndex(4) Then
                'Τυπώνω
                Print #1, .CellText(lngRow, "PickupPointDescription"); Tab(52); .CellText(lngRow, "PickupPointExactPoint"); Tab(103); .CellText(lngRow, "PickupPointTime")
                'Αναλυτική γραμμή
                intProcessedDetailLines = intProcessedDetailLines + 1
            End If
            
            'Αν έχω φτάσει στο όριο των γραμμών ανά σελίδα
            If intProcessedDetailLines > Val(intReportDetailLines) Then
                'Προσθέτω υποσέλιδο
                AddNextOrLastPageFooter intProcessedDetailLines, intReportDetailLines, lngTotalDetailLines, 24
                'Προσθέτω νέα σελίδα
                intPageNo = intPageNo + 1
                AddPageHeader intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
                'Προσθέτω την περιγραφή της διαδρομής (συνέχεια από την προηγούμενη σελίδα)
                AddPickupRouteDescription strPickupRouteDescription & " (συνέχεια)"
                intProcessedDetailLines = 8
            End If
            
        Next lngRow
        
        AddNextOrLastPageFooter intProcessedDetailLines, intReportDetailLines, lngTotalDetailLines, 26
    
    End With
    
    Close #1
    
End Function

Private Function AddPickupRouteDescription(strPickupRouteDescription)

    Print #1, strPickupRouteDescription
    Print #1, ""

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            If FindRoutesForDestination(txtDestinationID.text) Then
                blnProcessing = False
                EnableGrid grdDestinationsJoinRoutes, False
                HighlightRow grdDestinationsJoinRoutes, 1, 1, "", True
                UpdateButtons Me, 5, 0, 1, 1, 1, 1, 0
            Else
                UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
                If Not blnError Then
                    If blnProcessing Then
                        If MyMsgBox(4, strApplicationName, strStandardMessages(27), 1) Then
                        End If
                    Else
                        If MyMsgBox(1, strApplicationName, strStandardMessages(7), 1) Then
                        End If
                    End If
                End If
                blnError = False
                blnProcessing = False
                frmCriteria(0).Visible = True
                txtDestinationShortDescription.SetFocus
            End If
        Else
            UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
            If Not blnError Then
                If blnProcessing Then
                    If MyMsgBox(4, strApplicationName, strStandardMessages(27), 1) Then
                    End If
                Else
                    If MyMsgBox(1, strApplicationName, strStandardMessages(7), 1) Then
                    End If
                End If
            End If
            blnError = False
            blnProcessing = False
            frmCriteria(0).Visible = True
            txtDestinationShortDescription.SetFocus
        End If
    End If
    
End Function

Private Function SaveRecord()
    
    BeginTrans
        DeleteExistingRoutesAndPickupPointsForDestination (txtDestinationID.text)
        AddSelectedRoutesAndPickupPointsForDestination (txtDestinationID.text)
    CommitTrans
        
    If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
    End If
    
    ClearFields grdDestinationsJoinRoutes
    EnableFields txtDestinationShortDescription, txtDestinationDescription, cmdIndex(0), cmdIndex(1)
    frmCriteria(0).Visible = True
    txtDestinationShortDescription.SetFocus
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    
End Function

Private Sub cmdButton_Click(index As Integer)
                                                                
    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            SaveRecord
        Case 2
            PrintRecords
        Case 3
            ExportRecords
        Case 4
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select

End Sub

Private Function ExportRecords()

    Dim pdf As New ARExportPDF
    
    CreateUnicodeFile "ΣΗΜΕΙΑ ΠΑΡΑΛΑΒΗΣ ΠΡΟΟΡΙΣΜΟΥ", txtDestinationDescription.text, "", GetSetting(strApplicationName, "Settings", "Export Report Height")
    
    With rptOneLiner
        .Restart
        .Run False
        pdf.AcrobatVersion = 2
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = strReportsPathName & "ΣΗΜΕΙΑ ΠΑΡΑΛΑΒΗΣ ΠΡΟΟΡΙΣΜΟΥ " & txtDestinationDescription.text & ".pdf"
        pdf.Export .Pages
        If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
        End If
    End With
    
End Function


Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Συντ. προορισμού
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationShortDescription", "String", txtDestinationShortDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 5, 40, 1, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationShortDescription.text = tmpTableData.strFirstField
                txtDestinationDescription.text = tmpTableData.strSecondField
            End If
        Case 1
            'Περιγραφή προορισμού
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 5, 40, 1, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationShortDescription.text = tmpTableData.strFirstField
                txtDestinationDescription.text = tmpTableData.strSecondField
            End If
    End Select

End Sub

Private Sub cmdIndex_Validate(index As Integer, Cancel As Boolean)

    If txtDestinationShortDescription.text <> "" Then cmdIndex_Click 0
    
End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdDestinationsJoinRoutes, 44, GetSetting(strApplicationName, "Layout Strings", "grdDestinationsJoinRoutes"), "04NCISelectedPickupRouteID,04NCISelectedPickupPointID,04NCIPickupRouteID,04NCIPickupPointID,40NLNPickupRouteDescription,40NLNPickupPointDescription,40NLNPickupPointExactPoint,06NCTPickupPointTime", "-,-,-,-,-,Περιγραφή,Σημείο,Ώρα"
        Me.Refresh
        frmCriteria(0).Visible = True
        txtDestinationShortDescription.SetFocus
    End If
    
    'AddDummyLines grdDestinationsJoinRoutes, "SAS", "SAS", "99999", "99999", "", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "A00:00A"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyP And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 2
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()
    
    PositionControls Me, True, grdDestinationsJoinRoutes
    ColorizeControls Me, True
    SetUpGrid lstIconList, grdDestinationsJoinRoutes
    ClearFields txtDestinationID
    ClearFields txtDestinationShortDescription, txtDestinationDescription
    EnableFields txtDestinationShortDescription, txtDestinationDescription
    EnableFields cmdIndex(0), cmdIndex(1)
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    
End Sub

Private Sub grdDestinationsJoinRoutes_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdDestinationsJoinRoutes_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdDestinationsJoinRoutes.RowCount > 0 Then
        If grdDestinationsJoinRoutes.CellValue(grdDestinationsJoinRoutes.CurRow, "PickupPointID") <> "" Then
            grdDestinationsJoinRoutes.CellIcon(grdDestinationsJoinRoutes.CurRow, "SelectedPickupPointID") = IIf(grdDestinationsJoinRoutes.CellIcon(grdDestinationsJoinRoutes.CurRow, "SelectedPickupPointID") <> 3, lstIconList.ItemIndex(4), lstIconList.ItemIndex(1))
        End If
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdDestinationsJoinRoutes", grdDestinationsJoinRoutes.LayoutCol

End Sub

Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then
        ClearFields txtDestinationID, txtDestinationShortDescription
    End If

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 1
    
End Sub

Private Function FindRoutesForDestination(destinationID As String)

    'Local variables
    Dim lngRow As Long
    Dim rstRecordset As Recordset
    Dim strSQL As String
    
    'Αρχικές τιμές
    grdDestinationsJoinRoutes.Redraw = False
    
    'Βρίσκω τα δρομολόγια για τον επιλεγμένο προορισμό στον 'ενδιάμεσο' πίνακα
    strSQL = "SELECT DestinationID, RouteID, PickupPointID " _
        & "FROM DestinationsRoutesPickupPoints " _
        & "WHERE DestinationID = " & Val(destinationID) & " " _
        & "ORDER BY RouteID, PickupPointID"
    Set rstRecordset = CommonDB.OpenRecordset(strSQL)
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnErrors = False: FindRoutesForDestination = False: Exit Function
    'If rstRecordset.RecordCount <> 0 Then blnErrors = False: FindRoutesForDestination = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 5, 0, 0, 0, 0, 1, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Σημαδεύω τα δρομολόγια αν υπάρχουν στον 'ενδιάμεσο' πίνακα
    With rstRecordset
        Do While Not .EOF
            UpdateProgressBar Me
            For lngRow = 1 To grdDestinationsJoinRoutes.RowCount
                'Σημείο παραλαβής
                If !RouteID = grdDestinationsJoinRoutes.CellValue(lngRow, "PickupRouteID") And grdDestinationsJoinRoutes.CellValue(lngRow, "PickupPointID") = !PickupPointID Then
                    grdDestinationsJoinRoutes.CellIcon(lngRow, "SelectedPickupPointID") = lstIconList.ItemIndex(4)
                End If
            Next lngRow
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
    End With
    
    cmdButton(4).Caption = "Νέα αναζήτηση"
    frmProgress.Visible = False
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdDestinationsJoinRoutes
        FindRoutesForDestination = False
    Else
        FindRoutesForDestination = True
        blnProcessing = False
    End If
    
End Function

Private Function DeleteExistingRoutesAndPickupPointsForDestination(destinationID As String)

    Dim lngRow As Long
    Dim strSQL As String
    
    strSQL = "DELETE " _
        & "FROM DestinationsRoutesPickupPoints " _
        & "WHERE DestinationID = " & Val(destinationID)
    CommonDB.Execute (strSQL)
    
End Function

Private Function AddSelectedRoutesAndPickupPointsForDestination(strDestinationID As String)

    Dim lngRow As Long
    Dim rsDestinationsRoutesPickupPoints As Recordset
    Dim strSQL As String
    
    Set rsDestinationsRoutesPickupPoints = CommonDB.OpenRecordset("DestinationsRoutesPickupPoints")
        
    For lngRow = 1 To grdDestinationsJoinRoutes.RowCount
        If grdDestinationsJoinRoutes.CellIcon(lngRow, "SelectedPickupPointID") = 3 Then
            rsDestinationsRoutesPickupPoints.AddNew
            rsDestinationsRoutesPickupPoints!destinationID = Val(strDestinationID)
            rsDestinationsRoutesPickupPoints!RouteID = Val(grdDestinationsJoinRoutes.CellValue(lngRow, "PickupRouteID"))
            rsDestinationsRoutesPickupPoints!PickupPointID = Val(grdDestinationsJoinRoutes.CellValue(lngRow, "PickupPointID"))
            rsDestinationsRoutesPickupPoints.Update
        End If
    Next lngRow
    
    rsDestinationsRoutesPickupPoints.Close

End Function

Private Sub txtDestinationShortDescription_Change()

    If txtDestinationShortDescription.text = "" Then
        ClearFields txtDestinationID, txtDestinationDescription
    End If
    
End Sub

Private Sub txtDestinationShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtDestinationShortDescription_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestinationShortDescription.text <> "" Then cmdIndex_Click 0

End Sub

