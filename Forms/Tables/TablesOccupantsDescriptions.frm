VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesOccupantsDescriptions 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   8925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13875
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   13875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   12
      Top             =   7875
      Width           =   7515
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Δημιουργία"
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
         TabIndex        =   14
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Διαγραφή"
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
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
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   8625
      TabIndex        =   5
      Top             =   5850
      Width           =   4515
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "OccupantsDescriptions.OccupantDescriptionStatisticID"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox txtOccupantDescriptionStatisticID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox txtOccupantDescriptionID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "OccupantsDescriptions.OccupantDescriptionID"
         Top             =   75
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   825
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "TablesOccupantsDescriptions.frx":0000
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin iGrid300_10Tec.iGrid grdOccupantsDescriptions 
      Height          =   6240
      Left            =   8550
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1125
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   11007
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
   Begin UserControls.newText txtOccupantDescriptionDescription 
      Height          =   465
      Left            =   3150
      TabIndex        =   0
      Top             =   1125
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
   Begin UserControls.newText txtOccupantDescriptionStatisticDescription 
      Height          =   465
      Left            =   3150
      TabIndex        =   2
      Top             =   1650
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      Text            =   "ΝΑΙ"
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
      Left            =   3825
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1650
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
      PicNormal       =   "TablesOccupantsDescriptions.frx":1210
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   0
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   4
      Left            =   8550
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   1
      Left            =   8100
      Top             =   1125
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   7800
      Top             =   7350
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   3
      Left            =   2700
      Top             =   1125
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   13200
      Top             =   3600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3750
      Top             =   8550
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Ενημερώνει στατιστικά στοιχεία"
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
      TabIndex        =   10
      Top             =   1725
      Width           =   2265
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Χαρακτηρισμοί επιβαινόντων"
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
      TabIndex        =   4
      Top             =   75
      Width           =   6855
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Index           =   2
      Left            =   450
      TabIndex        =   3
      Top             =   1200
      Width           =   840
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
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   2
      Left            =   3150
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "TablesOccupantsDescriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim lngSelectedRow As Long

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtOccupantDescriptionID, txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticID, txtOccupantDescriptionStatisticDescription
            DisableFields txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticDescription
            grdOccupantsDescriptions.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", "OccupantsDescriptions", strApplicationName, "OccupantDescriptionID", txtOccupantDescriptionID.text, "True") Then
        PopulateGrid
        HighlightRow grdOccupantsDescriptions, lngSelectedRow, 1, "", True
        ClearFields txtOccupantDescriptionID, txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticID, txtOccupantDescriptionStatisticDescription
        DisableFields txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticDescription
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtOccupantDescriptionID, txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticID, txtOccupantDescriptionStatisticDescription
    EnableFields txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticDescription
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    txtOccupantDescriptionDescription.SetFocus

End Function

Private Function PopulateGrid()

    If FillGridFromDB("CommonDB", grdOccupantsDescriptions, "OccupantsDescriptions", "", "", "", 2, 0, 1) Then
        grdOccupantsDescriptions.SetFocus
        grdOccupantsDescriptions.SetCurCell 1, 1
    End If

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "OccupantsDescriptions", blnStatus, strApplicationName, "OccupantDescriptionID", txtOccupantDescriptionID.text, txtOccupantDescriptionDescription.text, txtOccupantDescriptionStatisticID.text, 1, strCurrentUser) <> 0 Then
        PopulateGrid
        HighlightRow grdOccupantsDescriptions, lngSelectedRow, 2, txtOccupantDescriptionDescription.text, True
        lngSelectedRow = 0
        ClearFields txtOccupantDescriptionID, txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticID, txtOccupantDescriptionStatisticDescription
        DisableFields txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticDescription
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Περιγραφή
    If Len(Trim(txtOccupantDescriptionDescription.text)) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtOccupantDescriptionDescription.SetFocus
        Exit Function
    End If
    
    'Ενημερώνει στατιστικά στοιχεία
    If txtOccupantDescriptionStatisticID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtOccupantDescriptionStatisticDescription.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Function SeekRecord()

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset

    Dim blnEnableDelete As Boolean
    
    If grdOccupantsDescriptions.RowCount = 0 Then Exit Function
    
    ClearFields txtOccupantDescriptionID, txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticID, txtOccupantDescriptionStatisticDescription
    DisableFields txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticDescription
    
    blnEnableDelete = SimpleSeek("Manifest", "TripOccupantDescriptionID", grdOccupantsDescriptions.CellValue(grdOccupantsDescriptions.CurRow, 1))
    
    If MainSeekRecord("CommonDB", "OccupantsDescriptions", "OccupantDescriptionID", grdOccupantsDescriptions.CellValue(grdOccupantsDescriptions.CurRow, 1), True, txtOccupantDescriptionID, txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticID) Then
        'Ενημερώνει στατιστικά στοιχεία
        Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtOccupantDescriptionStatisticID.text)
        txtOccupantDescriptionStatisticID.text = tmpRecordset.Fields(0)
        txtOccupantDescriptionStatisticDescription.text = tmpRecordset.Fields(1)
        '
        blnStatus = False
        lngSelectedRow = grdOccupantsDescriptions.CurRow
        EnableFields txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticDescription
        UpdateButtons Me, 4, 0, 1, IIf(blnEnableDelete, 1, 0), 1, 0
        txtOccupantDescriptionDescription.SetFocus
    End If
    
End Function

Private Sub cmdButton_Click(index As Integer)
                                                                
    Select Case index
        Case 0
            NewRecord
        Case 1
            SaveRecord
        Case 2
            DeleteRecord
        Case 3
            AbortProcedure False
        Case 4
            AbortProcedure True
    End Select

End Sub

Private Sub cmdIndex_Click(index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Ενημερώνει στατιστικά στοιχεία
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtOccupantDescriptionStatisticDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtOccupantDescriptionStatisticID.text = tmpTableData.strCode
                txtOccupantDescriptionStatisticDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdOccupantsDescriptions, 25, GetSetting(strApplicationName, "Layout Strings", "grdOccupantsDescriptions"), "04LNID,40LNName", "ID,Περιγραφή"
        Me.Refresh
        PopulateGrid
    End If

    'AddDummyLines grdOccupantsDescriptions, "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAΑΑΑΑΑΑΑΑΑΑ"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyEscape
            If cmdButton(3).Enabled Then cmdButton_Click 3: Exit Function
            If cmdButton(4).Enabled Then cmdButton_Click 4
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()
    
    UpdateColors Me, False
    SetUpGrid lstIconList, grdOccupantsDescriptions
    ClearFields txtOccupantDescriptionID, txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticID, txtOccupantDescriptionStatisticDescription
    DisableFields txtOccupantDescriptionDescription, txtOccupantDescriptionStatisticDescription
    UpdateButtons Me, 4, 1, 0, 0, 0, 1

End Sub

Private Sub grdOccupantsDescriptions_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub

Private Sub grdOccupantsDescriptions_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdOccupantsDescriptions_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()
    
    SaveSetting strApplicationName, "Layout Strings", "grdOccupantsDescriptions", grdOccupantsDescriptions.LayoutCol

End Sub

Private Sub txtOccupantDescriptionStatisticDescription_Change()

    If txtOccupantDescriptionStatisticDescription.text = "" Then
        ClearFields txtOccupantDescriptionStatisticID
    End If

End Sub


Private Sub txtOccupantDescriptionStatisticDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub


Private Sub txtOccupantDescriptionStatisticDescription_Validate(Cancel As Boolean)

    If txtOccupantDescriptionStatisticID.text = "" And txtOccupantDescriptionStatisticDescription.text <> "" Then cmdIndex_Click 0

End Sub


