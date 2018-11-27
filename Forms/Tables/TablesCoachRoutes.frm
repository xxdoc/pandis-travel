VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesCoachRoutes 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   17460
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   17460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   9
      Top             =   7875
      Width           =   7515
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   9600
      TabIndex        =   6
      Top             =   6225
      Width           =   4515
      Begin VB.TextBox txtPickupRouteID 
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
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "PickupRoutes.PickupRouteID"
         Top             =   75
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   450
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "TablesCoachRoutes.frx":0000
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin UserControls.newText txtPickupRouteDescription 
      Height          =   465
      Left            =   1725
      TabIndex        =   1
      Top             =   1650
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   60
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
   Begin iGrid300_10Tec.iGrid grdPickupRoutes 
      Height          =   6240
      Left            =   9525
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1125
      Width           =   6765
      _ExtentX        =   11933
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
   Begin UserControls.newText txtPickupRouteShortDescription 
      Height          =   465
      Left            =   1725
      TabIndex        =   5
      Top             =   1125
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   3
      Text            =   "ΑΑΑ"
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
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   16275
      Top             =   3225
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   5400
      Top             =   8550
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   9225
      Top             =   7350
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   2
      Left            =   9075
      Top             =   1500
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
      Left            =   1275
      Top             =   1575
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   12
      Left            =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   13
      Left            =   10200
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Δρομολόγια λεωφορείων"
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
      TabIndex        =   3
      Top             =   75
      Width           =   5895
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
      Index           =   3
      Left            =   450
      TabIndex        =   2
      Top             =   1725
      Width           =   840
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Συντ."
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
      TabIndex        =   0
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
      Index           =   0
      Left            =   1800
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
Attribute VB_Name = "TablesCoachRoutes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim lngSelectedRow As Long

Private Function ValidateFields()

    ValidateFields = False
    
    'Συντ.
    If Len(txtPickupRouteShortDescription.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPickupRouteShortDescription.SetFocus
        Exit Function
    End If
    If Len(Trim(txtPickupRouteShortDescription.text)) = 0 Or Left(txtPickupRouteShortDescription.text, 1) = " " Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtPickupRouteShortDescription.SetFocus
        Exit Function
    End If
    
    'Περιγραφή
    If Len(txtPickupRouteDescription.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPickupRouteDescription.SetFocus
        Exit Function
    End If
    If Len(Trim(txtPickupRouteDescription.text)) = 0 Or Left(txtPickupRouteDescription.text, 1) = " " Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtPickupRouteDescription.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtPickupRouteID, txtPickupRouteShortDescription, txtPickupRouteDescription
            DisableFields txtPickupRouteShortDescription, txtPickupRouteDescription
            grdPickupRoutes.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", "PickupRoutes", strApplicationName, "PickupRouteID", txtPickupRouteID.text, "True") Then
        PopulateGrid
        HighlightRow grdPickupRoutes, lngSelectedRow, 1, "", True
        ClearFields txtPickupRouteID, txtPickupRouteShortDescription, txtPickupRouteDescription
        DisableFields txtPickupRouteShortDescription, txtPickupRouteDescription
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If
    
End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtPickupRouteID, txtPickupRouteShortDescription, txtPickupRouteDescription
    EnableFields txtPickupRouteShortDescription, txtPickupRouteDescription
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    txtPickupRouteShortDescription.SetFocus

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "PickupRoutes", blnStatus, strApplicationName, "PickupRouteID", txtPickupRouteID.text, txtPickupRouteShortDescription.text, txtPickupRouteDescription.text, 1, strCurrentUser) <> 0 Then
        PopulateGrid
        HighlightRow grdPickupRoutes, lngSelectedRow, 2, txtPickupRouteShortDescription.text, True
        lngSelectedRow = 0
        ClearFields txtPickupRouteID, txtPickupRouteShortDescription, txtPickupRouteDescription
        DisableFields txtPickupRouteShortDescription, txtPickupRouteDescription
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If
    
End Function

Private Function SeekRecord()

    Dim blnEnableDelete As Boolean
    
    If grdPickupRoutes.RowCount = 0 Then Exit Function
    
    ClearFields txtPickupRouteID, txtPickupRouteShortDescription, txtPickupRouteDescription
    DisableFields txtPickupRouteShortDescription, txtPickupRouteDescription
    
    blnEnableDelete = SimpleSeek("PickupPoints", "PickupPointRouteID", grdPickupRoutes.CellValue(grdPickupRoutes.CurRow, 1))
    
    If MainSeekRecord("CommonDB", "PickupRoutes", "PickupRouteID", grdPickupRoutes.CellValue(grdPickupRoutes.CurRow, 1), True, txtPickupRouteID, txtPickupRouteShortDescription, txtPickupRouteDescription) Then
        blnStatus = False
        lngSelectedRow = grdPickupRoutes.CurRow
        EnableFields txtPickupRouteShortDescription, txtPickupRouteDescription
        UpdateButtons Me, 4, 0, 1, IIf(blnEnableDelete, 1, 0), 1, 0
        txtPickupRouteShortDescription.SetFocus
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

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdPickupRoutes, 25, GetSetting(strApplicationName, "Layout Strings", "grdPickupRoutes"), "04NCIID,04NCNShortDescription,40NLNDescription", "ID,Συντ.,Περιγραφή"
        Me.Refresh
        PopulateGrid
    End If
    
    'AddDummyLines grdPickupRoutes, "99999", "AAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"

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
    SetUpGrid lstIconList, grdPickupRoutes
    ClearFields txtPickupRouteID, txtPickupRouteShortDescription, txtPickupRouteDescription
    DisableFields txtPickupRouteShortDescription, txtPickupRouteDescription
    UpdateButtons Me, 4, 1, 0, 0, 0, 1
    
End Sub

Private Sub grdPickupRoutes_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub

Private Sub grdPickupRoutes_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPickupRoutes_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdPickupRoutes", grdPickupRoutes.LayoutCol

End Sub

Private Function PopulateGrid()
        
    If FillGridFromDB("CommonDB", grdPickupRoutes, "PickupRoutes", "", "", "", 3, 0, 1, 2) Then
        grdPickupRoutes.SetFocus
        grdPickupRoutes.SetCurCell 1, 1
    End If

End Function

