VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Begin VB.Form CommonSelectPrinter 
   BackColor       =   &H0000C000&
   BorderStyle     =   0  'None
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   450
      TabIndex        =   5
      Top             =   4875
      Width           =   3090
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   150
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   8438015
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Επιλογή"
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
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   1
         Left            =   1575
         TabIndex        =   7
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
         ForeColor       =   0
         PicOpacity      =   0
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   4125
      TabIndex        =   2
      Top             =   2700
      Width           =   4515
      Begin VB.TextBox txtShowInList 
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   780
      End
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
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "WhatPrinterPrints"
         Top             =   225
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   600
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "CommonSelectPrinter.frx":0000
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin iGrid300_10Tec.iGrid grdPrinterSelect 
      Height          =   3315
      Left            =   450
      TabIndex        =   0
      Top             =   1050
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   5847
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
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   2250
      Top             =   5700
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   2625
      Top             =   4350
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   9375
      Top             =   2250
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   0
      Left            =   0
      Top             =   1725
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Επιλογή εκτυπωτή"
      BeginProperty Font 
         Name            =   "Aka-Acid-Steelfish"
         Size            =   30
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   225
      TabIndex        =   1
      Top             =   75
      Width           =   2895
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
Attribute VB_Name = "CommonSelectPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSavedLayout As String

Private Function PopulateGrid()

    If FillGridFromDB("PrintersDB", grdPrinterSelect, "Printers", "PrinterName, PrinterReportDetailLines, PrinterReportTopMargin, PrinterReportLeftMargin, PrinterFontName, PrinterFontSize, PrinterTypeDescription ", "INNER JOIN PrinterTypes ON Printers.PrinterTypeID = PrinterTypes.PrinterTypeID ", txtShowInList.text & " = 1", 1, 0, 1, 2, 3, 4, 5, 6) Then
        grdPrinterSelect.SetFocus
        grdPrinterSelect.SetCurCell 1, 1
    Else
        Unload Me
    End If
        
End Function

Private Function UpdateVariables()

    With grdPrinterSelect
        strPrinterName = .CellText(.CurRow, "Name")
        strPrinterType = .CellText(.CurRow, "Type")
        intPrinterReportDetailLines = .CellText(.CurRow, "DetailLines")
        intPrinterReportTopMargin = .CellText(.CurRow, "TopMargin")
        intPrinterReportLeftMargin = .CellText(.CurRow, "LeftMargin")
        strPrinterFontName = .CellText(.CurRow, "FontName")
        strPrinterFontSize = .CellText(.CurRow, "FontSize")
    End With

End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            UpdateVariables
            AbortProcedure True
        Case 1
            AbortProcedure True
    End Select
    
End Sub

Private Function AbortProcedure(blnStatus)

    If blnStatus Then
        Unload Me
    End If

End Function

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdPrinterSelect, 25, GetSetting(strAppTitle, "Layout Strings", "grdPrinterSelect"), "40NLNName,04NLNDetailLines,04NLNTopMargin,04NLNLeftMargin,40NLNFontName,04NLNFontSize,40NLNType", "Ονομα,ΑΓ,ΕΠ,ΑΠ,Γραμματοσειρά,Μ,Τύπος"
        Me.Refresh
        PopulateGrid
    End If
    
    'AddDummyLines grdPrinterSelect, "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "", "", "", "", "", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)

    Dim CtrlDown
    
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyE And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyEscape
            cmdButton_Click 1
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    UpdateColors Me, False, grdPrinterSelect
    SetUpGrid lstIconList, grdPrinterSelect

End Sub

Private Sub grdPrinterSelect_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    cmdButton_Click 0
    
End Sub

Private Sub grdPrinterSelect_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPrinterSelect_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then cmdButton_Click 0
    
End Sub


Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdPrinterSelect", grdPrinterSelect.LayoutCol

End Sub

