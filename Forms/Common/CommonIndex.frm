VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Begin VB.Form CommonIndex 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8415
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3240
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmButtonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   75
      TabIndex        =   1
      Top             =   7350
      Width           =   3090
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   150
         TabIndex        =   3
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
         TabIndex        =   4
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
   Begin vbalIml6.vbalImageList lstIconList 
      Left            =   375
      Top             =   6450
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   4592
      Images          =   "CommonIndex.frx":0000
      Version         =   131072
      KeyCount        =   4
      Keys            =   "ÿÿÿ"
   End
   Begin iGrid300_10Tec.iGrid grdGrid 
      Height          =   6165
      Left            =   300
      TabIndex        =   2
      Top             =   900
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   10874
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   1
      DefaultRowHeight=   20
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
      ScrollBarStyle  =   2
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ευρετήριο"
      BeginProperty Font 
         Name            =   "Aka-Acid-Steelfish"
         Size            =   26.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Left            =   300
      TabIndex        =   0
      Top             =   75
      Width           =   1470
   End
   Begin VB.Shape shpShape 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   6315
      Left            =   225
      Top             =   825
      Width           =   2565
   End
End
Attribute VB_Name = "CommonIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            Me.Hide
        Case 1
            AbortProcedure
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function AbortProcedure()
    
    Dim lngCol As Long
    
    If cmdButton(1).Enabled Then
        For lngCol = 1 To grdGrid.ColCount
            grdGrid.CellValue(CommonIndex.grdGrid.CurRow, lngCol) = ""
        Next lngCol
    End If
    
    Me.Hide
    
End Function

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Select Case KeyCode
        Case vbKeyReturn
            cmdButton_Click 0
        Case vbKeyEscape
            cmdButton_Click 1
    End Select
    
End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdGrid
    ColorizeGrid grdGrid

End Sub

Private Sub grdGrid_ColHeaderMouseEnter(ByVal lCol As Long)

    grdGrid.Header.Buttons = True

End Sub

Private Sub grdGrid_ColHeaderMouseLeave(ByVal lCol As Long)

    grdGrid.Header.Buttons = False

End Sub


Private Sub grdGrid_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    Me.Hide

End Sub

Private Sub grdGrid_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then Me.Hide

End Sub
