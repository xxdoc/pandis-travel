VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Begin VB.Form CommonMessages 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6540
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   840
      Index           =   1
      Left            =   2775
      TabIndex        =   5
      Top             =   2250
      Width           =   1815
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   2
         Left            =   225
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   8421631
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "OK"
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
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   840
      Index           =   2
      Left            =   2125
      TabIndex        =   2
      Top             =   2250
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
         Caption         =   "Ναι"
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
         Caption         =   "Οχι"
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
   Begin vbalIml6.vbalImageList lslIcons 
      Left            =   5700
      Top             =   2400
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   48
      IconSizeY       =   48
      ColourDepth     =   24
      Size            =   38640
      Images          =   "CommonMessages.frx":0000
      Version         =   131072
      KeyCount        =   4
      Keys            =   "ÿÿÿ"
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      Height          =   1290
      Left            =   1050
      Top             =   825
      Width           =   5265
   End
   Begin VB.Label lblLine 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα Μήνυμα ΜήνυμαΜήνυμα"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   765
      Left            =   1125
      TabIndex        =   1
      Top             =   1125
      Width           =   5115
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgImage 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   225
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pandis Travel"
      BeginProperty Font 
         Name            =   "Aka-Acid-Steelfish"
         Size            =   26.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   630
      Left            =   150
      TabIndex        =   0
      Top             =   100
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3300
      Left            =   20
      Top             =   20
      Width           =   6510
   End
End
Attribute VB_Name = "CommonMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdButton_Click(Index As Integer)

    If Index = 0 Then
        cmdButton(0).Tag = "Pressed"
        Me.Hide
    End If
    If Index = 1 Then
        cmdButton(0).Tag = "NotPressed"
        Me.Hide
    End If
    If Index = 2 Then
        cmdButton(2).Tag = "Pressed"
        Me.Hide
    End If

End Sub

Private Sub cmdButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        cmdButton(0).Tag = "NotPressed"
        Unload Me
    End If
        
    If KeyCode = 37 Or KeyCode = 39 Then
        If Index = 0 Then
            cmdButton(1).SetFocus
        End If
        If Index = 1 Then
            cmdButton(0).SetFocus
        End If
    End If

End Sub

Private Sub Form_Load()
    
    frmButtonFrame(1).BackColor = Shape1.BackColor
    frmButtonFrame(2).BackColor = Shape1.BackColor
    
    cmdButton(0).Tag = "NotPressed"
    
End Sub

