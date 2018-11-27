VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesPickupPoints 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   15
   ClientTop       =   15
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
      TabIndex        =   30
      Top             =   4350
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "TablesPickupPoints.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "TablesPickupPoints.frx":001C
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
         TabIndex        =   32
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   20
         Top             =   8850
         Width           =   8940
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   21
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
            TabIndex        =   22
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
            TabIndex        =   23
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
            Index           =   2
            Left            =   3075
            TabIndex        =   24
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
            Index           =   4
            Left            =   5925
            TabIndex        =   25
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
            Index           =   3
            Left            =   4500
            TabIndex        =   26
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
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2565
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   5475
         Width           =   10740
         Begin UserControls.newText txtRouteShortDescription 
            Height          =   465
            Left            =   2475
            TabIndex        =   1
            Top             =   825
            Width           =   765
            _ExtentX        =   1349
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
         Begin UserControls.newText txtRouteDescription 
            Height          =   465
            Left            =   2475
            TabIndex        =   2
            Top             =   1350
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   820
            ForeColor       =   4194304
            MaxLength       =   60
            Text            =   "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
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
            Left            =   3300
            TabIndex        =   18
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
            PicNormal       =   "TablesPickupPoints.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   9900
            TabIndex        =   19
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
            PicNormal       =   "TablesPickupPoints.frx":05D2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   4
            Left            =   2625
            Top             =   525
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
            Left            =   3750
            TabIndex        =   17
            Top             =   75
            Width           =   6840
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   2
            Left            =   10275
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
            Left            =   2025
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
            Index           =   1
            Left            =   0
            Top             =   900
            Visible         =   0   'False
            Width           =   465
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
            TabIndex        =   15
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Ubuntu Condensed"
               Size            =   12
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   4
            Left            =   0
            TabIndex        =   10
            Top             =   2100
            Width           =   10740
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Συντ. διαδρομής"
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
            Width           =   1575
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Περιγραφή διαδρομής"
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
            Width           =   1575
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
            TabIndex        =   16
            Top             =   0
            Width           =   10740
         End
      End
      Begin VB.Frame frmFrameForGridButtons 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   75
         TabIndex        =   6
         Top             =   8175
         Width           =   10890
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   465
            Index           =   6
            Left            =   300
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   75
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   820
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Δημιουργία σημείου παραλαβής"
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
            Height          =   465
            Index           =   7
            Left            =   3750
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   75
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   820
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Αντιγραφή σημείων παραλαβής"
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
            Height          =   465
            Index           =   8
            Left            =   7200
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   75
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   820
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Διαγραφή σημείων παραλαβής"
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
         Height          =   1440
         Left            =   10950
         TabIndex        =   3
         Top             =   6600
         Width           =   4515
         Begin VB.TextBox txtDestinationID 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox txtPickupRouteID 
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
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "PickupRoutes.PickupRouteID"
            Top             =   75
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   825
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   4592
            Images          =   "TablesPickupPoints.frx":0B6C
            Version         =   131072
            KeyCount        =   4
            Keys            =   "ÿÿÿ"
         End
      End
      Begin iGrid300_10Tec.iGrid grdPickupPoints 
         Height          =   6615
         Left            =   75
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1500
         Width           =   18840
         _ExtentX        =   33232
         _ExtentY        =   11668
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
         Caption         =   "Σημεία παραλαβής επιβατών"
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
         TabIndex        =   5
         Top             =   75
         Width           =   6285
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
Attribute VB_Name = "TablesPickupPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean

Private Function AddGridLine()

    With grdPickupPoints
        .AddRow
        .CellIcon(.RowCount, "Status") = lstIconList.ItemIndex(2)
        .SetCurCell .RowCount, 4
        .SetFocus
    End With

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            EnableGrid grdPickupPoints, True
            HighlightRow grdPickupPoints, 1, 1, "", False
            grdPickupPoints.SetCurCell 1, 4
            DisableFields txtRouteShortDescription, txtRouteDescription, cmdIndex(0), cmdIndex(1)
            UpdateButtons Me, 8, 0, 1, 1, 1, 1, 0, 1, 1, 1
        Else
            UpdateButtons Me, 8, 0, 1, 1, 1, 1, 0, 1, 1, 1
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
            frmCriteria(0).Visible = False
            txtRouteShortDescription.SetFocus
        End If
    End If
    
End Function

Private Function RefreshList()
    
    On Error GoTo ErrTrap
    
    'SQL
    Dim intIndex As Long
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    'Local variables
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    
    'Πλέγμα
    With grdPickupPoints
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = True
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT PickupPoints.PickupPointID, PickupPoints.PickupPointRouteID, PickupPointHotelDescription, PickupPointExactPoint, PickupPointTime " _
        & "FROM PickupPoints "
    
    'Δρομολόγιο
    If txtPickupRouteID.text <> "" Then
        strThisParameter = "intRouteID Integer"
        strThisQuery = "PickupPointRouteID = intRouteID"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtPickupRouteID.text)
    End If
    
    'Προορισμός
    If txtDestinationID.text <> "" Then
        strThisParameter = "intDestinationID Integer"
        strThisQuery = "DestinationID = intDestinationID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtDestinationID.text)
    End If
        
    'Ταξινόμηση
    strOrder = " ORDER BY PickupPointHotelDescription, PickupPointTime"
    
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
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then blnError = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 8, 0, 0, 0, 0, 1, 0, 0, 0, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True

    'Γεμίζω το πλέγμα
    With rstRecordset
        grdPickupPoints.AddRow , , , , , , , rstRecordset.RecordCount
        lngRowCount = rstRecordset.RecordCount
        Do Until .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdPickupPoints.CellValue(lngRow, "ID") = !PickupPointID
            grdPickupPoints.CellValue(lngRow, "RouteID") = !PickupPointRouteID
            grdPickupPoints.CellValue(lngRow, "HotelDescription") = !PickupPointHotelDescription
            grdPickupPoints.CellValue(lngRow, "ExactPoint") = !PickupPointExactPoint
            grdPickupPoints.CellValue(lngRow, "Time") = !PickupPointTime
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
        .Close
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdPickupPoints
        RefreshList = 0
    Else
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Τελικές ενέργειες
    cmdButton(4).Caption = "Νέα αναζήτηση"
    frmProgress.Visible = False
    
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
    blnError = True
    ClearFields grdPickupPoints, frmProgress
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function ToggleGridLineToDelete()

    With grdPickupPoints
        .CellIcon(.CurRow, "Deleted") = IIf(.CellIcon(.CurRow, "Deleted") <= 0, lstIconList.ItemIndex(3), lstIconList.ItemIndex(1))
        .SetFocus
    End With

End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Δρομολόγιο
    If Len(txtPickupRouteID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtRouteShortDescription.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)
    
    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields grdPickupPoints
            EnableFields txtRouteShortDescription, txtRouteDescription, cmdIndex(0), cmdIndex(1)
            UpdateButtons Me, 8, 1, 0, 0, 0, 0, 1, 0, 0, 0
            frmCriteria(0).Visible = True
            txtRouteShortDescription.SetFocus
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function SaveRecord()
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngID As Long
    
    If Not ValidateGrid Then Exit Function
    
    With grdPickupPoints
        For lngRow = 1 To .RowCount
            'Add Record when Status = Blue and Deleted = Blank
            If (.CellIcon(lngRow, "Status") = 1) And (.CellIcon(lngRow, "Deleted") = -1 Or .CellIcon(lngRow, "Deleted") = 0) Then
                lngID = MainSaveRecord("CommonDB", "PickupPoints", True, strApplicationName, "PickupPointID", lngID, txtPickupRouteID.text, .CellValue(lngRow, "HotelDescription"), .CellValue(lngRow, "ExactPoint"), .CellValue(lngRow, "Time"), txtPickupRouteID.text, strCurrentUser)
            End If
            'Delete Existing Record when Status = Blank and Deleted = Red
            If (.CellIcon(lngRow, "Status") = -1) And (.CellIcon(lngRow, "Deleted") = 2) Then
                lngID = MainDeleteRecord("CommonDB", "PickupPoints", strApplicationName, "PickupPointID", .CellValue(lngRow, "ID"), False)
            End If
            'Update Existing Record when Status = Blank and Deleted = Blank
            If (.CellIcon(lngRow, "Status") = -1) And (.CellIcon(lngRow, "Deleted") = -1 Or .CellIcon(lngRow, "Deleted") = 0) Then
                lngID = MainSaveRecord("CommonDB", "PickupPoints", False, strApplicationName, "PickupPointID", .CellValue(lngRow, "ID"), .CellValue(lngRow, "RouteID"), .CellValue(lngRow, "HotelDescription"), .CellValue(lngRow, "ExactPoint"), .CellValue(lngRow, "Time"), .CellValue(lngRow, "RouteID"), strCurrentUser)
            End If
        Next lngRow
    End With
    
    If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
    End If
    
    ClearFields grdPickupPoints
    EnableFields txtRouteShortDescription, txtRouteDescription, cmdIndex(0), cmdIndex(1)
    UpdateButtons Me, 8, 1, 0, 0, 0, 0, 1, 0, 0, 0
    frmCriteria(0).Visible = True
    txtRouteShortDescription.SetFocus
    
End Function

Private Function ValidateGrid()

    Dim lngRow As Long
    Dim lngCol As Long
    
    ValidateGrid = False
    
    With grdPickupPoints
        For lngRow = 1 To .RowCount
            If grdPickupPoints.CellIcon(lngRow, "Deleted") <> 2 Then
                If .CellValue(lngRow, "HotelDescription") = "" Then
                    If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
                    End If
                    .SetFocus
                    .SetCurCell lngRow, "HotelDescription"
                    Exit Function
                End If
                If .CellValue(lngRow, "Time") = "" Then
                    If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
                    End If
                    .SetFocus
                    .SetCurCell lngRow, "Time"
                    Exit Function
                End If
                If Not IsDate(.CellValue(lngRow, "Time")) Or Len(.CellValue(lngRow, "Time")) <> 5 Then
                    If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
                    End If
                    .SetFocus
                    .SetCurCell lngRow, "Time"
                    Exit Function
                End If
            End If
        Next lngRow
    End With
    
    ValidateGrid = True

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
        Case 6
            AddGridLine
        Case 7
            CreatePickupPoints
        Case 8
            ToggleGridLineToDelete
    End Select

End Sub

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Συντ. δρομολογίου
            Set tmpRecordset = CheckForMatch("CommonDB", "PickupRoutes", "PickupRouteShortDescription", "String", txtRouteShortDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 5, 60, 1, 1, 0)
                txtPickupRouteID.text = tmpTableData.strCode
                txtRouteShortDescription.text = tmpTableData.strFirstField
                txtRouteDescription.text = tmpTableData.strSecondField
            End If
        Case 1
            'Περιγραφή δρομολογίου
            Set tmpRecordset = CheckForMatch("CommonDB", "PickupRoutes", "PickupRouteDescription", "String", txtRouteDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 5, 60, 1, 1, 0)
                txtPickupRouteID.text = tmpTableData.strCode
                txtRouteShortDescription.text = tmpTableData.strFirstField
                txtRouteDescription.text = tmpTableData.strSecondField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdPickupPoints, 44, GetSetting(strApplicationName, "Layout Strings", "grdPickupPoints"), _
            "04NCIID,04NCIRouteID,04NCIDestinationID,50NLNHotelDescription,50NLNExactPoint,07NCTTime,05NCNStatus,05NCNDeleted", _
            "ID,RouteID,DestinationID,Περιγραφή,Σημείο,Ώρα,Ν,Δ"
        Me.Refresh
        frmCriteria(0).Visible = True
        txtRouteShortDescription.SetFocus
    End If
    
    'AddDummyLines grdPickupPoints, "99999", "99999", "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "A00:00A", "N", "Δ"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(6).Enabled, vbKeyN And CtrlDown = 4 And cmdButton(6).Enabled
            cmdButton_Click 6
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(8).Enabled, vbKeyD And CtrlDown = 4 And cmdButton(8).Enabled
            cmdButton_Click 8
        Case vbKeyP And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()
    
    PositionControls Me, True, grdPickupPoints
    ColorizeControls Me, True
    SetUpGrid lstIconList, grdPickupPoints
    ClearFields txtPickupRouteID, txtDestinationID
    ClearFields txtRouteShortDescription, txtRouteDescription
    EnableFields txtRouteShortDescription, txtRouteDescription
    EnableFields cmdIndex(0), cmdIndex(1)
    UpdateButtons Me, 8, 1, 0, 0, 0, 0, 1, 0, 0, 0
    
End Sub

Private Sub grdPickupPoints_ColHeaderMouseEnter(ByVal lCol As Long)

    grdPickupPoints.Header.Buttons = True

End Sub

Private Sub grdPickupPoints_ColHeaderMouseLeave(ByVal lCol As Long)

    grdPickupPoints.Header.Buttons = False
    
End Sub

Private Sub grdPickupPoints_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngColCount As Long
    Dim lngRowCount As Long
    
    lngColCount = grdPickupPoints.colCount
    lngRowCount = grdPickupPoints.RowCount
    
    If grdPickupPoints.RowCount = 0 Then Exit Sub
    
    If grdPickupPoints.CurRow = 0 Then Exit Sub
    
    grdPickupPoints.Redraw = False
    
    For lngCol = 1 To lngColCount
        For lngRow = 1 To lngRowCount
            grdPickupPoints.CellForeColor(lngRow, lngCol) = grdPickupPoints.ForeColor
            grdPickupPoints.CellBackColor(lngRow, lngCol) = grdPickupPoints.BackColor
        Next lngRow
    Next lngCol
    
    For lngCol = 1 To lngColCount
        grdPickupPoints.CellForeColor(grdPickupPoints.CurRow, lngCol) = vbWhite
        grdPickupPoints.CellBackColor(grdPickupPoints.CurRow, lngCol) = &HC0C000
    Next lngCol
    
    grdPickupPoints.Redraw = True

End Sub

Private Sub grdPickupPoints_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdPickupPoints", grdPickupPoints.LayoutCol

End Sub

Private Sub txtRouteDescription_Change()

    If txtRouteDescription.text = "" Then
        ClearFields txtPickupRouteID, txtRouteShortDescription, txtRouteDescription
    End If

End Sub

Private Sub txtRouteDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 And txtPickupRouteID.text = "" Then cmdIndex_Click 1
    
End Sub

Private Sub txtRouteDescription_Validate(Cancel As Boolean)

    If txtPickupRouteID.text = "" And txtRouteDescription.text <> "" Then cmdIndex_Click 1
    
End Sub

Private Sub txtRouteShortDescription_Change()

    If txtRouteShortDescription.text = "" Then
        ClearFields txtPickupRouteID, txtRouteDescription
    End If

End Sub

Private Sub txtRouteShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 And txtPickupRouteID.text = "" Then cmdIndex_Click 0
    
End Sub

Private Sub txtRouteShortDescription_Validate(Cancel As Boolean)

    If txtPickupRouteID.text = "" And txtRouteShortDescription.text <> "" Then cmdIndex_Click 0
    
End Sub

Private Function PrintRecords()

    If SelectPrinter("PrinterPrintsReports") Then
        CreateUnicodeFile "ΣΗΜΕΙΑ ΠΑΡΑΛΑΒΗΣ ΔΡΟΜΟΛΟΓΙΟΥ", txtRouteDescription.text, "", intPrinterReportDetailLines
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

Private Function CreateUnicodeFile(strReportTitle, strReportSubTitle1, strReportSubTitle2, intReportDetailLines)

    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Open strUnicodeFile For Output As #1
    GoSub Headers
    
    With grdPickupPoints
        
        For lngRow = 1 To .RowCount
            
            'Εκτυπώνω τη γραμμή
            Print #1, .CellText(lngRow, "HotelDescription"); Tab(52); .CellText(lngRow, "ExactPoint"); Tab(103); .CellText(lngRow, "Time")
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            
            'Eject
            If intProcessedDetailLines > Val(intReportDetailLines) Then
                intPageNo = intPageNo + 1
                Print #1, ""
                Print #1, strStandardMessages(24)
                
                GoSub Headers
                
                Print #1, strStandardMessages(25)
                Print #1, ""
                intProcessedDetailLines = intProcessedDetailLines + 2
            End If
        
        Next lngRow
        
        Print #1, ""
        Print #1, strStandardMessages(26)
    
    End With
    
    Close #1
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 107, intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
    PrintColumnHeadings 1, "ΠΕΡΙΓΡΑΦΗ", 52, "ΣΗΜΕΙΟ", 103, "ΩΡΑ"
    Print #1, ""
    intProcessedDetailLines = 6
        
    Return
    
End Function

Private Function ExportRecords()

    Dim pdf As New ARExportPDF
    
    CreateUnicodeFile "ΣΗΜΕΙΑ ΠΑΡΑΛΑΒΗΣ ΔΡΟΜΟΛΟΓΙΟΥ", txtRouteDescription.text, "", GetSetting(strApplicationName, "Settings", "Export Report Height")
    
    With rptOneLiner
        .Restart
        .Run False
        pdf.AcrobatVersion = 2
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = strReportsPathName & "ΣΗΜΕΙΑ ΠΑΡΑΛΑΒΗΣ ΔΡΟΜΟΛΟΓΙΟΥ " & txtPickupRouteID.text & ".pdf"
        pdf.Export .Pages
        If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
        End If
    End With
    
End Function

Private Function CreatePickupPoints()

    Dim lngRow As Long
    
    With grdPickupPoints
        For lngRow = 1 To .RowCount Step 1
            .AddRow
            .CellValue(.RowCount, "RouteID") = .CellValue(lngRow, "RouteID")
            .CellValue(.RowCount, "HotelDescription") = .CellValue(lngRow, "HotelDescription")
            .CellValue(.RowCount, "ExactPoint") = .CellValue(lngRow, "ExactPoint")
            .CellIcon(.RowCount, "Status") = lstIconList.ItemIndex(2)
        Next lngRow
        .SetCurCell .RowCount, 4
        .SetFocus
    End With

End Function
