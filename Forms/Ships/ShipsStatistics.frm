VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form ShipsStatistics 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19170
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
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
      Left            =   12525
      TabIndex        =   32
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "ShipsStatistics.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "ShipsStatistics.frx":001C
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
         TabIndex        =   34
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
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
         TabIndex        =   24
         Top             =   8850
         Width           =   4515
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   150
            TabIndex        =   25
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3000
            TabIndex        =   26
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   1
            Left            =   1575
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            ForeColor       =   0
            PicOpacity      =   0
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3090
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   5625
         Width           =   7665
         Begin UserControls.newDate mskInvoiceDateIssueFrom 
            Height          =   465
            Left            =   1800
            TabIndex        =   1
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
         Begin UserControls.newDate mskInvoiceDateIssueTo 
            Height          =   465
            Left            =   3375
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
         Begin UserControls.newText txtShipDescription 
            Height          =   465
            Left            =   1800
            TabIndex        =   3
            Top             =   1350
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   0
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
         Begin UserControls.newText txtDestinationDescription 
            Height          =   465
            Left            =   1800
            TabIndex        =   4
            Top             =   1875
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   820
            ForeColor       =   0
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
            Left            =   6825
            TabIndex        =   22
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
            PicNormal       =   "ShipsStatistics.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   6825
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1875
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
            PicNormal       =   "ShipsStatistics.frx":05D2
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
            Left            =   2250
            Top             =   2325
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   3
            Left            =   3900
            Top             =   525
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
            Left            =   7200
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
            Left            =   1350
            Top             =   1050
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
            Top             =   975
            Visible         =   0   'False
            Width           =   465
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
            TabIndex        =   20
            Top             =   2625
            Width           =   7665
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
            Left            =   2700
            TabIndex        =   19
            Top             =   75
            Width           =   4815
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
            TabIndex        =   18
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Πλοίο"
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
            TabIndex        =   17
            Top             =   1425
            Width           =   915
         End
         Begin VB.Label lblLabel 
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
            Index           =   2
            Left            =   450
            TabIndex        =   16
            Top             =   900
            Width           =   915
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Προορισμός"
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
            TabIndex        =   15
            Top             =   1950
            Width           =   915
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
            TabIndex        =   21
            Top             =   0
            Width           =   7665
         End
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   7875
         TabIndex        =   5
         Top             =   6900
         Width           =   4515
         Begin VB.TextBox txtDestinationID 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
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
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox txtShipID 
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
            TabIndex        =   10
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
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "Ships.ShipID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "Table"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtTable 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
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
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "Destinations.DestinationID"
            Top             =   450
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1200
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   4592
            Images          =   "ShipsStatistics.frx":0B6C
            Version         =   131072
            KeyCount        =   4
            Keys            =   "ÿÿÿ"
         End
      End
      Begin iGrid300_10Tec.iGrid grdShipsStatistics 
         Height          =   7290
         Left            =   75
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1500
         Width           =   18840
         _ExtentX        =   33232
         _ExtentY        =   12859
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
      Begin VB.Label lblCriteria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Κριτήρια αναζήτησης"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2550
         TabIndex        =   31
         Top             =   1125
         Width           =   16365
      End
      Begin VB.Label lblSelectedGridLines 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Επιλεγμένες 0 εγγραφές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3975
         TabIndex        =   30
         Top             =   825
         Width           =   14940
      End
      Begin VB.Label lblSelectedGridTotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Σύνολα πάνε εδώ"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   3975
         TabIndex        =   29
         Top             =   525
         Width           =   14940
      End
      Begin VB.Label lblRecordCount 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Βρέθηκαν 99.999 εγγραφές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   315
         Left            =   75
         TabIndex        =   28
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Στατιστικά εκδρομών πλοίων"
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
         TabIndex        =   13
         Top             =   75
         Width           =   6390
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
Attribute VB_Name = "ShipsStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean
Dim blnRefreshList As Boolean

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If txtTable.text = "Sales" Then RefreshListFromSales Else RefreshListFromManifest
        If blnRefreshList Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskInvoiceDateIssueFrom.text, mskInvoiceDateIssueTo.text, txtShipDescription.text, txtDestinationDescription.text
            EnableGrid grdShipsStatistics, False
            HighlightRow grdShipsStatistics, 1, 1, "", True
            UpdateButtons Me, 2, 0, 1, 0
            Exit Function
        Else
            UpdateButtons Me, 2, 1, 0, 1
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
            mskInvoiceDateIssueFrom.SetFocus
        End If
    End If

End Function

Private Function RefreshListFromManifest()

    On Error GoTo ErrTrap

    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim rstTransactions As Recordset
    
    Dim lngRow As Long
    Dim lngTotalPersons As Long
    
    lngRowCount = 0
    frmCriteria(0).Visible = False
    blnRefreshList = False
    
    With grdShipsStatistics
        .Clear
        .Redraw = False
    End With
    
    'Στατιστικά από κατάσταση επιβαινόντων
    strSQL = "SELECT TripID, TripDate, TripLastName, ShipDescription, DestinationDescription " _
    & "FROM (((Manifest " _
    & "INNER JOIN Ships ON Manifest.TripShipID = Ships.ShipID) " _
    & "INNER JOIN Destinations ON Manifest.TripDestinationID = Destinations.DestinationID) " _
    & "INNER JOIN OccupantsDescriptions ON Manifest.TripOccupantDescriptionID = OccupantsDescriptions.OccupantDescriptionID) "
    
    'Από
    If mskInvoiceDateIssueFrom.text <> "" Then
        strThisParameter = "datFrom Date"
        strThisQuery = "Manifest.TripDate >= datFrom "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskInvoiceDateIssueFrom.text)
    End If
    
    'Εως
    If mskInvoiceDateIssueTo.text <> "" Then
        strThisParameter = "datTo Date"
        strThisQuery = "Manifest.TripDate <= datTo "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskInvoiceDateIssueTo.text)
    End If
    
    'Πλοίο
    If txtShipID.text <> "" Then
        strThisParameter = "intShip Integer"
        strThisQuery = "Manifest.TripShipID = intShip "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtShipID.text)
    End If
    
    'Προορισμός
    If txtDestinationID.text <> "" Then
        strThisParameter = "intDestination Integer"
        strThisQuery = "Manifest.TripDestinationID = intDestination "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtDestinationID.text)
    End If
    
    'Συμμετέχοντες (μόνο επιβάτες)
    strThisParameter = "lngStatistic Long"
    strThisQuery = "OccupantsDescriptions.OccupantDescriptionStatisticID = lngStatistic "
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = 1
    
    'Ταξινόμηση
    strOrder = " ORDER BY TripDate, ShipDescription"
    
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
    Set rstTransactions = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstTransactions.RecordCount = 0 Then blnErrors = False: RefreshListFromManifest = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstTransactions
    
    'Προσωρινά
    UpdateButtons Me, 2, 0, 1, 0
    cmdButton(1).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstTransactions
        Do While Not .EOF
            If Not blnProcessing Then Exit Do
            UpdateProgressBar Me
            grdShipsStatistics.AddRow
            lngRowCount = lngRowCount + 1
            lngRow = grdShipsStatistics.RowCount
            grdShipsStatistics.CellValue(lngRow, "TrnID") = !TripID
            grdShipsStatistics.CellValue(lngRow, "Date") = !TripDate
            grdShipsStatistics.CellValue(lngRow, "ShipDescription") = !ShipDescription
            grdShipsStatistics.CellValue(lngRow, "DestinationDescription") = !DestinationDescription
            grdShipsStatistics.CellValue(lngRow, "CompanyDescription") = !TripLastName
            grdShipsStatistics.CellValue(lngRow, "Persons") = 1
            lngTotalPersons = lngTotalPersons + grdShipsStatistics.CellValue(lngRow, "Persons")
            rstTransactions.MoveNext 'Επόμενη εγγραφή
            DoEvents 'Async!
        Loop
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdShipsStatistics
        RefreshListFromManifest = 0
    Else
        RefreshListFromManifest = lngRowCount
        blnRefreshList = True
        blnProcessing = False
    End If
    
    'Σύνολα
    If Not blnProcessing Then
        AddGrandTotalsToGrid 2, lngTotalPersons, 0
    End If
    
    'Τελικές ενέργειες
    cmdButton(1).Caption = "Νέα αναζήτηση"
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
    ClearFields grdShipsStatistics, frmProgress
    DisplayErrorMessage True, Err.Description

End Function

Private Function RefreshListFromSales()

    On Error GoTo ErrTrap

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
    
    Dim rstTransactions As Recordset

    Dim lngRow As Long
    Dim lngTotalPersons As Long
    Dim curTotalAmount As Currency
    
    lngRowCount = 0
    frmCriteria(0).Visible = False
    blnRefreshList = False
    
    With grdShipsStatistics
        .Clear
        .Redraw = False
    End With
    
    'Στατιστικά από πωλήσεις
    strSQL = "SELECT InvoiceTrnID, InvoiceDateIssue, InvoiceOutAdultsWithTransfer, InvoiceOutKidsWithTransfer, InvoiceOutFreeWithTransfer, InvoiceOutAdultsWithoutTransfer, InvoiceOutKidsWithoutTransfer, InvoiceOutFreeWithoutTransfer, InvoiceOutAdultsAmountWithTransfer, InvoiceOutKidsAmountWithTransfer, InvoiceOutAdultsAmountWithoutTransfer, InvoiceOutKidsAmountWithoutTransfer, InvoiceOutDirectAmount, " _
        & "ShipDescription, " _
        & "DestinationDescription, " _
        & "Description " _
        & "FROM (((((Invoices " _
        & "INNER JOIN InvoicesOut ON Invoices.InvoiceTrnID = InvoicesOut.InvoiceOutTrnID) " _
        & "INNER JOIN Customers ON Invoices.InvoicePersonID = Customers.ID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN Destinations ON InvoicesOut.InvoiceOutDestinationID = Destinations.DestinationID) " _
        & "INNER JOIN Ships ON InvoicesOut.InvoiceOutShipID = Ships.ShipID) "
    
    'Εγγραφές πωλήσεων
    strThisParameter = "strMasterRefersTo String"
    strThisQuery = "InvoiceMasterRefersTo = strMasterRefersTo "
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = "2"
    
    'Εκδρομές πλοίων ή λεωφορείων
    strThisParameter = "strSecondaryRefersTo String"
    strThisQuery = "InvoiceSecondaryRefersTo = strSecondaryRefersTo "
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = "1"
    
    'Εκδοση Από
    If mskInvoiceDateIssueFrom.text <> "" Then
        strThisParameter = "datFromDate Date"
        strThisQuery = "InvoiceDateIssue >= datFromDate "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateIssueFrom.text
    End If
    
    'Εκδοση Εως
    If mskInvoiceDateIssueTo.text <> "" Then
        strThisParameter = "datToDate Date"
        strThisQuery = "InvoiceDateIssue <= datToDate "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = mskInvoiceDateIssueTo.text
    End If
    
    'Προορισμός
    If txtDestinationID.text <> "" Then
        strThisParameter = "intDestinationID Integer"
        strThisQuery = "InvoiceOutDestinationID = intDestinationID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtDestinationID.text)
    End If
    
    'Πλοίο
    If txtShipID.text <> "" Then
        strThisParameter = "intShipID Integer"
        strThisQuery = "InvoiceOutShipID = intShipID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtShipID.text)
    End If
    
    'Ταξινόμηση
    strOrder = "ORDER BY InvoiceDateIssue, ShipDescription, DestinationDescription"
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strOrder
    Else
        TempQuery.SQL = strSQL & strOrder
    End If
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstTransactions = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstTransactions.RecordCount = 0 Then blnErrors = False: RefreshListFromSales = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstTransactions
    
    'Προσωρινά
    UpdateButtons Me, 2, 0, 1, 0
    cmdButton(1).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstTransactions
        Do While Not .EOF
            If Not blnProcessing Then Exit Do
            UpdateProgressBar Me
            grdShipsStatistics.AddRow
            lngRowCount = lngRowCount + 1
            lngRow = grdShipsStatistics.RowCount
            grdShipsStatistics.CellValue(lngRow, "TrnID") = !InvoiceTrnID
            grdShipsStatistics.CellValue(lngRow, "Date") = !InvoiceDateIssue
            grdShipsStatistics.CellValue(lngRow, "ShipDescription") = !ShipDescription
            grdShipsStatistics.CellValue(lngRow, "DestinationDescription") = !DestinationDescription
            grdShipsStatistics.CellValue(lngRow, "CompanyDescription") = !Description
            grdShipsStatistics.CellValue(lngRow, "Persons") = AddNumbers(!InvoiceOutAdultsWithTransfer, !InvoiceOutKidsWithTransfer, !InvoiceOutFreeWithTransfer, !InvoiceOutAdultsWithoutTransfer, !InvoiceOutKidsWithoutTransfer, !InvoiceOutFreeWithoutTransfer)
            grdShipsStatistics.CellValue(lngRow, "Amount") = AddNumbers(!InvoiceOutAdultsAmountWithTransfer, !InvoiceOutKidsAmountWithTransfer, !InvoiceOutAdultsAmountWithoutTransfer, !InvoiceOutKidsAmountWithoutTransfer, !InvoiceOutDirectAmount)
            lngTotalPersons = lngTotalPersons + grdShipsStatistics.CellValue(lngRow, "Persons")
            curTotalAmount = curTotalAmount + grdShipsStatistics.CellValue(lngRow, "Amount")
            rstTransactions.MoveNext 'Επόμενη εγγραφή
            DoEvents 'Async!
        Loop
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdShipsStatistics
        RefreshListFromSales = 0
    Else
        RefreshListFromSales = lngRowCount
        blnRefreshList = True
        blnProcessing = False
    End If
    
    'Σύνολα
    If Not blnProcessing Then
        AddGrandTotalsToGrid 2, lngTotalPersons, curTotalAmount
    End If
    
    'Τελικές ενέργειες
    cmdButton(1).Caption = "Νέα αναζήτηση"
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
    ClearFields grdShipsStatistics, frmProgress
    DisplayErrorMessage True, Err.Description
    
End Function


Private Function AddGrandTotalsToGrid(lngRowsToAdd, lngTotalPersons, curTotalAmount)
    
    With grdShipsStatistics
        .AddRow , , , , , , , lngRowsToAdd
        .CellValue(.RowCount, "ShipDescription") = "ΓΕΝΙΚΑ ΣΥΝΟΛΑ"
        .CellValue(.RowCount, "Persons") = lngTotalPersons
        .CellValue(.RowCount, "Amount") = curTotalAmount
    End With

End Function


Private Function UpdateCriteriaLabels(fromDate, toDate, ship, destination)

    Dim strCriteriaA As String

    strCriteriaA = "Από [ " & IIf(fromDate <> "", fromDate, "ΟΛΑ") & " ] Εως [ " & IIf(toDate <> "", toDate, "ΟΛΑ") & " ] "
    strCriteriaA = strCriteriaA & "Πλοίο [ " & IIf(ship <> "", ship, " ΟΛΑ ") & " ] "
    strCriteriaA = strCriteriaA & "Προορισμός [ " & IIf(destination <> "", destination, " ΟΛΟΙ ") & " ]"

    lblCriteria.Caption = strCriteriaA
    
End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            AbortProcedure False
        Case 2
            AbortProcedure True
    End Select
    
End Sub

Private Function ValidateFields()

    ValidateFields = False
    
    'Σωστό διάστημα
    If IsDate(mskInvoiceDateIssueFrom.text) And IsDate(mskInvoiceDateIssueTo.text) Then
        If CDate(mskInvoiceDateIssueFrom.text) > CDate(mskInvoiceDateIssueTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateIssueFrom.SetFocus
            Exit Function
        End If
    End If
    
    ValidateFields = True
    
End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
        ClearFields grdShipsStatistics
        frmCriteria(0).Visible = True
        mskInvoiceDateIssueFrom.SetFocus
        UpdateButtons Me, 2, 1, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Sub cmdIndex_Click(index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Πλοίο - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipDescription", "String", txtShipDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtShipID.text = tmpTableData.strCode
                txtShipDescription.text = tmpTableData.strFirstField
            End If
        Case 1
            'Προορισμός - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()
        
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdShipsStatistics, 44, GetSetting(strApplicationName, "Layout Strings", "grdShipsStatisticsFrom" & txtTable.text), _
            "05NCNTrnID,10NCDDate,40NLNShipDescription,40NLNDestinationDescription,40NLNCompanyDescription,10NRIXPersons,10NRFXAmount,04NCNSelected", _
            "TrnID,Ημερομηνία,Πλοίο,Προορισμός,Πελάτης,Σύνολο ατόμων,Σύνολο χρέωσης,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskInvoiceDateIssueFrom.SetFocus
    End If
            
    'AddDummyLines grdShipsStatistics, "99999", "A99/99/9999A", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "999999999", "9999999"
            
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
        Case vbKeyEscape
            If cmdButton(1).Enabled Then cmdButton_Click 1: Exit Function
            If cmdButton(2).Enabled Then cmdButton_Click 2
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdShipsStatistics
    PositionControls Me, True, grdShipsStatistics
    ColorizeControls Me, True
    
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    ClearFields txtShipID, txtDestinationID
    ClearFields mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, txtShipDescription, txtDestinationDescription
    ClearFields grdShipsStatistics
    
    EnableFields mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, txtShipDescription, txtDestinationDescription
    
    UpdateButtons Me, 2, 1, 0, 1

End Sub

Private Sub grdShipsStatistics_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdShipsStatistics_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdShipsStatistics_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdShipsStatistics.RowCount > 0 Then
        grdShipsStatistics.CellIcon(grdShipsStatistics.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdShipsStatistics, KeyCode, grdShipsStatistics.CurRow, "TrnID"))
        lblSelectedGridLines.Caption = CountSelected(grdShipsStatistics)
        If txtTable.text = "Sales" Then lblSelectedGridTotals.Caption = SumSelectedGridRows(grdShipsStatistics, False, "Persons", "Amount")
        If txtTable.text = "Manifest" Then lblSelectedGridTotals.Caption = SumSelectedGridRows(grdShipsStatistics, False, "Persons")
    End If

End Sub


Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdShipsStatisticsFrom" & txtTable.text, grdShipsStatistics.LayoutCol

End Sub

Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then
        ClearFields txtDestinationID
    End If

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtShipDescription_Change()

    If txtShipDescription.text = "" Then
        ClearFields txtShipID
    End If

End Sub

Private Sub txtShipDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub


Private Sub txtShipDescription_Validate(Cancel As Boolean)

    If txtShipID.text = "" And txtShipDescription.text <> "" Then cmdIndex_Click 0

End Sub


