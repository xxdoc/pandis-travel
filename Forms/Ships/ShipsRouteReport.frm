VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form ShipsRouteReport 
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
      Left            =   12450
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
         Picture         =   "ShipsRouteReport.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "ShipsRouteReport.frx":001C
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
         TabIndex        =   39
         Top             =   8850
         Width           =   14640
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   40
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
            Index           =   9
            Left            =   13050
            TabIndex        =   41
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
            Left            =   1650
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Επεξεργασία εγγραφής"
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
            Index           =   8
            Left            =   11625
            TabIndex        =   43
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
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3080
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   3
            Left            =   4500
            TabIndex        =   45
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   4
            Left            =   5925
            TabIndex        =   46
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   5
            Left            =   7350
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Εισαγωγή πληρώματος"
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
            Index           =   6
            Left            =   8775
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Εισαγωγή εγγραφών"
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
            Index           =   7
            Left            =   10200
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Εξαγωγή εγγραφών"
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
         Height          =   3315
         Left            =   7800
         TabIndex        =   13
         Top             =   5400
         Width           =   4515
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "Routes.RouteDepartureTime"
            Top             =   2325
            Width           =   3540
         End
         Begin VB.TextBox txtTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2325
            Width           =   780
         End
         Begin VB.TextBox txtRouteID 
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
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox txtTripID 
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
            TabIndex        =   24
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
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "Manifest.TripID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox Text3 
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
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "Manifest.TripRouteID"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Manifest.TripShipID"
            Top             =   825
            Width           =   3540
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
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox txtFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1200
            Width           =   780
         End
         Begin VB.TextBox txtVia 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox txtTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1950
            Width           =   780
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "Routes.RouteTo"
            Top             =   1950
            Width           =   3540
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Routes.RouteVia"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "Routes.RouteFrom"
            Top             =   1200
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   2700
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   4592
            Images          =   "ShipsRouteReport.frx":0038
            Version         =   131072
            KeyCount        =   4
            Keys            =   "ÿÿÿ"
         End
         Begin MSComDlg.CommonDialog OpenFileDialog 
            Left            =   675
            Top             =   2700
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3090
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   5625
         Width           =   7590
         Begin UserControls.newText txtShip 
            Height          =   465
            Left            =   1725
            TabIndex        =   2
            Top             =   1350
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
         Begin UserControls.newDate mskDate 
            Height          =   465
            Left            =   1725
            TabIndex        =   1
            Top             =   825
            Width           =   1455
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
         Begin UserControls.newText txtRoute 
            Height          =   465
            Left            =   1725
            TabIndex        =   3
            Top             =   1875
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   820
            Alignment       =   2
            ForeColor       =   0
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   6750
            TabIndex        =   30
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
            PicNormal       =   "ShipsRouteReport.frx":1248
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   2550
            TabIndex        =   31
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
            PicNormal       =   "ShipsRouteReport.frx":17E2
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
            Left            =   1950
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
            Left            =   2175
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
            Index           =   1
            Left            =   7125
            Top             =   1200
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
            Left            =   1275
            Top             =   975
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
            TabIndex        =   11
            Top             =   2625
            Width           =   7590
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
            Left            =   3225
            TabIndex        =   10
            Top             =   75
            Width           =   4215
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
            TabIndex        =   9
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Δρομολόγιο"
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
            Top             =   1950
            Width           =   840
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
            Index           =   2
            Left            =   450
            TabIndex        =   7
            Top             =   1425
            Width           =   840
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Ημερομηνία"
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
            TabIndex        =   6
            Top             =   900
            Width           =   840
         End
         Begin VB.Label lblRouteDescription 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Δρομολόγιο"
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
            Left            =   3000
            TabIndex        =   5
            Top             =   1950
            Width           =   840
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
            TabIndex        =   12
            Top             =   0
            Width           =   7590
         End
      End
      Begin iGrid300_10Tec.iGrid grdShipsRouteReport 
         Height          =   7290
         Left            =   75
         TabIndex        =   26
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
         TabIndex        =   38
         Top             =   1125
         Width           =   2565
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
         TabIndex        =   37
         Top             =   1125
         Width           =   16365
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
         TabIndex        =   36
         Top             =   525
         Width           =   14940
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
         TabIndex        =   35
         Top             =   825
         Width           =   14940
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Κατάσταση επιβατών πλοίων"
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
         TabIndex        =   27
         Top             =   75
         Width           =   6360
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
Attribute VB_Name = "ShipsRouteReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnErrorsFound As Boolean

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean



Private Function CheckForCorrectField(inputField, expectedField, errorMessage)

    If Trim(inputField) = Trim(expectedField) Then
        blnErrorsFound = False
    Else
        blnErrorsFound = True
        DisplayErrorMessage False, errorMessage
    End If

End Function

Private Function CreateFile()

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    
    Open strReportsPathName & "ΚΑΤΑΣΤΑΣΗ ΕΠΙΒΑΤΩΝ " + Right(mskDate.text, 4) & "-" + Mid(mskDate.text, 4, 2) & "-" + Left(mskDate.text, 2) & " " & txtShip.text + ".txt" For Output As #2
    
    With grdShipsRouteReport
        For lngRow = 1 To .RowCount
            Print #2, "" _
                & .CellText(lngRow, "Date") & " " _
                & .CellText(lngRow, "RouteID") & " " & Space(5 - Len(.CellText(lngRow, "RouteID"))) _
                & .CellText(lngRow, "DestinationID") & " " & Space(5 - Len(.CellText(lngRow, "DestinationID"))) _
                & .CellText(lngRow, "ShipID") & " " & Space(5 - Len(.CellText(lngRow, "ShipID"))) _
                & .CellText(lngRow, "OccupantDescriptionID") & " " & Space(5 - Len(.CellText(lngRow, "OccupantDescriptionID"))) _
                & .CellText(lngRow, "LastName") & " " & Space(40 - Len(.CellText(lngRow, "LastName"))) _
                & .CellText(lngRow, "FirstName") & " " & Space(40 - Len(.CellText(lngRow, "FirstName"))) _
                & .CellText(lngRow, "GenderID") & " " & Space(5 - Len(.CellText(lngRow, "GenderID"))) _
                & .CellText(lngRow, "AgeID") & " " & Space(5 - Len(.CellText(lngRow, "AgeID"))) _
                & .CellText(lngRow, "Care") & " " & Space(40 - Len(.CellText(lngRow, "Care"))) _
                & .CellText(lngRow, "Remarks") & " " & Space(40 - Len(.CellText(lngRow, "Remarks"))) _
                & .CellText(lngRow, "ShowInList") & " " & Space(5 - Len(.CellText(lngRow, "ShowInList"))) _
                & .CellText(lngRow, "User") & " " & Space(40 - Len(.CellText(lngRow, "User")))
        Next lngRow
    End With
    
    Close #2
    
    CreateFile = True
    
    Exit Function
    
ErrTrap:
    CreateFile = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function CreatePDF()

    On Error GoTo ErrTrap
    
    Dim pdf As New ARExportPDF
    
    With rptShipsRouteReport
        .Restart
        .Run False
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = strReportsPathName & "ΚΑΤΑΣΤΑΣΗ ΕΠΙΒΑΤΩΝ " + Right(mskDate.text, 4) & "-" + Mid(mskDate.text, 4, 2) & "-" + Left(mskDate.text, 2) & " " & txtShip.text & ".pdf"
        pdf.Export .Pages
    End With
    
    CreatePDF = True
    
    Exit Function
    
ErrTrap:
    CreatePDF = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function DeleteRecord()

    'Local variables
    Dim lngRow As Long
    Dim blnErrors As Boolean
    
    'Ερώτηση για διαγραφή
    If Not MyMsgBox(3, strAppTitle, strAppMessages(10), 2) Then
        grdShipsRouteReport.SetFocus
        Exit Function
    End If
    
    'Αρχείο λαθών
    Open strReportsPathName & "Errors.txt" For Append As #1

    'Διαγραφή επιλεγμένων εγγραφών
    With grdShipsRouteReport
        For lngRow = 1 To .RowCount
            If .CellIcon(lngRow, "Selected") = 3 Then
                If Not MainDeleteRecord("CommonDB", "Manifest", strAppTitle, "ID", .CellValue(lngRow, "ID"), "False") Then
                    Print #1, .CellValue(lngRow, "ID"); " " & .CellValue(lngRow, "Date") & " " & .CellValue(lngRow, "LastName") & " " & .CellValue(lngRow, "FirstName")
                    blnErrors = True
                End If
            End If
        Next lngRow
    End With

    'Αρχείο λαθών
    Close #1
    
    'Ελεγχος
    If blnErrors Then
        If MyMsgBox(4, strAppTitle, strAppMessages(6), 1) Then
        End If
        grdShipsRouteReport.SetFocus
        Exit Function
    End If

    'Μήνυμα ολοκλήρωσης
    If MyMsgBox(1, strAppTitle, strStandardMessages(8), 1) Then
    End If
    
    'Ανανεώνω τη λίστα
    cmdButton_Click 0

End Function

Private Function SeekRecord()

    With ShipsTransactions
        .Tag = "False"
        If .SeekRecord("Manifest", Val(grdShipsRouteReport.CellValue(grdShipsRouteReport.CurRow, "ID"))) Then
            If .Visible Then
                Unload Me
            Else
                .Show 1, Me
            End If
        End If
    End With

End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            SeekRecord
        Case 2
            DeleteRecord
        Case 3
            DoReport "Print"
        Case 4
            DoReport "CreatePDF"
        Case 5
            ImportCrew txtShipID.text
        Case 6
            ImportRecords
        Case 7
            DoReport "CreateFile"
        Case 8
            AbortProcedure False
        Case 9
            AbortProcedure True
    End Select
    
End Sub

Private Function DoReport(action As String)
    
    If action = "Print" Then
        If SelectPrinter("PrinterPrintsReports") Then
            If PrinterExists(strPrinterName) Then
                RunActiveReport
            Else
                If MyMsgBox(4, strAppTitle, strStandardMessages(18), 1) Then
                End If
                Exit Function
            End If
        End If
    End If
    
    If action = "CreatePDF" Then
        If CreatePDF Then
            If MyMsgBox(1, strAppTitle, strStandardMessages(8), 1) Then
            End If
        End If
    End If
    
    If action = "CreateFile" Then
        If CreateFile Then
            If MyMsgBox(1, strAppTitle, strStandardMessages(8), 1) Then
            End If
        End If
    End If
    
End Function

Private Function RunActiveReport()

    On Error GoTo ErrTrap
    
    With rptShipsRouteReport
        .Caption = lblCriteria.Caption
        .Restart
        If intPreviewReports = 1 Then
            .Zoom = -2
            .Printer.ColorMode = vbPRCMMonochrome
            .WindowState = vbMaximized
            .Run False
            .Show 1
        Else
            .Printer.DeviceName = strPrinterName
            .PrintReport False
            .Run True
        End If
    End With
    
    RunActiveReport = True
    
    Exit Function
    
ErrTrap:
    RunActiveReport = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function ValidateFields()

    'Αρχικές τιμές
    ValidateFields = False
    
    'Ημερομηνία
    If Not CheckDate(mskDate.text, strAppTitle) Then
        mskDate.SetFocus
        Exit Function
    End If
    
    'Πλοίο
    If txtShipID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtShip.SetFocus
        Exit Function
    End If
    
    'Δρομολόγιο
    If txtRouteID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtRoute.SetFocus
        Exit Function
    End If
    
    'Τελικές τιμές
    ValidateFields = True
    
End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function
    
    If Not blnStatus Then
        ClearFields grdShipsRouteReport, lblCriteria
        frmCriteria(0).Visible = True
        mskDate.SetFocus
        UpdateButtons Me, 9, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskDate.text, txtShip.text, lblRouteDescription.Caption
            EnableGrid grdShipsRouteReport, False
            HighlightRow grdShipsRouteReport, 1, 1, "", True
            UpdateButtons Me, 9, 0, 1, 0, 1, 1, 1, 1, 1, 1, 0
            Exit Function
        Else
            UpdateButtons Me, 9, 1, 0, 0, 0, 0, 0, 1, 0, 0, 1
            If Not blnError Then
                If blnProcessing Then
                    If MyMsgBox(4, strAppTitle, strStandardMessages(27), 1) Then
                    End If
                Else
                    If MyMsgBox(1, strAppTitle, strStandardMessages(7), 1) Then
                    End If
                End If
            End If
            blnError = False
            blnProcessing = False
            frmCriteria(0).Visible = True
            mskDate.SetFocus
        End If
    End If
    
    'If ValidateFields Then
    '    If RefreshList Then
    '        UpdateCriteriaLabels mskDate.text, txtShip.text, lblRouteDescription.Caption
    '        EnableGrid grdShipsRouteReport, False
    '        HighlightRow grdShipsRouteReport, 0, 1, "", True
    '        UpdateButtons Me, 9, 0, 1, 0, 1, 1, 1, 1, 1, 1, 0
    '        Exit Function
    '    Else
    '        If Not blnErrors Then DisplayMessageRecordsNotFound
    '        UpdateButtons Me, 9, 1, 0, 0, 0, 0, 0, 1, 0, 0, 1
    '        frmCriteria(0).Visible = True
    '        mskDate.SetFocus
    '    End If
    'End If

End Function

Private Function UpdateCriteriaLabels(myDate, myShip, myRoute)

    Dim strCriteriaA As String
    Dim strCriteriaB As String

    strCriteriaA = "Ημερομηνία [ " & IIf(myDate <> "", myDate, "ΟΛΑ") & " ] Πλοίο [ " & IIf(myShip <> "", myShip, "ΟΛΑ") & " ] "
    strCriteriaB = "Δρομολόγιο [ " & IIf(myRoute <> "", myRoute, "ΟΛΑ") & " ]"
    
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
End Function

Private Function RefreshList()

    'On Error GoTo ErrTrap
    
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
    
    'Local variables
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset

    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    frmCriteria(0).Visible = False

    'Πλέγμα
    With grdShipsRouteReport
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT TripID, TripDate, TripLastName, TripFirstName, TripCare, TripRemarks, TripRouteID, TripDestinationID, TripShipID, TripOccupantDescriptionID, TripGenderID, TripAgeID, Manifest.ShowInList, Manifest.User, OccupantDescriptionDescription, GenderDescription, AgeDescription " _
        & "FROM (((Manifest " _
        & "INNER JOIN Genders ON Manifest.TripGenderID = Genders.GenderID) " _
        & "INNER JOIN OccupantsDescriptions ON Manifest.TripOccupantDescriptionID = OccupantsDescriptions.OccupantDescriptionID) " _
        & "INNER JOIN Ages ON Manifest.TripAgeID = Ages.AgeID) "
    
    'Ημέρα
    If mskDate.text <> "" Then
        strThisParameter = "datDate Date"
        strThisQuery = "Manifest!TripDate = datDate "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskDate.text)
    End If
    
    'Πλοίο
    If txtShipID.text <> "" Then
        strThisParameter = "intShip Integer"
        strThisQuery = "Manifest!TripShipID = intShip "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtShipID.text)
    End If
    
    'Δρομολόγιο
    If txtRouteID.text <> "" Then
        strThisParameter = "intRoute Integer"
        strThisQuery = "Manifest!TripRouteID= intRoute "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtRouteID.text)
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY OccupantDescriptionDescription DESC, TripLastName, TripFirstName"
    
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
    If rstRecordset.RecordCount = 0 Then blnErrors = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, rstRecordset
    
    'Προσωρινά
    UpdateButtons Me, 9, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0
    cmdButton(8).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        grdShipsRouteReport.AddRow , , , , , , , rstRecordset.RecordCount
        lngRowCount = rstRecordset.RecordCount
        Do While Not .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdShipsRouteReport.CellValue(lngRow, "AA") = lngRow
            grdShipsRouteReport.CellValue(lngRow, "ID") = !TripID
            grdShipsRouteReport.CellValue(lngRow, "Date") = !TripDate
            grdShipsRouteReport.CellValue(lngRow, "LastName") = !TripLastName
            grdShipsRouteReport.CellValue(lngRow, "FirstName") = !TripFirstName
            grdShipsRouteReport.CellValue(lngRow, "Remarks") = !TripRemarks
            grdShipsRouteReport.CellValue(lngRow, "Care") = !TripCare
            grdShipsRouteReport.CellValue(lngRow, "OccupantDescription") = !OccupantDescriptionDescription
            grdShipsRouteReport.CellValue(lngRow, "Gender") = !GenderDescription
            grdShipsRouteReport.CellValue(lngRow, "Age") = !AgeDescription
            grdShipsRouteReport.CellValue(lngRow, "RouteID") = !TripRouteID
            grdShipsRouteReport.CellValue(lngRow, "DestinationID") = !TripDestinationID
            grdShipsRouteReport.CellValue(lngRow, "ShipID") = !TripShipID
            grdShipsRouteReport.CellValue(lngRow, "OccupantDescriptionID") = !TripOccupantDescriptionID
            grdShipsRouteReport.CellValue(lngRow, "GenderID") = !TripGenderID
            grdShipsRouteReport.CellValue(lngRow, "AgeID") = !TripAgeID
            grdShipsRouteReport.CellValue(lngRow, "ShowInList") = !ShowInList
            grdShipsRouteReport.CellValue(lngRow, "User") = !User
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
    End With
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdShipsRouteReport
        RefreshList = 0
    Else
        blnProcessing = False
        RefreshList = lngRowCount
    End If
    
    'Τελικές ενέργειες
    cmdButton(8).Caption = "Νέα αναζήτηση"
    frmProgress.Visible = False
    
    Exit Function
    
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
    blnErrors = True
    ClearFields grdShipsRouteReport, frmProgress
    DisplayErrorMessage True, Err.Description

End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Πλοίο - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipDescription", "String", txtShip.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 6, 0, 1, 3, 4, 5, 6, "ID", "Περιγραφή", "Σημαία", "Αρ. Νηολογίου", "Αρ. Ι.Μ.Ο.", "Διαχειριστής", 0, 40, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0)
            txtShipID.text = tmpTableData.strCode
            txtShip.text = tmpTableData.strFirstField
        Case 1
            'Δρομολόγιο πλοίου - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Routes", "RouteDescription", "String", txtRoute.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 6, 0, 1, 2, 3, 4, 5, "ID", "Συντ.", "Λιμένας εκκίνησης", "Ενδιάμεσοι λιμένες προσέγγισης", "Λιμένας τελικού προορισμού", "Ωρα", 0, 4, 40, 0, 40, 0, 1, 1, 0, 0, 0, 1)
            txtRouteID.text = tmpTableData.strCode
            txtRoute.text = tmpTableData.strFirstField
            lblRouteDescription.Caption = tmpTableData.strSecondField & " - " & tmpTableData.strFourthField
            txtFrom.text = tmpTableData.strSecondField
            txtVia.text = tmpTableData.strThirdField
            txtTo.text = tmpTableData.strFourthField
            txtTime.text = tmpTableData.strFifthField
    End Select

End Sub

Private Sub Form_Activate()
        
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdShipsRouteReport, 44, GetSetting(strAppTitle, "Layout Strings", "grdShipsRouteReport"), _
            "05NRNID,05NRNAA,10NCFDate,40NLNLastName,10NLNFirstName,40NLNRemarks,40NLNCare,40NLNOccupantDescription,40NLNGender,40NLNAge,05NLNRouteID,05NLNDestinationID,05NLNShipID,05NLNOccupantDescriptionID,05NLNGenderID,05NLNAgeID,05NLNShowInList,05NLNUser,05NCNSelected", _
            "ID,Α/Α,Ημερομηνία,Επώνυμο,Ονομα,Παρατηρήσεις,Ειδική φροντίδα,Ιδιότητα,Φύλο,Ηλικία,RouteID,DestinationID,ShipID,OccupantDescriptionID,GenderID,AgeID,ShowInList,User,Ε"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskDate.SetFocus
    End If
    
    'AddDummyLines grdShipsRouteReport, "99999", "999999", "A99/99/9999Α", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "AAAAAAAAAA", "1AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "2AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "ΙΔΙΟΤΗΤΑΑΑΑΑΑΑΑΑΑΑΑΑ", "ΦΥΛΟΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "ΗΛΙΚΙΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)

    Dim CtrlDown
    
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyE And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyP And CtrlDown And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyEscape
            If cmdButton(8).Enabled Then cmdButton_Click 8: Exit Function
            If cmdButton(9).Enabled Then cmdButton_Click 9
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdShipsRouteReport
    PositionControls Me, True, grdShipsRouteReport
    ColorizeControls Me, True
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    ClearFields txtTripID, txtRouteID, txtShipID
    ClearFields mskDate, txtShip, txtRoute
    ClearFields grdShipsRouteReport
    EnableFields mskDate, txtShip, txtRoute
    UpdateButtons Me, 9, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1
    
End Sub

Private Sub grdShipsRouteReport_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdShipsRouteReport_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    
    On Error GoTo ErrTrap
    
    'Local variables
    Dim txtCode As String
    
    txtCode = grdShipsRouteReport.CellValue(grdShipsRouteReport.CurRow, 1)
    
    If txtCode = "" Then
        cmdButton(1).Enabled = False
    Else
        cmdButton(1).Enabled = True
    End If

    Exit Sub
    
ErrTrap:
    Select Case Err.Number
        Case -2147220991
        Resume Next
    End Select

End Sub

Private Sub grdShipsRouteReport_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub grdShipsRouteReport_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdShipsRouteReport_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    Dim CtrlDown
    Dim lngRow As Long
    
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    'Επιλογή γραμμής
    If KeyCode = vbKeySpace Then
        With grdShipsRouteReport
            .CellIcon(grdShipsRouteReport.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdShipsRouteReport, KeyCode, .CurRow, 1))
            cmdButton(2).Enabled = False
            For lngRow = 1 To .RowCount
                If .CellIcon(lngRow, "Selected") >= 1 Then
                    cmdButton(2).Enabled = True
                    Exit Sub
                End If
            Next lngRow
        End With
    End If
    
    'Επιλογή όλων των γραμμών
    If grdShipsRouteReport.RowCount > 0 Then
        If KeyCode = vbKeyAdd And CtrlDown Then
            For lngRow = 1 To grdShipsRouteReport.RowCount
                grdShipsRouteReport.CellIcon(lngRow, "Selected") = 3
            Next lngRow
            cmdButton(2).Enabled = True
        End If
    End If
    
    'Αποεπιλογή όλων των γραμμών
    If grdShipsRouteReport.RowCount > 0 Then
        If KeyCode = vbKeySubtract And CtrlDown Then
            For lngRow = 1 To grdShipsRouteReport.RowCount
                grdShipsRouteReport.CellIcon(lngRow, "Selected") = 0
            Next lngRow
            cmdButton(2).Enabled = False
        End If
    End If
    
End Sub

Private Sub grdShipsRouteReport_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdShipsRouteReport", grdShipsRouteReport.LayoutCol

End Sub

Private Sub txtRoute_Change()

    If txtRoute.text = "" Then
        ClearFields txtRouteID, lblRouteDescription, txtFrom, txtVia, txtTo, txtTime
    End If

End Sub

Private Sub txtRoute_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtRoute_Validate(Cancel As Boolean)

    If txtRouteID.text = "" And txtRoute.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtShip_Change()

    If txtShip.text = "" Then
        ClearFields txtShipID
    End If
    
End Sub

Private Sub txtShip_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtShip_Validate(Cancel As Boolean)

    If txtShipID = "" And txtShip.text <> "" Then cmdIndex_Click 0
    
End Sub

Private Function ImportRecords()
    
    Dim strFileName As String
    Dim strLineWithError As String
    
    strFileName = ShowOpenFileDialog
    
    If strFileName <> "" Then
        strLineWithError = CheckFileForErrors(strFileName)
        If strLineWithError <> "" Then
            If MyMsgBox(4, strAppTitle, strStandardMessages(13), 1) Then
            End If
        Else
            strLineWithError = AppendRecords(strFileName)
            If strLineWithError <> "0" Then
                If MyMsgBox(4, strAppTitle, strStandardMessages(13), 1) Then
                End If
            Else
                If MyMsgBox(1, strAppTitle, strStandardMessages(8), 1) Then
                End If
            End If
        End If
    End If
    
End Function

Private Function ImportCrew(shipID As Integer)

    Dim strSQL As String
    Dim tmpRecordset As Recordset
    Dim strResult As String
    
    If MyMsgBox(2, strAppTitle, strAppMessages(12), 2) Then
        
        strSQL = "SELECT * FROM ShipsCrew WHERE CrewShipID = " & shipID
        Set tmpRecordset = CommonDB.OpenRecordset(strSQL, dbOpenSnapshot)
        
        With tmpRecordset
            If .RecordCount > 0 Then
                While Not .EOF
                    strResult = MainSaveRecord("CommonDB", "Manifest", True, strAppTitle, "ID", "", mskDate.text, Val(txtRouteID.text), grdShipsRouteReport.CellValue(1, "DestinationID"), Val(txtShipID.text), !CrewPropertyID, !CrewLastName, !CrewFirstName, !CrewGenderID, !CrewAgeID, "", "", "1", strCurrentUser)
                    .MoveNext
                Wend
            End If
            .Close
        End With
    
        If MyMsgBox(1, strAppTitle, strStandardMessages(8), 1) Then
        End If
        
        cmdButton_Click 0
    
    End If
    
End Function

Private Function ShowOpenFileDialog()

    On Error GoTo ErrTrap

    With OpenFileDialog
        .DefaultExt = "txt"
        .Filter = "Text Files (*.txt)|*.txt"
        .InitDir = strReportsPathName
        .ShowOpen
        .CancelError = True
        ShowOpenFileDialog = .fileName
        Exit Function
    End With

ErrTrap:
    ShowOpenFileDialog = ""
    Exit Function

End Function

Private Function CheckFileForErrors(fileName)

    Dim strLine As String
    Dim lngCurrentRecord As Long
    Dim lngCorrectLines As Long
    
    Open fileName For Input As #2
    lngCurrentRecord = 0
    lngCorrectLines = 0
    
    blnErrorsFound = False
    
    Do Until EOF(2)
        Line Input #2, strLine
        lngCurrentRecord = lngCurrentRecord + 1
        CheckForCorrectField Len(strLine), 258, "Η γραμμή " & lngCurrentRecord & " δεν έχει το σωστό μήκος"
        If Not blnErrorsFound Then CheckForCorrectField Mid(strLine, 1, 10), mskDate.text, "Η ημερομηνία δεν είναι σωστή"
        If Not blnErrorsFound Then CheckForCorrectField Mid(strLine, 12, 5), txtRouteID.text, "Το δρομολόγιο δεν είναι σωστό"
        If Not blnErrorsFound Then CheckForCorrectField Mid(strLine, 24, 5), txtShipID.text, "Το πλοίο δεν είναι σωστό"
        If Not blnErrorsFound Then
            lngCorrectLines = lngCorrectLines + 1
        Else
            CheckFileForErrors = True
            Exit Do
        End If
    Loop
    
    Close #2
    
End Function

Private Function AppendRecords(fileName)

    Dim strLine As String
    Dim lngCurrentRecord As Long
    Dim lngID As Long
    
    Open fileName For Input As #1
    AppendRecords = 0
    
    BeginTrans
    
    Do Until EOF(1)
        Line Input #1, strLine
        lngCurrentRecord = lngCurrentRecord + 1
        lngID = MainSaveRecord("CommonDB", "Manifest", True, strAppTitle, "ID", _
            "", _
            Mid(strLine, 1, 10), _
            Val(Mid(strLine, 12, 5)), _
            Val(Mid(strLine, 18, 5)), _
            Val(Mid(strLine, 24, 5)), _
            Val(Mid(strLine, 30, 5)), _
            Trim(Mid(strLine, 36, 40)), _
            Trim(Mid(strLine, 77, 40)), _
            Val(Mid(strLine, 118, 5)), _
            Val(Mid(strLine, 124, 5)), _
            Trim(Mid(strLine, 130, 40)), _
            Trim(Mid(strLine, 171, 40)), _
            Val(Mid(strLine, 212, 5)), _
            Trim(Mid(strLine, 218, 40)))
        If lngID = 0 Then
            AppendRecords = lngCurrentRecord
            Close #1
            Rollback
            Exit Do
        End If
    Loop
    
    If AppendRecords = 0 Then CommitTrans
    
    Close #1
    
End Function
