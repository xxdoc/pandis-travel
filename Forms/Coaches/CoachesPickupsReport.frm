VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form CoachesReport 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   -30
   ClientTop       =   15
   ClientWidth     =   19170
   ControlBox      =   0   'False
   FillColor       =   &H00800000&
   FillStyle       =   0  'Solid
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
      Left            =   12675
      TabIndex        =   38
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "CoachesPickupsReport.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "CoachesPickupsReport.frx":001C
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
         TabIndex        =   40
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
      TabIndex        =   5
      Top             =   75
      Width           =   18990
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   75
         TabIndex        =   27
         Top             =   8850
         Width           =   8940
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   28
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
            Index           =   5
            Left            =   7350
            TabIndex        =   29
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
            TabIndex        =   30
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
            Index           =   4
            Left            =   5925
            TabIndex        =   31
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
            TabIndex        =   32
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
            Index           =   3
            Left            =   4500
            TabIndex        =   33
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
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3690
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   5025
         Width           =   7815
         Begin UserControls.newText txtCustomerDescription 
            Height          =   465
            Left            =   1800
            TabIndex        =   3
            Top             =   1875
            Width           =   5040
            _ExtentX        =   8890
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
            TabIndex        =   2
            Top             =   1350
            Width           =   5040
            _ExtentX        =   8890
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   6900
            TabIndex        =   20
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
            PicNormal       =   "CoachesPickupsReport.frx":0038
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   6900
            TabIndex        =   21
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
            PicNormal       =   "CoachesPickupsReport.frx":05D2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newDate mskFrom 
            Height          =   465
            Left            =   1800
            TabIndex        =   0
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
         Begin UserControls.newDate mskTo 
            Height          =   465
            Left            =   3375
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
         Begin UserControls.newText txtRouteDescription 
            Height          =   465
            Left            =   1800
            TabIndex        =   4
            Top             =   2400
            Width           =   5040
            _ExtentX        =   8890
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   2
            Left            =   6900
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2400
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
            PicNormal       =   "CoachesPickupsReport.frx":0B6C
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
            Left            =   3000
            Top             =   2850
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
            Left            =   3825
            Top             =   525
            Visible         =   0   'False
            Width           =   465
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
            Index           =   1
            Left            =   450
            TabIndex        =   23
            Top             =   2475
            Width           =   915
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
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
            Index           =   0
            Left            =   450
            TabIndex        =   22
            Top             =   900
            Width           =   915
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   915
            Index           =   1
            Left            =   7275
            Top             =   1500
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   0
            Left            =   1350
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
            Index           =   2
            Left            =   0
            Top             =   1050
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
            Height          =   540
            Index           =   4
            Left            =   0
            TabIndex        =   14
            Top             =   3150
            Width           =   7815
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
            TabIndex        =   13
            Top             =   75
            Width           =   4890
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
            TabIndex        =   12
            Top             =   75
            Width           =   1740
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
            Index           =   6
            Left            =   450
            TabIndex        =   11
            Top             =   1425
            Width           =   915
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Πελάτης"
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
            TabIndex        =   10
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
            Index           =   2
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   7815
         End
      End
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2565
         Left            =   8025
         TabIndex        =   6
         Top             =   6150
         Width           =   4515
         Begin VB.TextBox txtTransferDestinationID 
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
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   450
            Width           =   780
         End
         Begin VB.TextBox Text7 
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
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "Transfers.DestinationID"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox txtCustomerID 
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
            TabIndex        =   35
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
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "Transfers.CustomerID"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
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
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "Routes.RouteID"
            Top             =   825
            Width           =   3540
         End
         Begin VB.TextBox txtRouteID 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
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
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox txtCallingForm 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "Called from form"
            Top             =   1575
            Width           =   3540
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Text            =   "RefersTo"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox txtRefersTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Top             =   1200
            Width           =   780
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1950
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   4592
            Images          =   "CoachesPickupsReport.frx":1106
            Version         =   131072
            KeyCount        =   4
            Keys            =   "ÿÿÿ"
         End
      End
      Begin iGrid300_10Tec.iGrid grdCoachesReport 
         Height          =   7290
         Left            =   75
         TabIndex        =   16
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   525
         Width           =   14940
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
         TabIndex        =   42
         Top             =   1125
         Width           =   16365
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
         TabIndex        =   41
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Κατάσταση επιβατών λεωφορείων"
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
         TabIndex        =   17
         Top             =   75
         Width           =   7560
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
Attribute VB_Name = "CoachesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskFrom.text, mskTo.text, txtDestinationDescription.text, txtCustomerDescription.text, txtRouteDescription.text
            EnableGrid grdCoachesReport, False
            HighlightRow grdCoachesReport, 1, 1, "", True
            UpdateButtons Me, 5, 0, 1, 1, 1, 1, 0
            Exit Function
        Else
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
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
            mskFrom.SetFocus
        End If
    End If

End Function

Private Function RunActiveReport()

    On Error GoTo ErrTrap
    
    With rptCoachesReport
        .Caption = lblTitle.Caption
        .Restart
        If intPreviewReports = 1 Then
            .Zoom = -2
            .Printer.ColorMode = vbPRCMMonochrome
            .WindowState = vbMaximized
            .Run False
            .Show 1
        Else
            If GetSetting(appName:=strApplicationName, Section:="Settings", Key:="IsDevelopment") = "1" Then
                MsgBox "Development Mode: Will not print!", vbInformation
                Exit Function
            Else
                .Printer.DeviceName = strPrinterName
                .PrintReport False
                .Run True
            End If
        End If
    End With
    
    RunActiveReport = True
    
    Exit Function
    
ErrTrap:
    RunActiveReport = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function UpdateCriteriaLabels(DateIssueFrom, DateIssueTo, destination, Person, Route)

    Dim strCriteriaA As String

    strCriteriaA = IIf(DateIssueFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & DateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(DateIssueTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & DateIssueTo & " ] ")
    strCriteriaA = strCriteriaA & IIf(destination = "", "Προορισμός [ ΟΛΟΙ ] ", "Προορισμός [ " & destination & " ]")
    strCriteriaA = strCriteriaA & IIf(Person = "", "Πελάτης [ ΟΛΟΙ ] ", "Πελάτης [ " & Person & " ]")
    strCriteriaA = strCriteriaA & IIf(Route = "", "Δρομολόγιο [ ΟΛΑ ] ", "Δρομολόγιο [ " & Route & " ]")
    
    lblCriteria.Caption = strCriteriaA
    
End Function

Private Function DoControlBreak(gridName As iGrid, totalName, ParamArray levelName() As Variant)

    Dim lngRow As Long
    Dim intLoop As Integer
    Dim level As Integer
    Dim curRouteTotal As Currency
    Dim curDailyTotal As Currency
    Dim curGrandTotal As Currency
    ReDim oldArea(UBound(levelName)) As String
    
    gridName.Redraw = False
    
    gridName.AddRow 1
    
    For intLoop = 0 To UBound(levelName)
        oldArea(intLoop) = gridName.CellValue(1, levelName(intLoop))
    Next intLoop
    
    lngRow = 1
    level = UBound(levelName)
    
    InitializeProgressBar Me, strApplicationName, gridName.RowCount
    
    Do While True
        Do While oldArea(level) = gridName.CellValue(lngRow, levelName(level))
            If oldArea(level) = gridName.CellValue(lngRow, levelName(level)) Then
                curRouteTotal = curRouteTotal + gridName.CellValue(lngRow, totalName)
            Else
                GoSub AddTotalLineAndUpdateLevels
            End If
            lngRow = lngRow + 1
            If lngRow >= gridName.RowCount Then Exit Do
            UpdateProgressBar Me
            DoEvents
            If Not blnProcessing Then
                DoControlBreak = False
                Exit Function
            End If
        Loop
        GoSub AddTotalLineAndUpdateLevels
        If level - 1 >= 0 Then
            If gridName.CellValue(lngRow, levelName(level - 1)) <> "" Then
                If oldArea(level - 1) <> gridName.CellValue(lngRow, levelName(level - 1)) Then
                    GoSub AddTotalLineAndUpdateLevels
                End If
            End If
        End If
        If lngRow >= gridName.RowCount Then
            Exit Do
        End If
    Loop
    
    GoSub AddDailyTotal
    GoSub AddGrandTotalLine
    
    gridName.Redraw = True
    
    Exit Function
    
AddTotalLineAndUpdateLevels:
    gridName.AddRow "", lngRow, , , , , , 1
    gridName.CellValue(lngRow, 3) = "     ΣΥΝΟΛΟ: " & curRouteTotal
    curDailyTotal = curDailyTotal + curRouteTotal
    gridName.CellForeColor(lngRow, 3) = vbCyan
    oldArea(level) = gridName.CellValue(lngRow + 1, levelName(level))
    If level - 1 >= 0 Then
        If gridName.CellValue(lngRow + 1, levelName(level - 1)) <> "" Then
            If oldArea(level - 1) <> gridName.CellValue(lngRow + 1, levelName(level - 1)) Then
                GoSub AddDailyTotal
            End If
            oldArea(level - 1) = gridName.CellValue(lngRow + 1, levelName(level - 1))
        End If
    End If
    lngRow = lngRow + 1
    
    curRouteTotal = 0
    
    Return
    
AddDailyTotal:
    If lngRow >= gridName.RowCount Then
        lngRow = gridName.RowCount
    Else
        lngRow = lngRow + 1
        gridName.AddRow "", lngRow
    End If
    gridName.CellForeColor(lngRow, 3) = vbCyan
    gridName.CellValue(lngRow, 3) = "     ΣΥΝΟΛΟ ΗΜΕΡΑΣ: " & curDailyTotal
    curGrandTotal = curGrandTotal + curDailyTotal
    curDailyTotal = 0
    
    Return

AddGrandTotalLine:
    gridName.AddRow ""
    gridName.CellValue(gridName.RowCount, 3) = "     ΓΕΝΙΚΟ ΣΥΝΟΛΟ: " & curGrandTotal
    gridName.CellForeColor(gridName.RowCount, 3) = vbCyan
    
    Return
    
End Function

Private Function RefreshList()

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
    
    'Local variables
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    frmCriteria(0).Visible = False
    
    'Πλέγμα
    With grdCoachesReport
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT " _
        & "TransferID, TransferDate, TransferAdults, TransferKids, TransferFree, TransferRemarks, TransferPickupPointDescription, " _
        & "PickUpPointHotelDescription, PickUpPointExactPoint, PickUpPointTime, PickupRouteDescription, " _
        & "PickupRouteDescription, " _
        & "Description, " _
        & "DestinationDescription , DestinationShortDescription " _
        & "FROM ((((Transfers " _
        & "LEFT JOIN PickupPoints ON Transfers.TransferPickupPointID = PickupPoints.PickUpPointID) " _
        & "LEFT JOIN PickupRoutes ON Transfers.TransferRouteID = PickupRoutes.PickupRouteID) " _
        & "LEFT JOIN Customers ON Transfers.TransferCustomerID = Customers.ID) " _
        & "INNER JOIN Destinations ON Transfers.TransferDestinationID = Destinations.DestinationID) "
    
    'Από
    If IsDate(mskFrom.text) Then
        strThisParameter = "datFrom Date"
        strThisQuery = "Transfers.TransferDate >= datFrom"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskFrom.text)
    End If
    
    'Εως
    If IsDate(mskTo.text) Then
        strThisParameter = "datTo Date"
        strThisQuery = "Transfers.TransferDate <= datTo"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskTo.text)
    End If
    
    'Προορισμός
    If txtTransferDestinationID.text <> "" Then
        strThisParameter = "intDestinationID Integer"
        strThisQuery = "Transfers.TransferDestinationID = intDestinationID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtTransferDestinationID.text)
    End If
    
    'Πελάτης
    If txtCustomerID.text <> "" Then
        strThisParameter = "intID Integer"
        strThisQuery = "Transfers.TransferCustomerID = intID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtCustomerID.text)
    End If
    
    'Δρομολόγιο
    If txtRouteID.text <> "" Then
        strThisParameter = "intRouteID Integer"
        strThisQuery = "Transfers.TransferRouteID = intRouteID "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = Val(txtRouteID.text)
    End If
    
    'Κανονική καταχώρηση = 1, Γρήγορη = 2
    strThisParameter = "intShowInList Integer"
    strThisQuery = "Transfers.ShowInList = intShowInList "
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtRefersTo.text)
    
    'Ταξινόμηση
    strOrder = IIf(txtRefersTo.text = "1", " ORDER BY TransferDate, PIckupRouteDescription, PickupPointTime, PickUpPointHotelDescription, Description", "ORDER BY Description")

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
    UpdateButtons Me, 5, 0, 0, 0, 0, 1, 0
    cmdButton(4).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            grdCoachesReport.AddRow
            lngRowCount = rstRecordset.RecordCount
            UpdateProgressBar Me
            lngRow = lngRow + 1
            grdCoachesReport.CellValue(lngRow, "TransferID") = !TransferID
            grdCoachesReport.CellValue(lngRow, "TransferDate") = !transferDate
            grdCoachesReport.CellValue(lngRow, "CustomerDescription") = !Description
            grdCoachesReport.CellValue(lngRow, "DestinationDescription") = !DestinationDescription
            grdCoachesReport.CellValue(lngRow, "RouteDescription") = !PickupRouteDescription
            grdCoachesReport.CellValue(lngRow, "PickupPointHotelDescription") = !PickupPointHotelDescription
            grdCoachesReport.CellValue(lngRow, "PickUpPointExactPoint") = !PickupPointExactPoint
            grdCoachesReport.CellValue(lngRow, "PickUpPointTime") = !PickupPointTime
            grdCoachesReport.CellValue(lngRow, "TransferAdults") = IIf(!TransferAdults > 0, !TransferAdults, "")
            grdCoachesReport.CellValue(lngRow, "TransferKids") = IIf(!TransferKids > 0, !TransferKids, "")
            grdCoachesReport.CellValue(lngRow, "TransferFree") = IIf(!TransferFree > 0, !TransferFree, "")
            grdCoachesReport.CellValue(lngRow, "TransferTotal") = !TransferAdults + !TransferKids + !TransferFree
            grdCoachesReport.CellValue(lngRow, "RefNo") = CreateReferenceNo(!DestinationShortDescription, !transferDate, !TransferID)
            grdCoachesReport.CellValue(lngRow, "TransferRemarks") = !TransferRemarks
            grdCoachesReport.CellValue(lngRow, "TransferDraftDescription") = !TransferPickupPointDescription
            rstRecordset.MoveNext
            DoEvents
            If Not blnProcessing Then Exit Do
        Loop
    End With
    
    'Να συνεχίσω;
    If blnProcessing Then
        If txtRefersTo.text = "1" Then
            DoControlBreak grdCoachesReport, "TransferTotal", "TransferDate", "RouteDescription"
        Else
            DoControlBreak grdCoachesReport, "TransferTotal", "CustomerDescription"
        End If
    End If
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdCoachesReport
        RefreshList = 0
    Else
        blnProcessing = False
        RefreshList = lngRowCount
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
    ClearFields grdCoachesReport, frmProgress
    DisplayErrorMessage True, Err.Description

End Function

Private Function RemoveTotals(gridName As iGrid)

    Dim lngRow As Long
    
    gridName.Redraw = False
    
    For lngRow = 1 To gridName.RowCount
        If gridName.CellValue(lngRow, "TransferDate") = "" Then
            gridName.RemoveRow (lngRow)
            lngRow = lngRow - 1
            If lngRow = gridName.RowCount Then
                Exit For
            End If
        End If
    Next lngRow
    
    gridName.Redraw = False

End Function

Private Sub cmdButton_Click(index As Integer)
                                
    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            EditRecord
        Case 2
            DoReport "Print"
        Case 3
            DoReport "CreatePDF"
        Case 4
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select

End Sub

Private Function DoReport(action As String)
    
    If action = "Print" Then
        If SelectPrinter("PrinterPrintsReports") Then
            If PrinterExists(strPrinterName) Then
                'Κανονική καταχώρηση
                If txtRefersTo.text = "1" Then
                    RemoveTotals grdCoachesReport
                    RunActiveReport
                    If txtRefersTo.text = "1" Then
                        DoControlBreak grdCoachesReport, "TransferTotal", "TransferDate", "RouteDescription"
                    Else
                        DoControlBreak grdCoachesReport, "TransferTotal", "Description"
                    End If
                End If
                'Γρήγορη καταχώρηση
                If txtRefersTo.text = "2" Then
                    If CreateUnicodeFile("ΚΑΤΑΣΤΑΣΗ ΜΕΤΑΦΟΡΩΝ", lblCriteria.Caption, "", intPrinterReportDetailLines) Then
                        With rptOneLiner
                            If intPreviewReports = 1 Then
                                .Restart
                                .Zoom = -2
                                .WindowState = vbMaximized
                                .Show 1
                            Else
                                .Printer.DeviceName = strPrinterName
                                .PrintReport False
                                .Run True
                            End If
                        End With
                    End If
                End If
            Else
                If MyMsgBox(4, strApplicationName, strStandardMessages(18), 1) Then
                End If
            End If
        End If
    End If
    
    If action = "CreatePDF" Then
        'Κανονική καταχώρηση
        If txtRefersTo.text = "1" Then
            RemoveTotals grdCoachesReport
            If CreatePDF("ΚΑΤΑΣΤΑΣΗ ΜΕΤΑΦΟΡΩΝ " & lblCriteria.Caption) Then
                If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
                End If
            End If
            If txtRefersTo.text = "1" Then
                DoControlBreak grdCoachesReport, "TransferTotal", "TransferDate", "RouteDescription"
            Else
                DoControlBreak grdCoachesReport, "TransferTotal", "Description"
            End If
        End If
        'Γρήγορη καταχώρηση
        If txtRefersTo.text = "2" Then
            If CreateUnicodeFile("ΚΑΤΑΣΤΑΣΗ ΜΕΤΑΦΟΡΩΝ", lblCriteria.Caption, "", GetSetting(strApplicationName, "Settings", "Export Report Height")) Then
                If CreateUnisexPDF("ΚΑΤΑΣΤΑΣΗ ΜΕΤΑΦΟΡΩΝ " & lblCriteria.Caption) Then
                    If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
                    End If
                End If
            End If
        End If
    End If
    
End Function

Private Function CreatePDF(fileName)

    On Error GoTo ErrTrap
    
    Dim pdf As New ARExportPDF
    
    With rptCoachesReport
        .Restart
        .Run False
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = Replace(fileName, "/", "-")
        pdf.fileName = Replace(pdf.fileName, "[", "")
        pdf.fileName = Replace(pdf.fileName, "]", "")
        pdf.fileName = Replace(pdf.fileName, "  ", " ")
        pdf.fileName = strReportsPathName & Replace(pdf.fileName, ":", "") & ".pdf"
        pdf.Export .Pages
    End With
    
    CreatePDF = True
    
    Exit Function
    
ErrTrap:
    CreatePDF = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Σωστό διάστημα
    If IsDate(mskFrom.text) And IsDate(mskTo.text) Then
        If CDate(mskFrom.text) > CDate(mskTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskFrom.SetFocus
            Exit Function
        End If
    End If
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function

    If Not blnStatus Then
        ClearFields grdCoachesReport, lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
        frmCriteria(0).Visible = True
        mskFrom.SetFocus
        UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Sub cmdIndex_Click(index As Integer)

    'Local variables
    Dim strShowInList As String
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset

    Select Case index
        'Προορισμός
        Case 0
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtTransferDestinationID.text = tmpTableData.strCode
                txtDestinationDescription.text = tmpTableData.strFirstField
            End If
        'Πελάτης
        Case 1
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "Description", "String", txtCustomerDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtCustomerID.text = tmpTableData.strCode
                txtCustomerDescription.text = tmpTableData.strFirstField
            End If
        'Δρομολόγιο
        Case 2
            Set tmpRecordset = CheckForMatch("CommonDB", "PickupRoutes", "PickupRouteDescription", "String", txtRouteDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 60, 1, 0)
                txtRouteID.text = tmpTableData.strCode
                txtRouteDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()
    
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdCoachesReport, 44, GetSetting(strApplicationName, "Layout Strings", grdCoachesReport.Tag), _
            "05NCNTransferID,12NCDTransferDate,40NLNCustomerDescription,40NLNDestinationDescription,50NLNRouteDescription,40NLNPickupPointHotelDescription,10NLNPickUpPointExactPoint,10NCTPickupPointTime,10NRITransferAdults,10NRITransferKids,10NRITransferFree,10NRITransferTotal,10NCΝRefNo,10NLNTransferRemarks,50NLNTransferDraftDescription", _
            "TransferID,Ημερομηνία,Πελάτης,Προορισμός,Δρομολόγιο,Σημείο παραλαβής,Ακριβές σημείο,Ωρα,Ενήλικες,Παιδιά,Δωρεάν,Σύνολο,Αριθμός αναφοράς,Παρατηρήσεις,Σημείο παραλαβής"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskFrom.SetFocus
    End If
    
    'AddDummyLines grdCoachesReport, "99999", "Α99/99/9999Α", "ΠΡΟΟΡΙΣΜΟΣ", "ΠΕΛΑΤΗΣ", "ΔΡΟΜΟΛΟΓΙΟ", "ΣΗΜΕΙΟ ΠΑΡΑΛΑΒΗΣ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "Α00:00Α", "999999", "999999", "999999", "999999", "Αριθμός αναφοράς", "Παρατηρήσεις", "Σημείο παραλαβής"
        
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
        Case vbKeyE And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyP And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyP And CtrlDown = 5 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdCoachesReport
    PositionControls Me, True, grdCoachesReport
    ColorizeControls Me, True
    ClearFields lblRecordCount, lblCriteria, lblSelectedGridLines, lblSelectedGridTotals
    ClearFields txtCustomerID, txtTransferDestinationID, txtRouteID
    ClearFields mskFrom, mskTo, txtDestinationDescription, txtCustomerDescription, txtRouteDescription
    ClearFields grdCoachesReport
    EnableFields mskFrom, mskTo, txtDestinationDescription, txtCustomerDescription, txtRouteDescription
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    
End Sub

Private Function EditRecord()
    
    Dim blnFound As Boolean
    
    'Κανονικές μεταφορές
    If txtRefersTo.text = "1" Then
        blnFound = CoachesPickupsStandard.SeekRecord(grdCoachesReport.CellValue(grdCoachesReport.CurRow, "TransferID"), txtRefersTo.text)
        If blnFound Then
            If txtCallingForm.text = "Transactions" Then
                Unload Me
                Exit Function
            Else
                CoachesPickupsStandard.Show 1, Me
            End If
        End If
        grdCoachesReport.SetFocus
    End If
    
    'Γρήγορες μεταφορές
    If txtRefersTo.text = "2" Then
        blnFound = CoachesPickupsBrief.SeekRecord(grdCoachesReport.CellValue(grdCoachesReport.CurRow, "TransferID"), txtRefersTo.text)
        If blnFound Then
            If txtCallingForm.text = "Transactions" Then
                Unload Me
                Exit Function
            Else
                CoachesPickupsBrief.Show 1, Me
            End If
        End If
        grdCoachesReport.SetFocus
    End If
    
End Function

Private Sub grdCoachesReport_ColHeaderMouseEnter(ByVal lCol As Long)

    grdCoachesReport.Header.Buttons = True

End Sub

Private Sub grdCoachesReport_ColHeaderMouseLeave(ByVal lCol As Long)

    grdCoachesReport.Header.Buttons = False
    
End Sub

Private Sub grdCoachesReport_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)

    cmdButton(1).Enabled = ChangeEditButtonStatus(grdCoachesReport, Me.Tag, lRow, 1)

End Sub

Private Sub grdCoachesReport_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub grdCoachesReport_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub grdCoachesReport_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", grdCoachesReport.Tag, grdCoachesReport.LayoutCol

End Sub

Private Sub txtCustomerDescription_Change()

    If txtCustomerDescription.text = "" Then txtCustomerID.text = ""

End Sub

Private Sub txtCustomerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtCustomerDescription_Validate(Cancel As Boolean)

    If txtCustomerID.text = "" And txtCustomerDescription.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then txtTransferDestinationID.text = ""

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtTransferDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 0
    
End Sub

Private Function CreateUnicodeFile(strReportTitle, strReportSubTitle1, strReportSubTitle2, intReportDetailLines)

    On Error GoTo ErrTrap
    
    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    'Μετρητές
    Dim intAdults As Integer

    'Αρχικές τιμές
    intAdults = 0
    intPageNo = 1
    
    Open strUnicodeFile For Output As #1

    'Επικεφαλίδες
    PrintHeadings 127, intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
    PrintColumnHeadings 1, "ΑΡ.ΑΝΑΦ.", 11, "ΠΕΛΑΤΗΣ", 42, "ΣΗΜΕΙΟ ΠΑΡΑΛΑΒΗΣ", 84, "ΑΤΟΜΑ", 90, "ΠΑΡΑΤΗΡΗΣΕΙΣ"
    Print #1, ""
    Print #1, ""
    
    'Εγγραφές
    intProcessedDetailLines = 7
    
    With grdCoachesReport
        For lngRow = 1 To grdCoachesReport.RowCount
            
            'Εκτυπώνω τη γραμμή
            Print #1, .CellText(lngRow, "RefNo"); _
                Tab(11); Left(.CellText(lngRow, "CustomerDescription"), 30); _
                Tab(42); Left(.CellText(lngRow, "TransferDraftDescription"), 41); _
                Tab(89 - Len((format(.CellText(lngRow, "TransferAdults"), "#,##0")))); format(.CellText(lngRow, "TransferAdults"), "#,##0"); _
                Tab(90); Left(.CellText(lngRow, "TransferRemarks"), 38)
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            
            'Eject
            If intProcessedDetailLines > intReportDetailLines Then
                intPageNo = intPageNo + 1
                PrintHeadings 127, intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
                PrintColumnHeadings 1, "ΑΡ.ΑΝΑΦ.", 11, "ΠΕΛΑΤΗΣ", 42, "ΣΗΜΕΙΟ ΠΑΡΑΛΑΒΗΣ", 84, "ΑΤΟΜΑ", 90, "ΠΑΡΑΤΗΡΗΣΕΙΣ"
                intProcessedDetailLines = 5
            End If
            
        Next lngRow
    End With
    
    Close #1
    
    CreateUnicodeFile = True
    
    Exit Function
    
ErrTrap:
    CreateUnicodeFile = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Sub txtRouteDescription_Change()

    If txtRouteDescription.text = "" Then txtRouteID.text = ""

End Sub

Private Sub txtRouteDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtRouteDescription_Validate(Cancel As Boolean)

    If txtRouteID.text = "" And txtRouteDescription.text <> "" Then cmdIndex_Click 2

End Sub


