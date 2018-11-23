VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form PersonLedger 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   9750
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15540
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProgress 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   990
      Left            =   4950
      TabIndex        =   20
      Top             =   4650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   465
         Left            =   150
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   820
         Picture         =   "PersonLedger.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "PersonLedger.frx":001C
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   22
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   15390
      Begin VB.Frame frmInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2190
         Left            =   300
         TabIndex        =   19
         Top             =   3375
         Width           =   4515
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "CustomersOrSuppliers"
            Top             =   1200
            Width           =   3540
         End
         Begin VB.TextBox txtCustomersOrSuppliers 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1200
            Width           =   780
         End
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "Invoices.InvoiceMasterRefersTo"
            Top             =   75
            Width           =   3540
         End
         Begin VB.TextBox txtInvoiceMasterRefersTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   75
            Width           =   780
         End
         Begin VB.TextBox txtExcursionTypeID 
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
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox txtPersonID 
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
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "Persons.ID"
            Top             =   450
            Width           =   3540
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
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "ExcursionTypes.ExcursionTypeID"
            Top             =   825
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1575
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   4592
            Images          =   "PersonLedger.frx":0038
            Version         =   131072
            KeyCount        =   4
            Keys            =   "ÿÿÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   3090
         Left            =   300
         TabIndex        =   15
         Top             =   5625
         Width           =   7440
         Begin UserControls.newText txtPersonDescription 
            Height          =   465
            Left            =   1575
            TabIndex        =   1
            Top             =   825
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
         Begin UserControls.newText txtExcursionTypeDescription 
            Height          =   465
            Left            =   1575
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
         Begin UserControls.newDate mskInvoiceDateIssueFrom 
            Height          =   465
            Left            =   1575
            TabIndex        =   3
            Top             =   1875
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
         Begin UserControls.newDate mskInvoiceDateIssueTo 
            Height          =   465
            Left            =   3075
            TabIndex        =   4
            Top             =   1875
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   6600
            TabIndex        =   31
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
            PicNormal       =   "PersonLedger.frx":1248
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   6600
            TabIndex        =   32
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
            PicNormal       =   "PersonLedger.frx":17E2
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   2
            Left            =   6975
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
            Index           =   1
            Left            =   1125
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
            TabIndex        =   30
            Top             =   2625
            Width           =   7440
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
            Left            =   2475
            TabIndex        =   29
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
            TabIndex        =   27
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Εκδρομές"
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
            TabIndex        =   18
            Top             =   1425
            Width           =   690
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Επωνυμία"
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
            TabIndex        =   17
            Top             =   900
            Width           =   690
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
            Index           =   0
            Left            =   450
            TabIndex        =   16
            Top             =   1950
            Width           =   690
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
            TabIndex        =   28
            Top             =   0
            Width           =   7440
         End
      End
      Begin VB.Frame frmButtonFrame 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   225
         TabIndex        =   6
         Top             =   8850
         Width           =   8940
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   8421376
            ButtonShape     =   3
            ButtonStyle     =   2
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
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   255
            ButtonShape     =   3
            ButtonStyle     =   2
            Caption         =   "Κλείσιμο"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   700
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
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   8421376
            ButtonShape     =   3
            ButtonStyle     =   2
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
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   8421376
            ButtonShape     =   3
            ButtonStyle     =   2
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
            ForeColor       =   0
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3075
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   8421376
            ButtonShape     =   3
            ButtonStyle     =   2
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
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            BackColor       =   8421376
            ButtonShape     =   3
            ButtonStyle     =   2
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
      Begin iGrid300_10Tec.iGrid grdPersonLedger 
         Height          =   7290
         Left            =   225
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1500
         Width           =   14940
         _ExtentX        =   26353
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
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Κριτήρια αναζήτησης"
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
         Left            =   4125
         TabIndex        =   14
         Top             =   1125
         Width           =   11040
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Καρτέλα συναλλασόμενου"
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
         TabIndex        =   13
         Top             =   75
         Width           =   6060
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
Attribute VB_Name = "PersonLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList Then
            UpdateCriteriaLabels txtPersonDescription.text, txtExcursionTypeDescription.text, mskInvoiceDateIssueFrom.text, mskInvoiceDateIssueTo.text
            EnableGrid grdPersonLedger, False
            HighlightRow grdPersonLedger, 2, "", True
            UpdateButtons Me, 5, 0, ChangeEditButtonStatus(grdPersonLedger, Me.Tag, 1, 1), 1, 1, 1, 0
            Exit Function
        Else
            If Not blnErrors Then DisplayMessageRecordsNotFound
            frmCriteria.Visible = True
            txtPersonDescription.SetFocus
        End If
    End If

End Function

Private Function UpdateCriteriaLabels(company, excursionType, fromDate, toDate)

    Dim strCriteriaA As String
    Dim strCriteriaB As String

    strCriteriaA = "Πελάτης: [" & company & "] Εκδρομές: [" & excursionType & "]"
    strCriteriaB = "Από: [" & fromDate & "] Έως: [" & toDate & "]"
    
    lblCriteria.Caption = strCriteriaA & " " & strCriteriaB
    
End Function

Private Function EditRecord()

    'Κινήσεις τιμολόγησης
    If grdPersonLedger.CellValue(grdPersonLedger.CurRow, "ShowInList") = "1" Then
        'Λεωφορεία
        If grdPersonLedger.CellValue(grdPersonLedger.CurRow, "ShipID") = "0" Then
            With InvoicesOut
                .txtInvoiceOutShipID.text = "0"
                .lblTitle.Caption = "Εκδρομές λεωφορείων"
                .txtShipDescription.Visible = False
                .cmdIndex(6).Visible = False
                .cmdIndex(7).Visible = False
                .Tag = "False"
                If .SeekRecord("Sales", grdPersonLedger.CellValue(grdPersonLedger.CurRow, "TripID")) Then .Show 1, Me
            End With
        End If
        'Πλοία
        If grdPersonLedger.CellValue(grdPersonLedger.CurRow, "ShipID") <> "0" Then
            With InvoicesOut
                .txtInvoiceOutShipID.text = "1"
                .lblTitle.Caption = "Εκδρομές πλοίων"
                .Tag = "False"
                If .SeekRecord("Sales", grdPersonLedger.CellValue(grdPersonLedger.CurRow, "TripID")) Then .Show 1, Me
            End With
        End If
    End If
    
    'Οικονομικές κινήσεις
    If grdPersonLedger.CellValue(grdPersonLedger.CurRow, "ShowInList") <> "1" Then
        With PersonsTransactions
            If .SeekRecord(grdPersonLedger.CellValue(grdPersonLedger.CurRow, "TripID"), grdPersonLedger.CellValue(grdPersonLedger.CurRow, "ShowInList")) Then
                .Tag = "False"
                .Show 1, Me
            End If
        End With
    End If
                
    'Κάνω focus
    grdPersonLedger.SetFocus

End Function

Private Function CreateUnicodeFile(strReportTitle, strReportSubTitle1, strReportSubTitle2, intReportDetailLines)

    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    'Μετρητές
    Dim intAdults As Integer
    Dim intKids As Integer
    Dim intFree As Integer
    Dim curAdultCompany As Currency
    Dim curKidCompany As Currency
    Dim curDebit As Currency
    Dim curCredit As Currency
    Dim curBalance As Currency

    'Αρχικές τιμές
    intAdults = 0
    intKids = 0
    intFree = 0
    curAdultsAmount = 0
    curKidsAmount = 0
    curDebit = 0
    curCredit = 0
    curBalance = 0
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Open strUnicodeFile For Output As #1
    GoSub Headers
    
    'Πλέγμα
    With grdPersonLedger
        For lngRow = 1 To grdPersonLedger.RowCount
            
            'Εκτυπώνω τη γραμμή
            Print #1, _
                Format(.CellText(lngRow, "Date"), "dd/mm/yy"); _
                Tab(10); .CellText(lngRow, "InvoiceDetails"); _
                Tab(24); .CellText(lngRow, "Destination"); _
                Tab(61 - Len((Format(.CellText(lngRow, "Adults"), "#,##0")))); Format(.CellText(lngRow, "Adults"), "#,##0"); _
                Tab(67 - Len((Format(.CellText(lngRow, "Kids"), "#,##0")))); Format(.CellText(lngRow, "Kids"), "#,##0"); _
                Tab(73 - Len((Format(.CellText(lngRow, "Free"), "#,##0")))); Format(.CellText(lngRow, "Free"), "#,##0"); _
                Tab(84 - Len((Format(.CellText(lngRow, "AdultsAmount"), "#,##0.00")))); Format(.CellText(lngRow, "AdultsAmount"), "#,##0.00"); _
                Tab(94 - Len((Format(.CellText(lngRow, "KidsAmount"), "#,##0.00")))); Format(.CellText(lngRow, "KidsAmount"), "#,##0.00"); _
                Tab(105 - Len((Format(.CellText(lngRow, "Debit"), "#,##0.00")))); Format(.CellText(lngRow, "Debit"), "#,##0.00"); _
                Tab(116 - Len((Format(.CellText(lngRow, "Credit"), "#,##0.00")))); Format(.CellText(lngRow, "Credit"), "#,##0.00"); _
                Tab(128 - Len((Format(.CellText(lngRow, "Balance"), "#,##0.00")))); Format(.CellText(lngRow, "Balance"), "#,##0.00")
            
            'Σύνολα
            If .CellText(lngRow, "TrnID") <> "" Then
                intAdults = intAdults + .CellValue(lngRow, "Adults")
                intKids = intKids + .CellValue(lngRow, "Kids")
                intFree = intFree + .CellValue(lngRow, "Free")
                curAdultsAmount = curAdultsAmount + .CellValue(lngRow, "AdultsAmount")
                curKidsAmount = curKidsAmount + .CellValue(lngRow, "KidsAmount")
                curDebit = curDebit + .CellValue(lngRow, "Debit")
                curCredit = curCredit + .CellValue(lngRow, "Credit")
                curBalance = curDebit - curCredit
            End If
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            
            'Eject
            If intProcessedDetailLines > intReportDetailLines Then
                Print #1, ""
                Print #1, Space(23) & "ΣΕ ΜΕΤΑΦΟΡΑ"; _
                Tab(61 - Len(Format(intAdults, "#,##0"))); Format(intAdults, "#,##0"); _
                Tab(67 - Len(Format(intKids, "#,##0"))); Format(intKids, "#,##0"); _
                Tab(73 - Len(Format(intFree, "#,##0"))); Format(intFree, "#,##0"); _
                Tab(84 - Len(Format(curAdultsAmount, "#,##0.00"))); Format(curAdultsAmount, "#,##0.00"); _
                Tab(94 - Len(Format(curKidsAmount, "#,##0.00"))); Format(curKidsAmount, "#,##0.00"); _
                Tab(105 - Len(Format(curDebit, "#,##0.00"))); Format(curDebit, "#,##0.00"); _
                Tab(116 - Len(Format(curCredit, "#,##0.00"))); Format(curCredit, "#,##0.00"); _
                Tab(128 - Len(Format(curBalance, "#,##0.00"))); Format(curBalance, "#,##0.00")
                
                GoSub Headers
                
                Print #1, Space(23) & "ΑΠΟ ΜΕΤΑΦΟΡΑ"; _
                    Tab(61 - Len(Format(intAdults, "#,##0"))); Format(intAdults, "#,##0"); _
                    Tab(67 - Len(Format(intKids, "#,##0"))); Format(intKids, "#,##0"); _
                    Tab(73 - Len(Format(intFree, "#,##0"))); Format(intFree, "#,##0"); _
                    Tab(84 - Len(Format(curAdultsAmount, "#,##0.00"))); Format(curAdultsAmount, "#,##0.00"); _
                    Tab(94 - Len(Format(curKidsAmount, "#,##0.00"))); Format(curKidsAmount, "#,##0.00"); _
                    Tab(105 - Len(Format(curDebit, "#,##0.00"))); Format(curDebit, "#,##0.00"); _
                    Tab(116 - Len(Format(curCredit, "#,##0.00"))); Format(curCredit, "#,##0.00"); _
                    Tab(128 - Len(Format(curBalance, "#,##0.00"))); Format(curBalance, "#,##0.00")
                Print #1, ""
                intProcessedDetailLines = intProcessedDetailLines + 2
            End If
            
        Next lngRow
    End With
    
    Close #1
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 127, intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
    PrintColumnHeadings 10, "ΣΤΟΙΧΕΙΟ", 56, "ΕΝΗΛΙ", 64, "ΠΑΙ", 71, "ΔΩ", 76, "ΧΡΕΩΣΕΙΣ", 86, "ΧΡΕΩΣΕΙΣ", 99, "ΣΥΝΟΛΟ"
    PrintColumnHeadings 1, "ΗΜΕΡ/ΝΙΑ", 10, "ΣΕΙΡΑ - Νο", 24, "ΠΡΟΟΡΙΣΜΟΣ", 58, "ΚΕΣ", 64, "ΔΙΑ", 69, "ΡΕΑΝ", 76, "ΕΝΗΛΙΚΩΝ", 87, "ΠΑΙΔΙΩΝ", 98, "ΧΡΕΩΣΗΣ", 109, "ΠΙΣΤΩΣΗ", 120, "ΥΠΟΛΟΙΠΟ"
    Print #1, ""
    intProcessedDetailLines = 7
      
    Return
    
End Function

Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
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
    
    'On Error GoTo ErrTrap
    
    If action = "Print" Then
        If SelectPrinter("PrinterPrintsReports") Then
            CreateUnicodeFile "ΚΑΡΤΕΛΑ ΠΕΛΑΤΗ " & txtPersonDescription.text & " ΕΚΔΡΟΜΕΣ " & txtExcursionTypeDescription.text, "ΑΠΟ " & mskInvoiceDateIssueFrom.text & " ΕΩΣ " & mskInvoiceDateIssueTo.text, "", intPrinterReportDetailLines
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
    End If
    
    If action = "CreatePDF" Then
        CreateUnicodeFile "ΚΑΡΤΕΛΑ ΠΕΛΑΤΗ " & txtPersonDescription.text & " ΕΚΔΡΟΜΕΣ " & txtExcursionTypeDescription.text, "ΑΠΟ " & mskInvoiceDateIssueFrom.text & " ΕΩΣ " & mskInvoiceDateIssueTo.text, "", GetSetting(strAppTitle, "Settings", "Export Report Height")
        CreateUnisexPDF "ΚΑΡΤΕΛΑ ΠΕΛΑΤΗ " & txtPersonDescription.text
        If MyMsgBox(1, strAppTitle, strStandardMessages(8), 1) Then
        End If
    End If
    
    Exit Function
    
ErrTrap:
    Close #1
    DisplayErrorMessage True, Err.description

End Function

Private Function ValidateFields()

    'Αρχικές τιμές
    ValidateFields = False
    
    'Πελάτης
    If txtPersonID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtPersonDescription.SetFocus
        Exit Function
    End If
    
    'Τύπος εκδρομής
    If txtExcursionTypeID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtExcursionTypeDescription.SetFocus
        Exit Function
    End If
    
    'Από
    If mskInvoiceDateIssueFrom.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskInvoiceDateIssueFrom.SetFocus
        Exit Function
    End If
    
    'Εως
    If mskInvoiceDateIssueTo.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskInvoiceDateIssueTo.SetFocus
        Exit Function
    End If
    
    'Σωστό διάστημα
    If IsDate(mskInvoiceDateIssueFrom.text) And IsDate(mskInvoiceDateIssueTo.text) Then
        If CDate(mskInvoiceDateIssueFrom.text) > CDate(mskInvoiceDateIssueTo.text) Then
            If MyMsgBox(4, strAppTitle, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateIssueFrom.SetFocus
            Exit Function
        End If
    End If

    ValidateFields = True
    
End Function

Private Function AbortProcedure(blnStatus)

    If Not blnStatus Then
        ClearFields grdPersonLedger, lblCriteria
        frmCriteria.Visible = True
        txtPersonDescription.SetFocus
        UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    End If
    
    If blnStatus Then
        Unload Me
    End If

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
    Dim rstInvoices As Recordset

    'Helpers
    Dim strFullInvoice As String
    Dim blnSoFarHasData As Boolean
    Dim blnPeriodHasData As Boolean
    
    'Τρέχουσα εγγραφή
    Dim lngAdultsLine As Long
    Dim lngKidsLine As Long
    Dim lngFreeLine As Long
    Dim curAdultsAmountLine As Currency
    Dim curKidsAmountLine As Currency
    Dim curDirectAmountLine As Currency
    Dim curDebitLine As Currency
    Dim curCreditLine As Currency
    
    'Προοδευτικό υπόλοιπο
    Dim curBalance As Currency
    
    'Προηγούμενη περίοδος
    'Ατομα
    Dim lngAdultsSoFar As Long
    Dim lngKidsSoFar As Long
    Dim lngFreeSoFar As Long
    'Ποσά
    Dim curAdultsAmountSoFar As Currency
    Dim curKidsAmountSoFar As Currency
    Dim curDirectAmountSoFar As Currency
    Dim curDebitSoFar As Currency
    Dim curCreditSoFar As Currency
    Dim curBalanceSoFar As Currency
    
    'Ζητούμενη περίοδος
    'Ατομα
    Dim lngAdultsPeriod As Long
    Dim lngKidsPeriod As Long
    Dim lngFreePeriod As Long
    'Ποσά
    Dim curAdultsAmountPeriod As Currency
    Dim curKidsAmountPeriod As Currency
    Dim curDirectAmountPeriod As Currency
    Dim curDebitPeriod As Currency
    Dim curCreditPeriod As Currency
    Dim curBalancePeriod As Currency
    
    'Στήλη χρέωσης
    Dim curDebitColumn As Currency
    Dim strDebitColumn As String
    
    'Στήλη πίστωσης
    Dim curCreditColumn As Currency
    Dim strCreditColumn As String
    
    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    frmCriteria.Visible = False
    blnPeriodHasData = False
    
    'Πλέγμα
    With grdPersonLedger
        .Clear
        .Redraw = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT " _
        & "InvoiceTrnID, InvoiceMasterRefersTo, InvoiceDateIssue, InvoiceNo, InvoicePersonID, CodeShortDescriptionA, CodeShortDescriptionB, CodeDescription, CodeBatch, CodeCustomers, CodeSuppliers " _
        & "FROM Invoices " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID "
 
    'Αγορές = 1, Πωλήσεις = 2
    strThisParameter = "strMaster String"
    strThisQuery = "(Invoices.InvoiceMasterRefersTo = strMaster"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = txtInvoiceMasterRefersTo.text
    
    'Πληρωμές = 3, Εισπράξεις = 4
    strThisParameter = "strSecondary String"
    strThisQuery = "Invoices.InvoiceMasterRefersTo = strSecondary)"
    strLogic = " OR "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Trim(Str(Val(txtInvoiceMasterRefersTo.text) + 2))
    
    'Συναλλασόμενος
    strThisParameter = "intPersonID Integer"
    strThisQuery = "Invoices.InvoicePersonID = intPersonID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = Val(txtPersonID.text)
    
    'Εως
    If IsDate(mskInvoiceDateIssueTo.text) Then
        strThisParameter = "datTo Date"
        strThisQuery = "Invoices.InvoiceDateIssue <= datTo "
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = CDate(mskInvoiceDateIssueTo.text)
    End If
    
    'Ταξινόμηση
    strOrder = " ORDER BY InvoiceDateIssue, InvoiceNo"
    
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
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        Do While Not .EOF
            GoSub CalculateSoFar
            If blnSoFarHasData Then GoSub UpdateSoFar
            Do While Not .EOF
                GoSub AddMasterInfoToTheGrid
                If !InvoiceMasterRefersTo = "1" Or !InvoiceMasterRefersTo = "2" Then
                    strSQL = "SELECT * FROM InvoicesOut INNER JOIN Destinations ON InvoicesOut.InvoiceOutDestinationID = Destinations.DestinationID WHERE InvoiceOutTrnID = " & !InvoiceTrnID
                    TempQuery.SQL = strSQL
                    Set rstInvoices = TempQuery.OpenRecordset()
                    If rstInvoices.RecordCount = 1 Then
                        GoSub CalculateDebitLine
                        GoSub CalculatePeriodTotals
                        GoSub AddDebitLineToTheGrid
                    End If
                End If
                If !InvoiceMasterRefersTo = "3" Or !InvoiceMasterRefersTo = "4" Then
                    strSQL = "SELECT * FROM PaymentsIn INNER JOIN PaymentWays ON PaymentsIn.PaymentInPaymentWayID = PaymentWays.PaymentWayID WHERE PaymentInTrnID = " & !InvoiceTrnID
                    TempQuery.SQL = strSQL
                    Set rstInvoices = TempQuery.OpenRecordset()
                    If rstInvoices.RecordCount = 1 Then
                        GoSub CalculateCreditLine
                        GoSub AddCreditLineToTheGrid
                    End If
                End If
                GoSub AddBalanceCellToTheGrid
                .MoveNext
                blnPeriodHasData = True
            Loop
        Loop
    End With
    
    If blnPeriodHasData Then GoSub ShowPeriodTotals
    If blnSoFarHasData And blnPeriodHasData Then GoSub ShowGrandTotals
        
    'Τρόπος επιστροφής
    RefreshList = True
    
    'Τελικές ενέργειες
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
    
CalculateSoFar:
    Dim strAdults As String
    Dim strKids As String
    Dim strFree As String
    With rstRecordset
        Do Until !InvoiceDateIssue >= CDate(mskInvoiceDateIssueFrom.text)
            If !InvoiceMasterRefersTo = "1" Or !InvoiceMasterRefersTo = "2" Then
                strSQL = "SELECT * FROM InvoicesOut WHERE InvoiceOutTrnID = " & !InvoiceTrnID
                TempQuery.SQL = strSQL
                Set rstInvoices = TempQuery.OpenRecordset()
                If rstInvoices.RecordCount = 1 Then
                    GoSub CalculateDebitLine
                    GoSub CalculateSoFarTotals
                End If
            End If
            If !InvoiceMasterRefersTo = "3" Or !InvoiceMasterRefersTo = "4" Then
                strSQL = "SELECT * FROM PaymentsIn WHERE PaymentInTrnID = " & !InvoiceTrnID
                TempQuery.SQL = strSQL
                Set rstInvoices = TempQuery.OpenRecordset()
                If rstInvoices.RecordCount = 1 Then
                    curCreditLine = IIf(!CodeCustomers = "-", rstInvoices!PaymentInAmount, -rstInvoices!PaymentInAmount)
                    curCreditSoFar = curCreditSoFar + curCreditLine
                    curBalanceSoFar = curBalanceSoFar - curCreditLine
                End If
            End If
            UpdateProgressBar Me
            .MoveNext
            blnSoFarHasData = True
            If .EOF Then
                Exit Do
            End If
        Loop
    End With
    
    Return
    
UpdateSoFar:
    With grdPersonLedger
        grdPersonLedger.AddRow
        lngRow = grdPersonLedger.RowCount
        .CellValue(lngRow, "Destination") = "ΠΡΟΗΓΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ"
        .CellValue(lngRow, "Adults") = Format(lngAdultsSoFar, "#,##0")
        .CellValue(lngRow, "Kids") = Format(lngKidsSoFar, "#,##0")
        .CellValue(lngRow, "Free") = Format(lngFreeSoFar, "#,##0")
        .CellValue(lngRow, "AdultsAmount") = Format(curAdultsAmountSoFar, "#,##0.00")
        .CellValue(lngRow, "KidsAmount") = Format(curKidsAmountSoFar, "#,##0.00")
        .CellValue(lngRow, "DirectAmount") = Format(curDirectAmountSoFar, "#,##0.00")
        .CellValue(lngRow, "Debit") = Format(curDebitSoFar, "#,##0.00")
        .CellValue(lngRow, "Credit") = Format(curCreditSoFar, "#,##0.00")
        .CellValue(lngRow, "Balance") = Format(curBalanceSoFar, "#,##0.00")
        grdPersonLedger.AddRow
        lngRow = grdPersonLedger.RowCount
    End With
    
    Return
    
ShowPeriodTotals:
    With grdPersonLedger
        grdPersonLedger.AddRow
        lngRow = grdPersonLedger.RowCount
        grdPersonLedger.AddRow
        lngRow = grdPersonLedger.RowCount
        .CellValue(lngRow, "Destination") = "ΖΗΤΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ"
        .CellValue(lngRow, "Adults") = Format(lngAdultsPeriod, "#,##0")
        .CellValue(lngRow, "Kids") = Format(lngKidsPeriod, "#,##0")
        .CellValue(lngRow, "Free") = Format(lngFreePeriod, "#,##0")
        .CellValue(lngRow, "AdultsAmount") = Format(curAdultsAmountPeriod, "#,##0.00")
        .CellValue(lngRow, "KidsAmount") = Format(curKidsAmountPeriod, "#,##0.00")
        .CellValue(lngRow, "DirectAmount") = Format(curDirectAmountPeriod, "#,##0.00")
        .CellValue(lngRow, "Debit") = Format(curDebitPeriod, "#,##0.00")
        .CellValue(lngRow, "Credit") = Format(curCreditPeriod, "#,##0.00")
        .CellValue(lngRow, "Balance") = Format(curBalancePeriod, "#,##0.00")
    End With
    
    Return
    
ShowGrandTotals:
    With grdPersonLedger
        grdPersonLedger.AddRow
        lngRow = grdPersonLedger.RowCount
        grdPersonLedger.AddRow
        lngRow = grdPersonLedger.RowCount
        .CellValue(lngRow, "Destination") = "ΓΕΝΙΚΑ ΣΥΝΟΛΑ"
        .CellValue(lngRow, "Adults") = Format(lngAdultsSoFar + lngAdultsPeriod, "#,##0")
        .CellValue(lngRow, "Kids") = Format(lngKidsSoFar + lngKidsPeriod, "#,##0")
        .CellValue(lngRow, "Free") = Format(lngFreeSoFar + lngFreePeriod, "#,##0")
        .CellValue(lngRow, "AdultsAmount") = Format(curAdultsAmountSoFar + curAdultsAmountPeriod, "#,##0.00")
        .CellValue(lngRow, "KidsAmount") = Format(curKidsAmountSoFar + curKidsAmountPeriod, "#,##0.00")
        .CellValue(lngRow, "DirectAmount") = Format(curDirectAmountSoFar + curDirectAmountPeriod, "#,##0.00")
        .CellValue(lngRow, "Debit") = Format(curDebitSoFar + curDebitPeriod, "#,##0.00")
        .CellValue(lngRow, "Credit") = Format(curCreditSoFar + curCreditPeriod, "#,##0.00")
        .CellValue(lngRow, "Balance") = Format(curBalanceSoFar + curBalancePeriod, "#,##0.00")
    End With
    
    Return
    
    
CalculateDebitLine:
    With rstRecordset
        lngAdultsLine = IIf(!CodeCustomers = "+", rstInvoices!InvoiceOutAdultsWithTransfer + rstInvoices!InvoiceOutAdultsWithoutTransfer, -rstInvoices!InvoiceOutAdultsWithTransfer - rstInvoices!InvoiceOutAdultsWithoutTransfer)
        lngKidsLine = IIf(!CodeCustomers = "+", rstInvoices!InvoiceOutKidsWithTransfer + rstInvoices!InvoiceOutKidsWithoutTransfer, -rstInvoices!InvoiceOutKidsWithTransfer - rstInvoices!InvoiceOutKidsWithoutTransfer)
        lngFreeLine = IIf(!CodeCustomers = "+", rstInvoices!InvoiceOutFreeWithTransfer + rstInvoices!InvoiceOutFreeWithoutTransfer, -rstInvoices!InvoiceOutFreeWithTransfer - rstInvoices!InvoiceOutFreeWithoutTransfer)
        curAdultsAmountLine = IIf(!CodeCustomers = "+", rstInvoices!InvoiceOutAdultsAmountWithTransfer + rstInvoices!InvoiceOutAdultsAmountWithoutTransfer, -rstInvoices!InvoiceOutAdultsAmountWithTransfer - rstInvoices!InvoiceOutAdultsAmountWithoutTransfer)
        curKidsAmountLine = IIf(!CodeCustomers = "+", rstInvoices!InvoiceOutKidsAmountWithTransfer + rstInvoices!InvoiceOutKidsAmountWithoutTransfer, -rstInvoices!InvoiceOutKidsAmountWithTransfer - rstInvoices!InvoiceOutKidsAmountWithoutTransfer)
        curDirectAmountLine = IIf(!CodeCustomers = "+", rstInvoices!InvoiceOutDirectAmount, -rstInvoices!InvoiceOutDirectAmount)
        curDebitLine = curAdultsAmountLine + curKidsAmountLine + curDirectAmountLine
        curCreditLine = 0
    End With

    Return
    
CalculateCreditLine:
    curCreditLine = IIf(rstRecordset!CodeCustomers = "-", rstInvoices!PaymentInAmount, -rstInvoices!PaymentInAmount)
    curCreditPeriod = curCreditPeriod + curCreditLine
    curBalancePeriod = curBalancePeriod - curCreditLine
    
    Return

CalculatePeriodTotals:
    lngAdultsPeriod = lngAdultsPeriod + lngAdultsLine
    lngKidsPeriod = lngKidsPeriod + lngKidsLine
    lngFreePeriod = lngFreePeriod + lngFreeLine
    curAdultsAmountPeriod = curAdultsAmountPeriod + curAdultsAmountLine
    curKidsAmountPeriod = curKidsAmountPeriod + curKidsAmountLine
    curDirectAmountPeriod = curDirectAmountPeriod + curDirectAmountLine
    curDebitPeriod = curDebitPeriod + curDebitLine
    curCreditPeriod = curCreditPeriod + curCreditLine
    curBalancePeriod = curBalancePeriod + curAdultsAmountLine + curKidsAmountLine + curDirectAmountLine - curCreditLine
    
    Return
    
CalculateSoFarTotals:
    lngAdultsSoFar = lngAdultsSoFar + lngAdultsLine
    lngKidsSoFar = lngKidsSoFar + lngKidsLine
    lngFreeSoFar = lngFreeSoFar + lngFreeLine
    curAdultsAmountSoFar = curAdultsAmountSoFar + curAdultsAmountLine
    curKidsAmountSoFar = curKidsAmountSoFar + curKidsAmountLine
    curDirectAmountSoFar = curDirectAmountSoFar + curDirectAmountLine
    curDebitSoFar = curDebitSoFar + curDebitLine
    curBalanceSoFar = curBalanceSoFar + curAdultsAmountLine + curKidsAmountLine + curDirectAmountLine
    
    Return

AddDebitLineToTheGrid:
    grdPersonLedger.CellValue(lngRow, "Destination") = rstInvoices!DestinationDescription
    grdPersonLedger.CellValue(lngRow, "Adults") = Format(lngAdultsLine, "#,##0")
    grdPersonLedger.CellValue(lngRow, "Kids") = Format(lngKidsLine, "#,##0")
    grdPersonLedger.CellValue(lngRow, "Free") = Format(lngFreeLine, "#,##0")
    grdPersonLedger.CellValue(lngRow, "AdultsAmount") = Format(curAdultsAmountLine, "#,##0.00")
    grdPersonLedger.CellValue(lngRow, "KidsAmount") = Format(curKidsAmountLine, "#,##0.00")
    grdPersonLedger.CellValue(lngRow, "DirectAmount") = Format(curDirectAmountLine, "#,##0.00")
    grdPersonLedger.CellValue(lngRow, "Debit") = Format(curDebitLine, "#,##0.00")
    grdPersonLedger.CellValue(lngRow, "Credit") = Format(curCreditLine, "#,##0.00")
    
    Return
    
AddCreditLineToTheGrid:
    grdPersonLedger.CellValue(lngRow, "Destination") = rstRecordset!CodeDescription
    grdPersonLedger.CellValue(lngRow, "Adults") = 0
    grdPersonLedger.CellValue(lngRow, "Kids") = 0
    grdPersonLedger.CellValue(lngRow, "Free") = 0
    grdPersonLedger.CellValue(lngRow, "AdultsAmount") = 0
    grdPersonLedger.CellValue(lngRow, "KidsAmount") = 0
    grdPersonLedger.CellValue(lngRow, "DirectAmount") = 0
    grdPersonLedger.CellValue(lngRow, "Debit") = 0
    grdPersonLedger.CellValue(lngRow, "Credit") = Format(curCreditLine, "#,##0.00")
    
    Return
   
AddMasterInfoToTheGrid:
    lngRow = lngRow + 1
    grdPersonLedger.AddRow
    With rstRecordset
        grdPersonLedger.CellValue(lngRow, "TrnID") = !InvoiceTrnID
        grdPersonLedger.CellValue(lngRow, "Date") = !InvoiceDateIssue
        strFullInvoice = !CodeShortDescriptionB & Space(3 - Len(!CodeShortDescriptionB)) & " "
        strFullInvoice = strFullInvoice & IIf(!CodeBatch <> "", !CodeBatch, "0") & " "
        strFullInvoice = strFullInvoice & Right("00000" & !InvoiceNo, 5)
        grdPersonLedger.CellValue(lngRow, "InvoiceDetails") = strFullInvoice
    End With
    
    UpdateProgressBar Me
    
    Return
    
AddBalanceCellToTheGrid:
    grdPersonLedger.CellValue(lngRow, "Balance") = Format(curBalanceSoFar + curBalancePeriod, "#,##0.00")
    
    Return
    
ErrTrap:
    blnErrors = True
    ClearFields grdPersonLedger, frmProgress
    DisplayErrorMessage True, Err.description

End Function

Private Sub cmdIndex_Click(Index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Companies
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "Description", "String", txtPersonDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Επωνυμία", 0, 40, 1, 0)
            txtPersonID.text = tmpTableData.strCode
            txtPersonDescription.text = tmpTableData.strFirstField
        Case 1
            'Excursion type
            Set tmpRecordset = CheckForMatch("CommonDB", "ExcursionTypes", "ExcursionTypeDescription", "String", txtExcursionTypeDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Εκδρομές", 0, 40, 1, 0)
            txtExcursionTypeID.text = tmpTableData.strCode
            txtExcursionTypeDescription.text = tmpTableData.strFirstField
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdPersonLedger, 44, GetSetting(strAppTitle, "Layout Strings", "grdPersonLedger"), _
            "05NCITrnID,12NCDDate,50NLNInvoiceDetails,40NLNDestination,10NRIAdults,10NRIKids,10NRIFree,10NRFXAdultsAmount,10NRFXKidsAmount,10NRFXDirectAmount,10NRFXDebit,10NRFCredit,10NRFXBalance", _
            "TrnID,Εκδοση,Παραστατικό,Προορισμός,Ενήλικες,Παιδιά,Δωρεάν,Χρέωση ενηλίκων,Χρέωση παιδιών,Απευθείας ποσό,Σύνολο χρέωσης,Πίστωση,Προοδευτικό Υπόλοιπο"
        Me.Refresh
        frmCriteria.Visible = True
        txtPersonDescription.SetFocus
    End If
            
    'AddDummyLines grdPersonLedger, "99999", "A99/99/9999A", "ΑΑΑΑΑΑΑΑΑΑΑΑ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "999999", "999999", "999999", "-9999999", "-9999999", "-9999999", "-9999999", "-9999999", "-9999999"
   
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
        Case vbKeyP And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    UpdateColors Me, True, grdPersonLedger
    SetUpGrid grdPersonLedger, lstIconList
    ClearFields txtPersonID, txtPersonDescription, txtExcursionTypeID, txtExcursionTypeDescription, mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, lblCriteria
    EnableFields txtPersonDescription, txtExcursionTypeDescription, mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo
    UpdateButtons Me, 5, 1, 0, 0, 0, 0, 1
    
    txtPersonID.text = "35"
    txtPersonDescription.text = "TUI HELLAS"
    txtExcursionTypeID.text = "2"
    txtExcursionTypeDescription.text = "ΠΛΟΙΩΝ"
    mskInvoiceDateIssueFrom.text = "01/01/2018"
    mskInvoiceDateIssueTo.text = "31/12/2018"

End Sub

Private Sub grdPersonLedger_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdPersonLedger_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    
    cmdButton(1).Enabled = ChangeEditButtonStatus(grdPersonLedger, Me.Tag, lRow, 1)

End Sub

Private Sub grdPersonLedger_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub grdPersonLedger_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonLedger_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cmdButton(1).Enabled Then cmdButton_Click 1

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdPersonLedger", grdPersonLedger.LayoutCol

End Sub

Private Sub txtExcursionTypeDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtExcursionTypeDescription_Validate(Cancel As Boolean)

    If txtExcursionTypeID.text = "" And txtExcursionTypeDescription.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtPersonDescription_Change()

    If txtPersonDescription.text = "" Then txtPersonID.text = ""

End Sub

Private Sub txtPersonDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtPersonDescription_Validate(Cancel As Boolean)

    If txtPersonID.text = "" And txtPersonDescription.text <> "" Then cmdIndex_Click 0

End Sub

Private Sub txtExcursionTypeDescription_Change()

    If txtExcursionTypeDescription.text = "" Then txtExcursionTypeID.text = ""
    
End Sub

