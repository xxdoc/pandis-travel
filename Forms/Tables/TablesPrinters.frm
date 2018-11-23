VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesPrinters 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   15
   ClientTop       =   105
   ClientWidth     =   16470
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   16470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   48
      Top             =   7875
      Width           =   7515
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   8550
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3825
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
      PicNormal       =   "TablesPrinters.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   3750
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4275
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
      PicNormal       =   "TablesPrinters.frx":059A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   3750
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3750
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
      PicNormal       =   "TablesPrinters.frx":0B34
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      DragIcon        =   "TablesPrinters.frx":10CE
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   11625
      TabIndex        =   33
      Top             =   5625
      Width           =   4515
      Begin VB.TextBox txtPrinterEAFDSSID 
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
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox Text5 
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
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Printers.PrinterEafdssID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtPrinterPrintsReportsID 
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
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
      Begin VB.TextBox txtPrinterPrintsInvoicesID 
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Printers.PrinterPrintsReportsID"
         Top             =   1575
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
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "Printers.PrinterPrintsInvoicesID"
         Top             =   825
         Width           =   3540
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
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Printers.PrinterID"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox txtPrinterID 
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin VB.TextBox txtPrinterTypeID 
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
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
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Printers.PrinterTypeID"
         Top             =   450
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   1950
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "TablesPrinters.frx":1998
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin UserControls.newInteger mskPrinterInvoiceHeight 
      Height          =   465
      Left            =   3075
      TabIndex        =   7
      Top             =   5325
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin iGrid300_10Tec.iGrid grdAppPrinters 
      Height          =   3090
      Left            =   9825
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1125
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   5450
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
   Begin UserControls.newText txtPrinterName 
      Height          =   465
      Left            =   1950
      TabIndex        =   0
      Top             =   1125
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      Locked          =   -1  'True
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
   Begin UserControls.newText txtPrinterEAFDSSString 
      Height          =   465
      Left            =   3075
      TabIndex        =   6
      Top             =   4800
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   10
      Text            =   "ΑΑΑΑΑΑΑΑΑΑ"
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
   Begin UserControls.newInteger mskPrinterInvoiceDetailLines 
      Height          =   465
      Left            =   3075
      TabIndex        =   8
      Top             =   5850
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newInteger mskPrinterInvoiceTopMargin 
      Height          =   465
      Left            =   3075
      TabIndex        =   9
      Top             =   6375
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newInteger mskPrinterReportDetailLines 
      Height          =   465
      Left            =   7875
      TabIndex        =   11
      Top             =   4350
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newInteger mskPrinterReportTopMargin 
      Height          =   465
      Left            =   7875
      TabIndex        =   12
      Top             =   4875
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newInteger mskPrinterReportLeftMargin 
      Height          =   465
      Left            =   7875
      TabIndex        =   14
      Top             =   5400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newInteger mskPrinterFontSize 
      Height          =   465
      Left            =   1950
      TabIndex        =   3
      Top             =   2700
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   2
      Text            =   "99"
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UserControls.newText txtPrinterFontName 
      Height          =   465
      Left            =   1950
      TabIndex        =   2
      Top             =   2175
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
   Begin iGrid300_10Tec.iGrid grdAvailablePrinters 
      Height          =   2940
      Left            =   9825
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4275
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   5186
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
   Begin UserControls.newText txtPrinterTypeDescription 
      Height          =   465
      Left            =   1950
      TabIndex        =   1
      Top             =   1650
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
   Begin UserControls.newText txtPrintsReportsDescription 
      Height          =   465
      Left            =   7875
      TabIndex        =   10
      Top             =   3825
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
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
   Begin UserControls.newText txtPrintsInvoicesDescription 
      Height          =   465
      Left            =   3075
      TabIndex        =   4
      Top             =   3750
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
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
   Begin UserControls.newText txtEafdssDescription 
      Height          =   465
      Left            =   3075
      TabIndex        =   5
      Top             =   4275
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Παραστατικά "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3990
      Index           =   0
      Left            =   450
      TabIndex        =   21
      Tag             =   "SameColorAsBackground"
      Top             =   3225
      Width           =   4740
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   2
         Left            =   4275
         Top             =   1575
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   1
         Left            =   2175
         Top             =   1500
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   0
         Left            =   3000
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   390
         Index           =   6
         Left            =   2775
         Top             =   3600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   5
         Left            =   0
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Με σήμανση"
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
         Height          =   255
         Index           =   15
         Left            =   450
         TabIndex        =   28
         Top             =   1125
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Εκτυπώνει παραστατικά"
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
         Height          =   255
         Index           =   14
         Left            =   450
         TabIndex        =   27
         Top             =   600
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Υψος σε γραμμές"
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
         TabIndex        =   26
         Top             =   2175
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Αναλυτικές γραμμές"
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
         Index           =   4
         Left            =   450
         TabIndex        =   25
         Top             =   2700
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Συμβολοσειρά σήμανσης"
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
         TabIndex        =   24
         Top             =   1650
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Επάνω περιθώριο"
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
         Index           =   7
         Left            =   450
         TabIndex        =   23
         Top             =   3225
         Width           =   1740
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Αναφορές "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3990
      Index           =   1
      Left            =   5250
      TabIndex        =   22
      Tag             =   "SameColorAsBackground"
      Top             =   3225
      Width           =   4140
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   4
         Left            =   3675
         Top             =   525
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   9
         Left            =   4500
         Top             =   1575
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   8
         Left            =   1950
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   7
         Left            =   2475
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   3
         Left            =   0
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Εκτυπώνει αναφορές"
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
         Index           =   12
         Left            =   450
         TabIndex        =   32
         Top             =   600
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Αναλυτικές γραμμές"
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
         TabIndex        =   31
         Top             =   1125
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Επάνω περιθώριο"
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
         Index           =   9
         Left            =   450
         TabIndex        =   30
         Top             =   1650
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Αριστερό περιθώριο"
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
         Index           =   10
         Left            =   450
         TabIndex        =   29
         Top             =   2175
         Width           =   1515
         WordWrap        =   -1  'True
      End
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   6975
      TabIndex        =   44
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
      PicNormal       =   "TablesPrinters.frx":2BA8
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   14
      Left            =   1500
      Top             =   1275
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
      Left            =   3900
      Top             =   0
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
      Top             =   3450
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   7050
      Top             =   8550
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   14475
      Top             =   2025
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   9375
      Top             =   7200
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   11
      Left            =   9975
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   10
      Left            =   9375
      Top             =   3225
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Τύπος"
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
      Index           =   8
      Left            =   450
      TabIndex        =   20
      Top             =   1650
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Εκτυπωτές"
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
      TabIndex        =   18
      Top             =   75
      Width           =   2475
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Γραμματοσειρά"
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
      Index           =   11
      Left            =   450
      TabIndex        =   16
      Top             =   2175
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Μέγεθος"
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
      Index           =   13
      Left            =   450
      TabIndex        =   15
      Top             =   2700
      Width           =   1065
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Όνομα"
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
      TabIndex        =   13
      Top             =   1200
      Width           =   1065
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
Attribute VB_Name = "TablesPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim lngSelectedRow  As Long

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtPrinterID, txtPrinterName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
            DisableFields txtPrinterName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3)
            grdAppPrinters.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("PrintersDB", "Printers", strAppTitle, "ID", txtPrinterID.text, "True") Then
        PopulateGrid
        HighlightRow grdAppPrinters, lngSelectedRow, 1, "", True
        ClearFields txtPrinterID, txtPrinterName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
        DisableFields txtPrinterName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3)
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If

End Function

Private Function FindPrinters()

    Dim prt As Printer
    Dim strSavedLayout As String
    
    With grdAvailablePrinters
        With .AddCol(sKey:="ID", sHeader:="ID", lWidth:=254, eHdrTextFlags:=igTextCenter)
            .eTextFlags = igTextCenter
        End With
        With .AddCol(sKey:="PrinterName", sHeader:="Όνομα", lWidth:=254, eHdrTextFlags:=igTextCenter)
            .eTextFlags = igTextLeft
        End With
    End With
    
    For Each prt In Printers
        grdAvailablePrinters.AddRow
        grdAvailablePrinters.CellValue(grdAvailablePrinters.RowCount, "PrinterName") = prt.DeviceName
    Next
    
    strSavedLayout = GetSetting(strAppTitle, "Layout Strings", "grdPrinters"): grdAvailablePrinters.LayoutCol = strSavedLayout

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtPrinterID, txtPrinterName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
    EnableFields txtPrinterName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3)
    InitializeFields mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, mskPrinterFontSize, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    txtPrinterName.SetFocus

End Function

Private Function PopulateGrid()

    If FillGridFromDB("PrintersDB", grdAppPrinters, "Printers", "", "", "", 2, 0, 1) Then
        grdAppPrinters.SetFocus
        grdAppPrinters.SetCurCell 1, 1
    End If

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("PrintersDB", "Printers", blnStatus, strAppTitle, "ID", txtPrinterID.text, txtPrinterName.text, txtPrinterTypeID.text, txtPrinterPrintsInvoicesID.text, txtPrinterEAFDSSID.text, txtPrinterEAFDSSString.text, mskPrinterInvoiceHeight.text, mskPrinterInvoiceDetailLines.text, mskPrinterInvoiceTopMargin.text, txtPrinterPrintsReportsID.text, mskPrinterReportDetailLines.text, mskPrinterReportTopMargin.text, mskPrinterReportLeftMargin.text, txtPrinterFontName.text, mskPrinterFontSize.text, 1, strCurrentUser) <> 0 Then
        PopulateGrid
        HighlightRow grdAppPrinters, lngSelectedRow, 2, txtPrinterName.text, True
        lngSelectedRow = 0
        ClearFields txtPrinterID, txtPrinterName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
        DisableFields txtPrinterName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3)
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If

End Function

Private Function SeekRecord()
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    If grdAppPrinters.RowCount = 0 Then Exit Function
    
    ClearFields txtPrinterID, txtPrinterName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
    DisableFields txtPrinterName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3)
    
    If MainSeekRecord("PrintersDB", "Printers", "ID", grdAppPrinters.CellValue(grdAppPrinters.CurRow, 1), True, txtPrinterID, txtPrinterName, txtPrinterTypeID, txtPrinterPrintsInvoicesID, txtPrinterEAFDSSID, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, txtPrinterFontName, mskPrinterFontSize) Then
        'Τύπος εκτυπωτή
        Set tmpRecordset = CheckForMatch("PrintersDB", "PrinterTypes", "PrinterTypeID", "Numeric", txtPrinterTypeID.text)
        txtPrinterTypeID.text = tmpRecordset.Fields(0)
        txtPrinterTypeDescription.text = tmpRecordset.Fields(1)
        'Εκτυπώνει παραστατικά;
        Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtPrinterPrintsInvoicesID.text)
        txtPrinterPrintsInvoicesID.text = tmpRecordset.Fields(0)
        txtPrintsInvoicesDescription.text = tmpRecordset.Fields(1)
        'Με σήμανση;
        Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtPrinterEAFDSSID.text)
        txtPrinterEAFDSSID.text = tmpRecordset.Fields(0)
        txtEafdssDescription.text = tmpRecordset.Fields(1)
        'Εκτυπώνει αναφορές;
        Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtPrinterPrintsReportsID.text)
        txtPrinterPrintsReportsID.text = tmpRecordset.Fields(0)
        txtPrintsReportsDescription.text = tmpRecordset.Fields(1)
        '
        blnStatus = False
        lngSelectedRow = grdAppPrinters.CurRow
        EnableFields txtPrinterName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3)
        UpdateButtons Me, 4, 0, 1, 1, 1, 0
        txtPrinterName.SetFocus
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Περιγραφή
    If txtPrinterName.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtPrinterName.SetFocus
        Exit Function
    End If
    
    'Τύπος
    If txtPrinterTypeID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtPrinterTypeDescription.SetFocus
        Exit Function
    End If
    
    'Εκτυπώνει παραστατικά
    If txtPrinterPrintsInvoicesID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtPrintsInvoicesDescription.SetFocus
        Exit Function
    End If
    
    'Εκτυπώνει παραστατικά με σήμανση
    If txtPrinterEAFDSSID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtEafdssDescription.SetFocus
        Exit Function
    End If
    
    'Υψος παραστατικού
    If mskPrinterInvoiceHeight.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPrinterInvoiceHeight.SetFocus
        Exit Function
    End If
    
    'Αναλυτικές γραμμές παραστατικού
    If mskPrinterInvoiceDetailLines.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPrinterInvoiceDetailLines.SetFocus
        Exit Function
    End If
    
    'Επάνω περιθώριο παραστατικού
    If mskPrinterInvoiceTopMargin.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPrinterInvoiceTopMargin.SetFocus
        Exit Function
    End If
    
    'Εκτυπώνει αναφορές
    If txtPrinterPrintsReportsID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtPrintsReportsDescription.SetFocus
        Exit Function
    End If
    
    'Αναλυτικές γραμμές αναφορών
    If mskPrinterReportDetailLines.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPrinterReportDetailLines.SetFocus
        Exit Function
    End If
    
    'Επάνω περιθώριο αναφορών
    If mskPrinterReportTopMargin.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPrinterReportTopMargin.SetFocus
        Exit Function
    End If
    
    'Αριστερό περιθώριο αναφορών
    If mskPrinterReportLeftMargin.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPrinterReportLeftMargin.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

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
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Τύπος εκτυπωτή
            Set tmpRecordset = CheckForMatch("PrintersDB", "PrinterTypes", "PrinterTypeDescription", "String", txtPrinterTypeDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterTypeID.text = tmpTableData.strCode
            txtPrinterTypeDescription.text = tmpTableData.strFirstField
        Case 1
            'Εκτυπώνει παραστατικά;
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtPrintsInvoicesDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterPrintsInvoicesID.text = tmpTableData.strCode
            txtPrintsInvoicesDescription.text = tmpTableData.strFirstField
        Case 2
            'Με σήμανση;
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtEafdssDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterEAFDSSID.text = tmpTableData.strCode
            txtEafdssDescription.text = tmpTableData.strFirstField
        Case 3
            'Εκτυπώνει αναφορές;
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtPrintsReportsDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtPrinterPrintsReportsID.text = tmpTableData.strCode
            txtPrintsReportsDescription.text = tmpTableData.strFirstField
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        FindPrinters
        Me.Tag = "False"
        AddColumnsToGrid grdAppPrinters, 25, GetSetting(strAppTitle, "Layout Strings", "grdPrinters"), "04NCIID,40NLNDescription", "ID,Όνομα"
        Me.Refresh
        PopulateGrid
    End If
    
    'AddDummyLines grdAppPrinters, "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"

End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyEscape
            If cmdButton(3).Enabled Then cmdButton_Click 3: Exit Function
            If cmdButton(4).Enabled Then cmdButton_Click 4
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Sub Form_Load()
    
    UpdateColors Me, False
    SetUpGrid lstIconList, grdAppPrinters, grdAvailablePrinters
    ClearFields txtPrinterID, txtPrinterName, txtPrinterTypeID, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrinterPrintsInvoicesID, txtPrintsInvoicesDescription, txtPrinterEAFDSSID, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrinterPrintsReportsID, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin
    DisableFields txtPrinterName, txtPrinterTypeDescription, txtPrinterFontName, mskPrinterFontSize, txtPrintsInvoicesDescription, txtEafdssDescription, txtPrinterEAFDSSString, mskPrinterInvoiceHeight, mskPrinterInvoiceDetailLines, mskPrinterInvoiceTopMargin, txtPrintsReportsDescription, mskPrinterReportDetailLines, mskPrinterReportTopMargin, mskPrinterReportLeftMargin, cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3)
    UpdateButtons Me, 4, 1, 0, 0, 0, 1
    
End Sub

Private Sub grdAvailablePrinters_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    If txtPrinterName.Enabled Then
        txtPrinterName.text = grdAvailablePrinters.CellValue(lRow, "PrinterName")
        txtPrinterName.SetFocus
    End If

End Sub

Private Sub grdAppPrinters_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub

Private Sub grdAppPrinters_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdAppPrinters_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()
    
    SaveSetting strAppTitle, "Layout Strings", "grdPrinters", grdAppPrinters.LayoutCol

End Sub

Private Sub txtEafdssDescription_Change()

    If txtEafdssDescription.text = "" Then
        ClearFields txtPrinterEAFDSSID
    End If

End Sub

Private Sub txtEafdssDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtEafdssDescription_Validate(Cancel As Boolean)

    If txtPrinterEAFDSSID.text = "" And txtEafdssDescription.text <> "" Then cmdIndex_Click 2

End Sub

Private Sub txtPrintsInvoicesDescription_Change()

    If txtPrintsInvoicesDescription.text = "" Then
        ClearFields txtPrinterPrintsInvoicesID
    End If

End Sub

Private Sub txtPrintsInvoicesDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtPrintsInvoicesDescription_Validate(Cancel As Boolean)

    If txtPrinterPrintsInvoicesID.text = "" And txtPrintsInvoicesDescription.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtPrintsReportsDescription_Change()

    If txtPrintsReportsDescription.text = "" Then
        ClearFields txtPrinterPrintsReportsID
    End If
    
End Sub

Private Sub txtPrintsReportsDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3
    
End Sub

Private Sub txtPrintsReportsDescription_Validate(Cancel As Boolean)

    If txtPrinterPrintsReportsID.text = "" And txtPrintsReportsDescription.text <> "" Then cmdIndex_Click 3
    
End Sub

Private Sub txtPrinterTypeDescription_Change()

    If txtPrinterTypeDescription.text = "" Then
        ClearFields txtPrinterTypeID
    End If

End Sub

Private Sub txtPrinterTypeDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtPrinterTypeDescription_Validate(Cancel As Boolean)

    If txtPrinterTypeID.text = "" And txtPrinterTypeDescription.text <> "" Then cmdIndex_Click 0

End Sub

