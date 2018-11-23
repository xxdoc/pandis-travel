VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesPrices 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   16620
   ControlBox      =   0   'False
   Icon            =   "TablesPriceLists.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   16620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   34
      Top             =   7875
      Width           =   10365
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   35
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
         Index           =   6
         Left            =   8775
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
         Index           =   5
         Left            =   7350
         TabIndex        =   39
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
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   3
         Left            =   4500
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
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
         Index           =   4
         Left            =   5925
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Μαζική επεξεργασία"
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
      Height          =   2190
      Left            =   7725
      TabIndex        =   23
      Top             =   5100
      Width           =   4515
      Begin VB.TextBox Text20 
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
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Prices.ShowInList"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtShowInList 
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
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox txtPriceDestinationID 
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox txtPriceID 
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
         TabIndex        =   28
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
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Prices.PriceID"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Prices.PriceDestinationID"
         Top             =   825
         Width           =   3540
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Prices.PriceCustomerID"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox txtPriceCustomerID 
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   1575
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "TablesPriceLists.frx":0442
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin VB.Frame frmFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Χρεώσεις χωρίς μεταφορά "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2265
      Index           =   2
      Left            =   3675
      TabIndex        =   12
      Tag             =   "SameColorAsBackground"
      Top             =   4125
      Width           =   3165
      Begin Dacara_dcButton.dcButton cmdHalf 
         Height          =   315
         Index           =   1
         Left            =   1575
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   8421376
         ButtonShape     =   3
         ButtonStyle     =   2
         Caption         =   "½"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin UserControls.newFloat mskPriceAdultWithoutTransfer 
         Height          =   465
         Left            =   1500
         TabIndex        =   21
         Top             =   525
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         MaxLength       =   8
         Text            =   "9.999,99"
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
      Begin UserControls.newFloat mskPriceKidWithoutTransfer 
         Height          =   465
         Left            =   1500
         TabIndex        =   22
         Top             =   1425
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         MaxLength       =   8
         Text            =   "9.999,99"
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
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   18
         Left            =   1050
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   16
         Left            =   2700
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   15
         Left            =   0
         Top             =   675
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   390
         Index           =   14
         Left            =   1725
         Top             =   1875
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   13
         Left            =   1650
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Παιδιά"
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
         Index           =   5
         Left            =   450
         TabIndex        =   14
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Ενήλικες"
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
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame frmFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Περίοδος "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1890
      Index           =   0
      Left            =   450
      TabIndex        =   4
      Tag             =   "SameColorAsBackground"
      Top             =   2175
      Width           =   3165
      Begin UserControls.newDate mskPriceFrom 
         Height          =   465
         Left            =   1200
         TabIndex        =   15
         Top             =   525
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   820
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
      Begin UserControls.newDate mskPriceTo 
         Height          =   465
         Left            =   1200
         TabIndex        =   16
         Top             =   1050
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   820
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
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   9
         Left            =   2700
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   390
         Index           =   6
         Left            =   1575
         Top             =   1500
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   2
         Left            =   750
         Top             =   525
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
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
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   7
         Left            =   1275
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Έως"
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
         TabIndex        =   6
         Top             =   1125
         Width           =   315
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Από"
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
         TabIndex        =   5
         Top             =   600
         Width           =   315
      End
   End
   Begin iGrid300_10Tec.iGrid grdPrices 
      Height          =   6240
      Left            =   7650
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1125
      Width           =   7815
      _ExtentX        =   13785
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
   Begin UserControls.newText txtPriceCustomerDescription 
      Height          =   465
      Left            =   1800
      TabIndex        =   1
      Top             =   1125
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
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
   Begin UserControls.newText txtPriceDestinationDescription 
      Height          =   465
      Left            =   1800
      TabIndex        =   2
      Top             =   1650
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
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
   Begin VB.Frame frmFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Χρεώσεις με μεταφορά "
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2265
      Index           =   1
      Left            =   450
      TabIndex        =   9
      Tag             =   "SameColorAsBackground"
      Top             =   4125
      Width           =   3165
      Begin Dacara_dcButton.dcButton cmdHalf 
         Height          =   315
         Index           =   0
         Left            =   1575
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   8421376
         ButtonShape     =   3
         ButtonStyle     =   2
         Caption         =   "½"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   12
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin UserControls.newFloat mskPriceAdultWithTransfer 
         Height          =   465
         Left            =   1500
         TabIndex        =   18
         Top             =   525
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         MaxLength       =   8
         Text            =   "9.999,99"
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
      Begin UserControls.newFloat mskPriceKidWithTransfer 
         Height          =   465
         Left            =   1500
         TabIndex        =   19
         Top             =   1425
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   820
         Alignment       =   1
         MaxLength       =   8
         Text            =   "9.999,99"
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
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   17
         Left            =   1050
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   390
         Index           =   12
         Left            =   1575
         Top             =   1875
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   11
         Left            =   2700
         Top             =   825
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   10
         Left            =   1725
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   8
         Left            =   0
         Top             =   675
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Ενήλικες"
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
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Παιδιά"
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
         Top             =   1500
         Width           =   615
      End
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   6825
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1125
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
      PicNormal       =   "TablesPriceLists.frx":1652
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   6825
      TabIndex        =   33
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
      PicNormal       =   "TablesPriceLists.frx":1BEC
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   7725
      Top             =   7350
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   15450
      Top             =   3450
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   5325
      Top             =   8550
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   1
      Left            =   1350
      Top             =   1350
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   3
      Left            =   7500
      Top             =   0
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
      Left            =   7200
      Top             =   1125
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   4
      Left            =   0
      Top             =   1425
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Τιμοκατάλογοι εκδρομών"
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
      Left            =   225
      TabIndex        =   8
      Top             =   75
      Width           =   5610
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Index           =   11
      Left            =   450
      TabIndex        =   3
      Top             =   1725
      Width           =   915
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Index           =   7
      Left            =   450
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   -75
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
Attribute VB_Name = "TablesPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnStatus As Boolean
Dim lngSelectedRow As Long
Dim blnInputingData As Boolean

Private Function BatchProcessing()

    With UtilsBatchPriceListEdit
        .Tag = "True"
        .Show 1, Me
    End With
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", "Prices", strAppTitle, "PriceID", txtPriceID.text, "True") Then
        blnInputingData = False
        If PopulateGrid Then
            HighlightRow grdPrices, lngSelectedRow, 1, "", True
        End If
        ClearFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer
        DisableFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer, cmdHalf(0), cmdHalf(1)
        UpdateButtons Me, 6, 1, 0, 0, 0, 1, 1, 0
    End If
        
End Function

Private Function PopulateGrid()

    If FillGridFromDB("CommonDB", grdPrices, "Prices", "", "", "PriceCustomerID= " & Val(txtPriceCustomerID.text) & " AND PriceDestinationID = " & Val(txtPriceDestinationID.text), 4, 0, 1, 2, 3, 4, 5, 6, 7, 8) Then
        PopulateGrid = True
    End If

End Function

Private Function NewRecord()
    
    blnStatus = True
    blnInputingData = True
    ClearFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer
    DisableFields txtPriceCustomerDescription, txtPriceDestinationDescription
    EnableFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer, cmdHalf(0), cmdHalf(1)
    UpdateButtons Me, 6, 0, 1, 0, 0, 0, 1, 0
    InitializeFields mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer
    mskPriceFrom.SetFocus
    
End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "Prices", blnStatus, strAppTitle, "PriceID", txtPriceID.text, txtPriceCustomerID.text, txtPriceDestinationID.text, mskPriceFrom.text, mskPriceTo.text, mskPriceAdultWithTransfer.text, mskPriceKidWithTransfer.text, mskPriceAdultWithoutTransfer.text, mskPriceKidWithoutTransfer.text, txtShowInList.text, strCurrentUser) <> 0 Then
        blnInputingData = False
        PopulateGrid
        HighlightRow grdPrices, lngSelectedRow, 4, mskPriceFrom.text, True
        lngSelectedRow = 0
        ClearFields txtPriceID, mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer
        DisableFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer, cmdHalf(0), cmdHalf(1)
        UpdateButtons Me, 6, 1, 0, 0, 0, 1, 1, 0
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If

End Function

Private Function SeekRecord()

    If grdPrices.RowCount = 0 Then Exit Function
    
    ClearFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer
    DisableFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer
    
    If MainSeekRecord("CommonDB", "Prices", "PriceID", grdPrices.CellValue(grdPrices.CurRow, "PriceID"), True, _
        txtPriceID, _
        txtPriceCustomerID, _
        txtPriceDestinationID, _
        mskPriceFrom, _
        mskPriceTo, _
        mskPriceAdultWithTransfer, _
        mskPriceKidWithTransfer, _
        mskPriceAdultWithoutTransfer, _
        mskPriceKidWithoutTransfer) Then
        blnStatus = False
        lngSelectedRow = grdPrices.CurRow
        blnInputingData = True
        DisableFields txtPriceCustomerDescription, txtPriceDestinationDescription
        EnableFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer, cmdHalf(0), cmdHalf(1)
        UpdateButtons Me, 6, 0, 1, 1, 0, 0, 1, 0
        mskPriceFrom.SetFocus
    End If

End Function

Private Function UpdateForm()

    If Not ValidateCriteria Then Exit Function
    
    If PopulateGrid Then
        DisableFields txtPriceCustomerDescription, txtPriceDestinationDescription
        EnableFields grdPrices
        grdPrices.SetCurCell 1, 4
        grdPrices.SetFocus
    End If
    
    UpdateButtons Me, 6, 1, 0, 0, 0, 0, 1, 0
    
End Function

Private Function ValidateCriteria()

    ValidateCriteria = False
    
    'Πελάτης
    If txtPriceCustomerID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtPriceCustomerDescription.SetFocus
        Exit Function
    End If

    'Προορισμός
    If txtPriceDestinationID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtPriceDestinationDescription.SetFocus
        Exit Function
    End If

    ValidateCriteria = True

End Function

Private Function ValidateFields()

    ValidateFields = False
    
    If Not IsDate(mskPriceFrom.text) Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskPriceFrom.SetFocus
        Exit Function
    End If
    If Not IsDate(mskPriceTo.text) Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskPriceTo.SetFocus
        Exit Function
    End If
    
    If CDate(mskPriceTo.text) < CDate(mskPriceFrom.text) Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(10), 1) Then
        End If
        mskPriceTo.SetFocus
        Exit Function
    End If
    
    If mskPriceAdultWithTransfer.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPriceAdultWithTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskPriceAdultWithTransfer.text) > 9999 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskPriceAdultWithTransfer.SetFocus
        Exit Function
    End If
    
    If mskPriceKidWithTransfer.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPriceKidWithTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskPriceKidWithTransfer.text) > 9999 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskPriceKidWithTransfer.SetFocus
        Exit Function
    End If
    
    If mskPriceAdultWithoutTransfer.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPriceAdultWithoutTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskPriceAdultWithoutTransfer.text) > 9999 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskPriceAdultWithoutTransfer.SetFocus
        Exit Function
    End If
    
    If mskPriceKidWithoutTransfer.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskPriceKidWithoutTransfer.SetFocus
        Exit Function
    End If
    If CCur(mskPriceKidWithoutTransfer.text) > 9999 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskPriceKidWithoutTransfer.SetFocus
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
            UpdateForm
        Case 4
            BatchProcessing
        Case 5
            AbortProcedure False
        Case 6
            AbortProcedure True
    End Select

End Sub

Private Sub cmdHalf_Click(index As Integer)

    'Ενήλικες
    If index = 0 And mskPriceAdultWithTransfer.text <> "" Then mskPriceKidWithTransfer.text = Format(CCur(mskPriceAdultWithTransfer.text / 2), "#,##0.00")
    'Παιδιά
    If index = 1 And mskPriceAdultWithoutTransfer.text <> "" Then mskPriceKidWithoutTransfer.text = Format(CCur(mskPriceAdultWithoutTransfer.text / 2), "#,##0.00")
    
End Sub

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Πελάτης - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "Description", "String", txtPriceCustomerDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "PriceID", "Επωνυμία", 0, 40, 1, 0)
            txtPriceCustomerID.text = tmpTableData.strCode
            txtPriceCustomerDescription.text = tmpTableData.strFirstField
        Case 1
            'Προορισμός - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtPriceDestinationDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "PriceID", "Περιγραφή", 0, 40, 1, 0)
            txtPriceDestinationID.text = tmpTableData.strCode
            txtPriceDestinationDescription.text = tmpTableData.strFirstField
    End Select

End Sub

Private Sub Form_Activate()
        
    If Me.Tag = "True" Then
        Me.Tag = "False"
        DisableFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer, cmdHalf(0), cmdHalf(1)
        EnableFields txtPriceCustomerDescription, txtPriceDestinationDescription
        AddColumnsToGrid grdPrices, 44, GetSetting(strAppTitle, "Layout Strings", "grdPrices"), "04NCIPriceID,04NCICompanyID,04NCIDestinationID,10NCDFrom,10NCDTo,10NRFXAdultWithTransfer,10NRFXKidWithTransfer,10NRFXAdultWithoutTransfer,10NRFXKidWithoutTransfer", "ID,CompanyID,DestinationID,Από,Έως,Ενήλικες Με Μεταφορά,Παιδιά Με Μεταφορά,Ενήλικες Χωρίς Μεταφορά,Παιδιά Χωρίς Μεταφορά"
        txtPriceCustomerDescription.SetFocus
    End If
    
    'AddDummyLines grdPrices, "99999", "99999", "99999", "Α99/99/9999Α", "Α99/99/9999Α", "999.99", "999.99", "999.99", "999.99"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF10 And cmdButton(3).Enabled, vbKeyC And CtrlDown And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyF7 And cmdButton(4).Enabled, vbKeyF And CtrlDown And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyEscape
            If cmdButton(5).Enabled Then cmdButton_Click 5: Exit Function
            If cmdButton(6).Enabled Then cmdButton_Click 6
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If blnInputingData Then
            If MyMsgBox(3, strAppTitle, strStandardMessages(3), 2) Then
                blnInputingData = False
                ClearFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer
                DisableFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer, cmdHalf(0), cmdHalf(1)
                grdPrices.SetFocus
                UpdateButtons Me, 6, 1, 0, 0, 0, 0, 1, 0
            End If
        Else
            ClearFields txtPriceCustomerDescription, txtPriceDestinationDescription, txtPriceCustomerID, txtPriceDestinationID, grdPrices
            EnableFields txtPriceCustomerDescription, txtPriceDestinationDescription
            UpdateButtons Me, 6, 0, 0, 0, 1, 1, 0, 1
            txtPriceCustomerDescription.SetFocus
            Exit Function
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Sub Form_Load()
        
    UpdateColors Me, False
    SetUpGrid lstIconList, grdPrices
    ClearFields txtPriceCustomerDescription, txtPriceDestinationDescription, mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer
    DisableFields mskPriceFrom, mskPriceTo, mskPriceAdultWithTransfer, mskPriceKidWithTransfer, mskPriceAdultWithoutTransfer, mskPriceKidWithoutTransfer, cmdHalf(0), cmdHalf(1)
    UpdateButtons Me, 6, 0, 0, 0, 1, 1, 0, 1
    
End Sub

Private Sub grdPrices_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub

Private Sub grdPrices_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPrices_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()
    
    SaveSetting strAppTitle, "Layout Strings", "grdPrices", grdPrices.LayoutCol

End Sub

Private Sub mskPriceAdultWithoutTransfer_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF6 Then cmdHalf_Click (1)

End Sub

Private Sub mskPriceAdultWithTransfer_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF6 Then cmdHalf_Click (0)
    
End Sub

Private Sub mskPriceKidWithoutTransfer_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF6 Then cmdHalf_Click (1)

End Sub

Private Sub mskPriceKidWithTransfer_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF6 Then cmdHalf_Click (0)

End Sub

Private Sub txtPriceCustomerDescription_Change()

    If txtPriceCustomerDescription.text = "" Then
        ClearFields txtPriceCustomerID
    End If
    
End Sub

Private Sub txtPriceCustomerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    
End Sub

Private Sub txtPriceCustomerDescription_Validate(Cancel As Boolean)

    If txtPriceCustomerID.text = "" And txtPriceCustomerDescription.text <> "" Then cmdIndex_Click 0
    
End Sub

Private Sub txtPriceDestinationDescription_Change()

    If txtPriceDestinationDescription.text = "" Then
        ClearFields txtPriceDestinationID
    End If

End Sub

Private Sub txtPriceDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtPriceDestinationDescription_Validate(Cancel As Boolean)

    If txtPriceDestinationID.text = "" And txtPriceDestinationDescription.text <> "" Then cmdIndex_Click 1

End Sub

