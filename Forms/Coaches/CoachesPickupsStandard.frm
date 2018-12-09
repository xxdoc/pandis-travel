VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form CoachesPickupsStandard 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   44
      Top             =   7350
      Width           =   8940
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   45
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
         Index           =   5
         Left            =   7350
         TabIndex        =   46
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "Εύρεση"
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
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   10350
      TabIndex        =   11
      Top             =   0
      Width           =   4515
      Begin VB.TextBox txtDestinationID 
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1950
         Width           =   780
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Transfers.TransferDestinationID"
         Top             =   1950
         Width           =   3540
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
      Begin VB.TextBox Text6 
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
         Text            =   "Transfers.TransferRouteID"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox txtPickupPointID 
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox txtTransferTypeID 
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox txtCustomerID 
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
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox txtTransferID 
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   75
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
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Transfers.CustomerID"
         Top             =   450
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
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Transfers.TransferTransferTypeID"
         Top             =   825
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
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Transfers.TransferPickupPointID"
         Top             =   1200
         Width           =   3540
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Transfers.TransferID"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Text            =   "RefersTo"
         Top             =   2325
         Width           =   3540
      End
      Begin VB.TextBox txtRefersTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Top             =   2325
         Width           =   780
      End
   End
   Begin UserControls.newDate mskDate 
      Height          =   465
      Left            =   2250
      TabIndex        =   0
      Top             =   1125
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
   Begin UserControls.newText txtCustomerDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   1
      Top             =   1650
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
   Begin UserControls.newText txtTransferTypeDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   2
      Top             =   2175
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
      Left            =   2250
      TabIndex        =   3
      Top             =   2700
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
   Begin UserControls.newText txtPickupPointDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   4
      Top             =   3225
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   50
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
   Begin UserControls.newText txtRemarks 
      Height          =   465
      Left            =   2250
      TabIndex        =   10
      Top             =   6375
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
   Begin UserControls.newText txtRouteDescription 
      Height          =   465
      Left            =   3000
      TabIndex        =   6
      Top             =   3750
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
   Begin UserControls.newText txtRouteShortDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   5
      Top             =   3750
      Width           =   690
      _ExtentX        =   1217
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
   Begin UserControls.newInteger mskAdults 
      Height          =   465
      Left            =   2250
      TabIndex        =   7
      Top             =   4275
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   3
      Text            =   "999"
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
   Begin UserControls.newInteger mskKids 
      Height          =   465
      Left            =   2250
      TabIndex        =   8
      Top             =   4800
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   3
      Text            =   "999"
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
   Begin UserControls.newInteger mskFree 
      Height          =   465
      Left            =   2250
      TabIndex        =   9
      Top             =   5325
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   3
      Text            =   "999"
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
   Begin UserControls.newInteger mskTotal 
      Height          =   465
      Left            =   2250
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5850
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   820
      Enabled         =   0   'False
      ForeColor       =   0
      Text            =   "999"
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   7275
      TabIndex        =   39
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
      PicNormal       =   "CoachesPickupsStandard.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7275
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2175
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
      PicNormal       =   "CoachesPickupsStandard.frx":059A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   7275
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2700
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
      PicNormal       =   "CoachesPickupsStandard.frx":0B34
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   7275
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3225
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
      PicNormal       =   "CoachesPickupsStandard.frx":10CE
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   5
      Left            =   8025
      TabIndex        =   43
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
      PicNormal       =   "CoachesPickupsStandard.frx":1668
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   3750
      TabIndex        =   51
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
      PicNormal       =   "CoachesPickupsStandard.frx":1C02
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   0
      Left            =   1800
      Top             =   2625
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   3825
      Top             =   6825
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   12
      Left            =   0
      Top             =   2100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   9375
      Top             =   6900
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3975
      Top             =   8025
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Μεταφορές επιβατών"
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
      TabIndex        =   38
      Top             =   75
      Width           =   4995
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Παρατηρήσεις"
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
      Index           =   14
      Left            =   450
      TabIndex        =   37
      Top             =   6450
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Σύνολο"
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
      TabIndex        =   36
      Top             =   5925
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Δωρεάν"
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
      TabIndex        =   35
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Index           =   11
      Left            =   450
      TabIndex        =   34
      Top             =   4875
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Index           =   10
      Left            =   450
      TabIndex        =   33
      Top             =   4350
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Διαδρομή"
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
      TabIndex        =   32
      Top             =   3825
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Σημείο παραλαβής"
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
      TabIndex        =   31
      Top             =   3300
      Width           =   1365
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
      Index           =   7
      Left            =   450
      TabIndex        =   30
      Top             =   2775
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Τύπος μετακίνησης"
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
      TabIndex        =   29
      Top             =   2250
      Width           =   1365
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
      Index           =   5
      Left            =   450
      TabIndex        =   28
      Top             =   1725
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
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
      Index           =   2
      Left            =   450
      TabIndex        =   27
      Top             =   1200
      Width           =   1365
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
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   13
      Left            =   2700
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "CoachesPickupsStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnStatus As Boolean
Dim blnCancel As Boolean

Private Function AbortProcedure(blnStatus)

    'If monthlyCalendar.Visible Then monthlyCalendar.Visible = False: Exit Function

    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            blnCancel = True
            ClearFields txtTransferID, txtCustomerID, txtTransferTypeID, txtPickupPointID, txtRouteID, txtDestinationID
            ClearFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
            'ClearFields lblWeekday, mskTotal
            ClearFields mskTotal
            DisableFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
            UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("ΜεταφορέςΗμερολόγιο"), 0, 1), 0, 1
        End If
    End If
    
    If blnStatus Then Unload Me
    
End Function

Private Function DeleteRecord()

    If MainDeleteRecord("CommonDB", "Transfers", strApplicationName, "ID", txtTransferID.text, "True") Then
        blnCancel = True
        ClearFields txtTransferID, txtCustomerID, txtTransferTypeID, txtPickupPointID, txtRouteID, txtDestinationID
        ClearFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
        'ClearFields lblWeekday, mskTotal
        ClearFields mskTotal
        DisableFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("ΜεταφορέςΗμερολόγιο"), 0, 1), 0, 1
    End If

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
            FindRecords
        Case 4
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select
    
End Sub

Private Function FindRecords()

    With CoachesReport
        .Tag = "True"
        .txtRefersTo.text = txtRefersTo.text
        .txtCallingForm.text = "Transactions"
        .grdCoachesReport.Tag = "grdCoachesReportStandard"
        .Show 1, Me
    End With

End Function

Private Function SaveRecord()

    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "Transfers", blnStatus, strApplicationName, "ID", txtTransferID.text, mskDate.text, txtCustomerID.text, txtTransferTypeID.text, txtPickupPointID.text, txtRouteID.text, mskAdults.text, mskKids.text, mskFree.text, txtDestinationID.text, "", txtRemarks.text, 1, strCurrentUser) <> 0 Then
        blnCancel = True
        ClearFields txtTransferID, txtCustomerID, txtTransferTypeID, txtPickupPointID, txtRouteID, txtDestinationID
        ClearFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
        'ClearFields lblWeekday, mskTotal
        ClearFields mskTotal
        DisableFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If

End Function

Private Function ValidateFields()
    
    ValidateFields = False
    
    'Ημερομηνία
    If mskDate.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskDate.SetFocus
        Exit Function
    End If
    If Not IsDate(mskDate.text) Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskDate.SetFocus
        Exit Function
    End If
    
    'Πελάτης
    If txtCustomerID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtCustomerDescription.SetFocus
        Exit Function
    End If

    'Τύπος μετακίνησης
    If txtTransferTypeID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtTransferTypeDescription.SetFocus
        Exit Function
    End If

    'Σημείο παραλαβής
    If txtPickupPointID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPickupPointDescription.SetFocus
        Exit Function
    End If
    
    'Δρομολόγιο
    If txtRouteID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtRouteDescription.SetFocus
        Exit Function
    End If
    
    'Ενήλικες
    If mskAdults.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskAdults.SetFocus
        Exit Function
    End If
    
    'Παιδιά
    If mskKids.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskKids.SetFocus
        Exit Function
    End If
    
    'Δωρεάν
    If mskFree.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskFree.SetFocus
        Exit Function
    End If
    
    'Προορισμός
    If txtDestinationID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtDestinationDescription.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Function NewRecord()

    blnStatus = True
    blnCancel = False
    ClearFields txtTransferID, txtCustomerID, txtTransferTypeID, txtPickupPointID, txtRouteID, txtDestinationID
    ClearFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
    'ClearFields lblWeekday, mskTotal
    ClearFields mskTotal
    EnableFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    UpdateButtons Me, 5, 0, 1, 0, 0, 1, 0
    mskDate.SetFocus
        
    InitializeFields mskAdults, mskKids, mskFree, mskTotal
        
End Function

Private Sub cmdIndex_Click(index As Integer)

    Dim strShowInList As String
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    Dim strSQL As String
    Dim intSize As Integer

    Select Case index
        Case 0
            'Ημερολόγιο
            'ShowMonthlyCalendar Me, monthlyCalendar
        Case 1
            'Πελάτης
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "Description", "String", txtCustomerDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtCustomerID.text = tmpTableData.strCode
                txtCustomerDescription.text = tmpTableData.strFirstField
            End If
        Case 2
            'Τύπος μετακίνησης
            Set tmpRecordset = CheckForMatch("CommonDB", "TransferTypes", "TransferTypeDescription", "String", txtTransferTypeDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtTransferTypeID.text = tmpTableData.strCode
                txtTransferTypeDescription.text = tmpTableData.strFirstField
            End If
        Case 3
            'Προορισμός
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestinationDescription.text = tmpTableData.strFirstField
            End If
        Case 4
            'Δρομολόγιο - Αν έχω δώσει προορισμό, βρίσκω τα δρομολόγια που είναι συνδεδεμένα με τον δοσμένο προορισμό
            If txtDestinationID.text <> "" Then
                strSQL = "SELECT DISTINCT DestinationsRoutesPickupPoints.RouteID, DestinationID, RouteShortDescription, RouteDescription " _
                    & "FROM DestinationsRoutesPickupPoints " _
                    & "INNER JOIN PickupRoutes ON DestinationsRoutesPickupPoints.RouteID = PickupRoutes.RouteID " _
                    & "WHERE DestinationID = " & txtDestinationID.text & " " _
                    & IIf(txtRouteShortDescription.text <> "", "AND RouteShortDescription = '" & txtRouteShortDescription.text & "'", "")
                Set tmpRecordset = FindAndReturnRecords(strSQL)
                If tmpRecordset.RecordCount > 0 Then
                    tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 2, 3, "ID", "Συντ.", "Δρομολόγιο", 0, 5, 40, 1, 1, 0)
                    txtRouteID.text = tmpTableData.strCode
                    txtRouteShortDescription.text = tmpTableData.strFirstField
                    txtRouteDescription.text = tmpTableData.strSecondField
                End If
            Else
                ClearFields txtRouteID, txtRouteShortDescription, txtRouteDescription
            End If
        Case 5
            'Σημείο παραλαβής - έχω δώσει προορισμό - έχω δώσει δρομολόγιο
            If txtDestinationID.text <> "" And txtRouteID.text <> "" Then
                intSize = Len(txtPickupPointDescription.text)
                strSQL = "SELECT DestinationID, RouteID, DestinationsRoutesPickupPoints.PickupPointID, PickupPointHotelDescription, PickupPointTime " _
                    & "FROM DestinationsRoutesPickupPoints " _
                    & "INNER JOIN PickupPoints ON DestinationsRoutesPickupPoints.PickupPointID = PickupPoints.PickupPointID " _
                    & "WHERE RouteID = " & txtRouteID.text & " AND DestinationID = " & txtDestinationID.text & " " _
                    & IIf(txtPickupPointDescription.text <> "", "AND Left(PickupPointHotelDescription, " & intSize & ") = '" & txtPickupPointDescription.text & "'", "")
                Set tmpRecordset = FindAndReturnRecords(strSQL)
                If tmpRecordset.RecordCount > 0 Then
                    tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 2, 3, 4, "ID", "Περιγραφή", "Ωρα", 0, 40, 7, 1, 0, 1)
                    txtPickupPointID.text = tmpTableData.strCode
                    txtPickupPointDescription.text = tmpTableData.strFirstField
                End If
            End If
            'Σημείο παραλαβής - έχω δώσει προορισμό - δεν έχω δώσει δρομολόγιο
            If txtDestinationID.text <> "" And txtRouteID.text = "" Then
                intSize = Len(txtPickupPointDescription.text)
                strSQL = "SELECT DestinationID, RouteID, DestinationsRoutesPickupPoints.PickupPointID, PickupPointHotelDescription, PickupPointTime " _
                    & "FROM DestinationsRoutesPickupPoints " _
                    & "INNER JOIN PickupPoints ON DestinationsRoutesPickupPoints.PickupPointID = PickupPoints.PickupPointID " _
                    & "WHERE DestinationID = " & txtDestinationID.text & " " _
                    & IIf(txtPickupPointDescription.text <> "", "AND Left(PickupPointHotelDescription, " & intSize & ") = '" & txtPickupPointDescription.text & "'", "")
                Set tmpRecordset = FindAndReturnRecords(strSQL)
                If tmpRecordset.RecordCount > 0 Then
                    tmpTableData = DisplayIndex(tmpRecordset, 2, True, 4, 1, 2, 3, 4, "ID", "RouteID", "Περιγραφή", "Ωρα", 0, 0, 40, 7, 1, 0, 0, 1)
                    txtPickupPointID.text = tmpTableData.strFirstField
                    txtPickupPointDescription.text = tmpTableData.strSecondField
                    txtRouteID.text = tmpTableData.strCode
                    FindRoute
                End If
            End If
    End Select

End Sub

Private Function FindRoute()

    Dim rsTable As Recordset
    
    Set rsTable = CommonDB.OpenRecordset("PickupRoutes")
    With rsTable
        .index = "PickupRouteID"
        .Seek "=", Val(txtRouteID.text)
        If Not .NoMatch Then
            txtRouteID.text = !PickupRouteID
            txtRouteShortDescription.text = !PickupRouteShortDescription
            txtRouteDescription.text = !PickupRouteDescription
            txtRouteShortDescription.Locked = True
        Else
            txtRouteID.text = ""
            txtRouteShortDescription.text = ""
            txtRouteDescription.text = ""
            txtRouteShortDescription.Locked = False
        End If
        .Close
    End With

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
        
End Sub

Public Function SeekRecord(myTransferID, myRefersTo)

    Dim tmpRecordset As Recordset
    Dim tmpTableData As typTableData
    
    ClearFields txtTransferID, txtCustomerID, txtTransferTypeID, txtPickupPointID, txtRouteID, txtDestinationID
    ClearFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
    'ClearFields lblWeekday, mskTotal
    ClearFields mskTotal
    DisableFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    
    SeekRecord = False
    
    If MainSeekRecord("CommonDB", "Transfers", "ID", myTransferID, True, txtTransferID, mskDate, txtCustomerID, txtTransferTypeID, txtPickupPointID, txtRouteID, mskAdults, mskKids, mskFree, txtDestinationID, txtPickupPointDescription, txtRemarks) Then
        'Πελάτης
        Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "ID", "Numeric", txtCustomerID.text)
        txtCustomerID.text = tmpRecordset.Fields(0)
        txtCustomerDescription.text = tmpRecordset.Fields(1)
        'Τύπος μετακίνησης
        Set tmpRecordset = CheckForMatch("CommonDB", "TransferTypes", "TransferTypeID", "Numeric", txtTransferTypeID.text)
        txtTransferTypeID.text = tmpRecordset.Fields(0)
        txtTransferTypeDescription.text = tmpRecordset.Fields(1)
        'Προορισμός - 1 από 3
        Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationID", "Numeric", txtDestinationID.text)
        txtDestinationID.text = tmpRecordset.Fields(0)
        txtDestinationDescription.text = tmpRecordset.Fields(2)
        'Δρομολόγιο - 2 από 3
        Set tmpRecordset = CheckForMatch("CommonDB", "PickupRoutes", "PickupRouteID", "Numeric", txtRouteID.text)
        txtRouteID.text = tmpRecordset.Fields(0)
        txtRouteShortDescription.text = tmpRecordset.Fields(1)
        txtRouteDescription.text = tmpRecordset.Fields(2)
        'Σημείο παραλαβής - 3 από 3
        Set tmpRecordset = CheckForMatch("CommonDB", "PickupPoints", "PickupPointID", "Numeric", txtPickupPointID.text)
        txtPickupPointID.text = tmpRecordset.Fields(0)
        txtPickupPointDescription.text = tmpRecordset.Fields(2)
        'Τα υπόλοιπα
        EnableFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
        mskTotal.text = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
        'lblWeekday.Caption = FindWeekDay(mskDate.text)
        txtRefersTo.text = myRefersTo
        blnCancel = False
        blnStatus = False
        SeekRecord = True
        UpdateButtons Me, 5, 0, 1, 1, 0, 1, 0
    End If
    
End Function

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyInsert And cmdButton(0).Enabled, vbKeyN And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyF7 And cmdButton(3).Enabled, vbKeyF And CtrlDown = 4 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyEscape
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Function
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12 And CtrlDown
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    UpdateColors Me, False
    blnCancel = True
    ClearFields txtTransferID, txtCustomerID, txtTransferTypeID, txtPickupPointID, txtRouteID, txtDestinationID
    ClearFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
    'ClearFields lblWeekday, mskTotal
    ClearFields mskTotal
    DisableFields mskDate, txtCustomerDescription, txtTransferTypeDescription, txtDestinationDescription, txtPickupPointDescription, txtRouteShortDescription, txtRouteDescription, mskAdults, mskKids, mskFree, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5)
    'monthlyCalendar.Visible = False
    UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1

End Sub

Private Sub monthlyCalendar_DblClick()

    'mskDate.text = format(monthlyCalendar.Value, "dd/mm/yyyy")
    'monthlyCalendar.Visible = False

End Sub

Private Sub monthlyCalendar_KeyPress(KeyAscii As Integer)

    'If KeyAscii = vbKeyReturn Then
    '    mskDate.text = format(monthlyCalendar.Value, "dd/mm/yyyy")
    '    monthlyCalendar.Visible = False
    'End If

End Sub

Private Sub mskDate_GotFocus()

    'lblWeekday.Caption = FindWeekDay(mskDate.text)
    
End Sub

Private Sub mskDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    
End Sub

Private Sub mskDate_LostFocus()

    'lblWeekday.Caption = FindWeekDay(mskDate.text)

End Sub

Private Sub mskDate_Validate(Cancel As Boolean)

    'lblWeekday.Caption = FindWeekDay(mskDate.text)
    
End Sub

Private Sub txtCustomerDescription_Change()

    If txtCustomerDescription.text = "" Then
        ClearFields txtCustomerID
    End If

End Sub

Private Sub txtCustomerDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtCustomerDescription_Validate(Cancel As Boolean)

    If txtCustomerID.text = "" And txtCustomerDescription.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then
        ClearFields txtDestinationID
    End If

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub

Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 3

End Sub

Private Sub txtPickupPointDescription_Change()

    If txtPickupPointDescription.text = "" Then
        ClearFields txtPickupPointID
        txtRouteShortDescription.Locked = False
    End If

End Sub

Private Sub txtPickupPointDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 5

End Sub

Private Sub txtPickupPointDescription_Validate(Cancel As Boolean)

    If txtPickupPointID.text = "" And txtPickupPointDescription.text <> "" Then cmdIndex_Click 5

End Sub

Private Sub txtRouteDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3

End Sub

Private Sub txtRouteShortDescription_Change()

    If txtRouteShortDescription.text = "" Then
        ClearFields txtRouteID, txtRouteDescription, txtPickupPointID, txtPickupPointDescription
    End If

End Sub

Private Sub txtRouteShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4

End Sub

Private Sub txtRouteShortDescription_Validate(Cancel As Boolean)

    If txtRouteID.text = "" And txtRouteShortDescription.text <> "" Then cmdIndex_Click 4

End Sub

Private Sub mskAdults_Validate(Cancel As Boolean)

    If Not blnCancel Then
        mskTotal.text = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
    End If

End Sub

Private Sub mskFree_Validate(Cancel As Boolean)

    If Not blnCancel Then
        mskTotal.text = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
    End If

End Sub

Private Sub mskKids_Validate(Cancel As Boolean)
    
    If Not blnCancel Then
        mskTotal.text = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
    End If

End Sub

Private Sub txtTransferTypeDescription_Change()

    If txtTransferTypeDescription.text = "" Then
        ClearFields txtTransferTypeID
    End If

End Sub

Private Sub txtTransferTypeDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtTransferTypeDescription_Validate(Cancel As Boolean)

    If txtTransferTypeID.text = "" And txtTransferTypeDescription.text <> "" Then cmdIndex_Click 2

End Sub

Private Function FindAndReturnRecords(strSQL) As Recordset

   Dim tmpRecordset As Recordset
   
   Set tmpRecordset = CommonDB.OpenRecordset(strSQL, dbOpenSnapshot)
   Set FindAndReturnRecords = tmpRecordset

End Function

