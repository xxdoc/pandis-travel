VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form CoachesPickupsBrief 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15000
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
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
      TabIndex        =   35
      Top             =   6300
      Width           =   8940
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
   Begin MSComCtl2.MonthView monthlyCalendar 
      Height          =   2370
      Left            =   12075
      TabIndex        =   34
      Top             =   2100
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   25231362
      CurrentDate     =   43031
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   10350
      TabIndex        =   19
      Top             =   75
      Width           =   4515
      Begin VB.TextBox txtDestinationShortDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Destinations.DestinationShortDescription"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtRefersTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
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
         Text            =   "RefersTo"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox Text4 
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
         Text            =   "Transfers.TransferID"
         Top             =   75
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
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Transfers.CustomerID"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox txtTransferID 
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
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
         TabIndex        =   22
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
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Transfers.DestinationID"
         Top             =   825
         Width           =   3540
      End
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
   End
   Begin UserControls.newDate mskDate 
      Height          =   465
      Left            =   2250
      TabIndex        =   1
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
   Begin UserControls.newText txtCompanyDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   2
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
   Begin UserControls.newText txtDestinationDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   3
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
   Begin UserControls.newText txtRemarks 
      Height          =   465
      Left            =   2250
      TabIndex        =   8
      Top             =   5325
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
   Begin UserControls.newInteger mskAdults 
      Height          =   465
      Left            =   2250
      TabIndex        =   5
      Top             =   3225
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
      TabIndex        =   6
      Top             =   3750
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
   Begin UserControls.newInteger mskTotal 
      Height          =   465
      Left            =   2250
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4800
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   820
      Enabled         =   0   'False
      ForeColor       =   0
      Text            =   "9.999"
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
   Begin UserControls.newText txtPickupPointDescription 
      Height          =   465
      Left            =   2250
      TabIndex        =   4
      Top             =   2700
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   820
      ForeColor       =   0
      MaxLength       =   50
      Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
      Left            =   3750
      TabIndex        =   31
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
      PicNormal       =   "CoachesPickupsBrief.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   7275
      TabIndex        =   32
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
      PicNormal       =   "CoachesPickupsBrief.frx":059A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7275
      TabIndex        =   33
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
      PicNormal       =   "CoachesPickupsBrief.frx":0B34
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   4350
      Top             =   6975
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   4050
      Top             =   5775
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   9375
      Top             =   6000
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
      Left            =   1800
      Top             =   2550
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
      Top             =   2250
      Visible         =   0   'False
      Width           =   465
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
      TabIndex        =   28
      Top             =   2775
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
      TabIndex        =   18
      Top             =   1200
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
      TabIndex        =   17
      Top             =   1725
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
      TabIndex        =   16
      Top             =   2250
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
      TabIndex        =   15
      Top             =   3300
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
      TabIndex        =   14
      Top             =   3825
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
      TabIndex        =   13
      Top             =   4350
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
      TabIndex        =   12
      Top             =   4875
      Width           =   1365
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
      TabIndex        =   11
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label lblWeekday 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ημέρα"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   1200
      Width           =   465
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
      TabIndex        =   0
      Top             =   75
      Width           =   4995
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
      Left            =   2550
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "CoachesPickupsBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnStatus As Boolean
Dim blnCancel As Boolean

Private Function AbortProcedure(blnStatus)

    If monthlyCalendar.Visible Then monthlyCalendar.Visible = False: Exit Function

    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strStandardMessages(3), 2) Then
            blnStatus = False
            blnCancel = True
            ClearFields txtTransferID, txtCustomerID, txtTransferDestinationID, txtDestinationShortDescription
            ClearFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
            ClearFields lblWeekday, mskTotal
            DisableFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
            UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("CoachesReport"), 0, 1), 0, 1
        End If
    End If
    
    If blnStatus Then Unload Me
    
End Function

Private Function DeleteRecord()

    If MainDeleteRecord("CommonDB", "Transfers", strAppTitle, "ID", txtTransferID.text, "True") Then
        blnCancel = True
        ClearFields txtTransferID, txtCustomerID, txtTransferDestinationID, txtDestinationShortDescription
        ClearFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
        ClearFields lblWeekday, mskTotal
        DisableFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
        UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("CoachesReport"), 0, 1), 0, 1
    End If

End Function

Private Function DisplayRecordID(myDestinationDescription, myTransferDate, myRecordID)

     MyMsgBox 1, strAppTitle, strAppMessages(8) & " " & CreateReferenceNo(myDestinationDescription, myTransferDate, myRecordID), 1

End Function

Private Sub cmdButton_Click(Index As Integer)
                
    Select Case Index
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
        .grdCoachesReport.Tag = "grdCoachesReportBrief"
        .Show 1, Me
    End With
    
End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    txtTransferID.text = MainSaveRecord("CommonDB", "Transfers", blnStatus, strAppTitle, "ID", txtTransferID.text, mskDate.text, txtCustomerID.text, 0, 0, 0, mskAdults.text, mskKids.text, mskFree.text, txtTransferDestinationID.text, txtPickupPointDescription.text, txtRemarks.text, "2", strCurrentUser)
    
    If txtTransferID.text <> 0 Then
        blnCancel = True
        DisplayRecordID txtDestinationShortDescription.text, mskDate.text, txtTransferID.text
        ClearFields txtTransferID, txtCustomerID, txtTransferDestinationID, txtDestinationShortDescription
        ClearFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
        ClearFields lblWeekday, mskTotal
        DisableFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
        UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("CoachesReport"), 0, 1), 0, 1
    End If
    
End Function

Private Function ValidateFields()
    
    ValidateFields = False
    
    'Ημερομηνία
    If mskDate.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskDate.SetFocus
        Exit Function
    End If
    If Not IsDate(mskDate.text) Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskDate.SetFocus
        Exit Function
    End If
    
    'Πελάτης
    If txtCustomerID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtCompanyDescription.SetFocus
        Exit Function
    End If

    'Προορισμός
    If txtTransferDestinationID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtDestinationDescription.SetFocus
        Exit Function
    End If
    
    'Ενήλικες
    If mskAdults.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskAdults.SetFocus
        Exit Function
    End If
    
    'Παιδιά
    If mskKids.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskKids.SetFocus
        Exit Function
    End If
    
    'Δωρεάν
    If mskFree.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        mskFree.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Function NewRecord()

    blnStatus = True
    blnCancel = False
    ClearFields txtTransferID, txtCustomerID, txtTransferDestinationID, txtDestinationShortDescription
    ClearFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
    ClearFields lblWeekday, mskTotal
    EnableFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
    UpdateButtons Me, 5, 0, 1, 0, 0, 1, 0
    mskDate.SetFocus
        
    InitializeFields mskAdults, mskKids, mskFree, mskTotal
        
End Function

Private Sub cmdIndex_Click(Index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset

    Select Case Index
        Case 0
            'Ημερολόγιο
            ShowMonthlyCalendar Me, monthlyCalendar
        Case 1
            'Πελάτης
            Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "Description", "String", txtCompanyDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 5, 40, 1, 0)
            txtCustomerID.text = tmpTableData.strCode
            txtCompanyDescription.text = tmpTableData.strFirstField
        Case 2
            'Προορισμός
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestinationDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 2, "ID", "Συντ.", "Περιγραφή", 0, 5, 40, 1, 1, 0)
            txtTransferDestinationID.text = tmpTableData.strCode
            txtDestinationShortDescription.text = tmpTableData.strFirstField
            txtDestinationDescription.text = tmpTableData.strSecondField
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
        
End Sub

Public Function SeekRecord(myTransferID, myRefersTo)

    Dim tmpRecordset As Recordset
    Dim tmpTableData As typTableData
    
    ClearFields txtTransferID, txtCustomerID, txtTransferDestinationID, txtDestinationShortDescription
    ClearFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
    ClearFields lblWeekday, mskTotal
    DisableFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
    
    SeekRecord = False
    
    If MainSeekRecord("CommonDB", "Transfers", "ID", myTransferID, True, txtTransferID, mskDate, txtCustomerID, , , , mskAdults, mskKids, mskFree, txtTransferDestinationID, txtPickupPointDescription, txtRemarks) Then
        'Πελάτης
        Set tmpRecordset = CheckForMatch("CommonDB", "Customers", "ID", "Numeric", txtCustomerID.text)
        txtCustomerID.text = tmpRecordset.Fields(0)
        txtCompanyDescription.text = tmpRecordset.Fields(1)
        'Προορισμός
        Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationID", "Numeric", txtTransferDestinationID.text)
        txtTransferDestinationID.text = tmpRecordset.Fields(0)
        txtDestinationShortDescription.text = tmpRecordset.Fields(1)
        txtDestinationDescription.text = tmpRecordset.Fields(2)
        'Τα υπόλοιπα
        EnableFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
        mskTotal.text = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
        lblWeekday.Caption = FindWeekDay(mskDate.text)
        txtRefersTo.text = myRefersTo
        blnCancel = False
        blnStatus = False
        SeekRecord = True
        UpdateButtons Me, 5, 0, 1, 1, 0, 1, 0
     End If
    
End Function

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
        Case vbKeyF7 And cmdButton(3).Enabled, vbKeyF And CtrlDown And cmdButton(3).Enabled
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
    ClearFields txtTransferID, txtCustomerID, txtTransferDestinationID, txtDestinationShortDescription
    ClearFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
    ClearFields lblWeekday, mskTotal
    DisableFields mskDate, txtCompanyDescription, txtDestinationDescription, txtPickupPointDescription, mskAdults, mskKids, mskFree, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2)
    monthlyCalendar.Visible = False
    UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
   
End Sub

Private Sub monthlyCalendar_DblClick()

    mskDate.text = Format(monthlyCalendar.Value, "dd/mm/yyyy")
    monthlyCalendar.Visible = False

End Sub

Private Sub monthlyCalendar_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        mskDate.text = Format(monthlyCalendar.Value, "dd/mm/yyyy")
        monthlyCalendar.Visible = False
    End If

End Sub

Private Sub mskAdults_Validate(Cancel As Boolean)

    If Not blnCancel Then
        mskTotal.text = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
    End If

End Sub

Private Sub mskDate_GotFocus()

    lblWeekday.Caption = FindWeekDay(mskDate.text)

End Sub

Private Sub mskDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    
End Sub

Private Sub mskDate_LostFocus()

    lblWeekday.Caption = FindWeekDay(mskDate.text)
    
End Sub

Private Sub mskDate_Validate(Cancel As Boolean)

    lblWeekday.Caption = FindWeekDay(mskDate.text)
    
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

Private Sub mskKids_Change()

    If Not blnCancel Then
        mskTotal.text = AddNumbers(mskAdults.text, mskKids.text, mskFree.text)
    End If

End Sub

Private Sub txtCompanyDescription_Change()

    If txtCompanyDescription.text = "" Then
        ClearFields txtCustomerID
    End If

End Sub

Private Sub txtCompanyDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtCompanyDescription_Validate(Cancel As Boolean)

    If txtCustomerID.text = "" And txtCompanyDescription.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtDestinationDescription_Change()

    If txtDestinationDescription.text = "" Then
        ClearFields txtTransferDestinationID, txtDestinationShortDescription
    End If

End Sub

Private Sub txtDestinationDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtDestinationDescription_Validate(Cancel As Boolean)

    If txtTransferDestinationID.text = "" And txtDestinationDescription.text <> "" Then cmdIndex_Click 2

End Sub

Private Function FindAndReturnRecords(strSQL) As Recordset

   Dim tmpRecordset As Recordset
   
   Set tmpRecordset = CommonDB.OpenRecordset(strSQL, dbOpenSnapshot)
   Set FindAndReturnRecords = tmpRecordset

End Function

