VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form Persons 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   18150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   33
      Top             =   6825
      Width           =   8940
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "ƒÁÏÈÔıÒ„ﬂ·"
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   8421631
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   " ÎÂﬂÛÈÏÔ"
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "¡ÔËﬁÍÂıÛÁ"
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "ƒÈ·„Ò·ˆﬁ"
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "≈˝ÒÂÛÁ"
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
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
         ButtonShape     =   3
         ButtonStyle     =   4
         Caption         =   "¡ÍıÒÔ"
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
      Caption         =   "Customer"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   10200
      TabIndex        =   21
      Top             =   2475
      Width           =   4515
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
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "999"
         Top             =   75
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
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "InvoiceMasterRefersTo"
         Top             =   75
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
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
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "CustomersOrSuppliers"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox txtPersonVATStateID 
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
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
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "PersonVATStateID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtPersonTaxOfficeID 
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
         Top             =   825
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
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "PersonTaxOfficeID"
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
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "PersonID"
         Top             =   450
         Width           =   3540
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
   End
   Begin UserControls.newText txtDescription 
      Height          =   465
      Left            =   2625
      TabIndex        =   0
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
   Begin UserControls.newText txtPhones 
      Height          =   465
      Left            =   2625
      TabIndex        =   5
      Top             =   3750
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
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
   Begin UserControls.newText txtAddress 
      Height          =   465
      Left            =   2625
      TabIndex        =   4
      Top             =   3225
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   50
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
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
   Begin UserControls.newText txtTaxNo 
      Height          =   465
      Left            =   2625
      TabIndex        =   1
      Top             =   1650
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   15
      Text            =   "AAAAAAAAAAAAAAA"
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
   Begin UserControls.newText txtVATStateDescription 
      Height          =   465
      Left            =   2625
      TabIndex        =   7
      Top             =   4800
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
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
   Begin UserControls.newText txtTaxOfficeDescription 
      Height          =   465
      Left            =   2625
      TabIndex        =   2
      Top             =   2175
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
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
   Begin UserControls.newText txtAccountCode 
      Height          =   465
      Left            =   2625
      TabIndex        =   9
      Top             =   5850
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   15
      Text            =   "AAAAAAAAAAAAAAA"
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
   Begin UserControls.newText txtPersonInCharge 
      Height          =   465
      Left            =   2625
      TabIndex        =   6
      Top             =   4275
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
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
   Begin UserControls.newText txtProfession 
      Height          =   465
      Left            =   2625
      TabIndex        =   3
      Top             =   2700
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
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
   Begin UserControls.newText txtEmail 
      Height          =   465
      Left            =   2625
      TabIndex        =   8
      Top             =   5325
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   820
      ForeColor       =   4194304
      MaxLength       =   40
      Text            =   "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
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
      Left            =   7650
      TabIndex        =   28
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
      PicNormal       =   "Persons.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   8100
      TabIndex        =   29
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
      PicNormal       =   "Persons.frx":059A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7650
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4800
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
      PicNormal       =   "Persons.frx":0B34
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   1
      Left            =   0
      Top             =   7050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "”ıÌ·ÎÎ·Û¸ÏÂÌÔÚ"
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
      TabIndex        =   19
      Top             =   75
      Width           =   3870
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   9375
      Top             =   6375
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3600
      Top             =   7500
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   3000
      Top             =   6300
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   0
      Left            =   2175
      Top             =   900
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
      Left            =   3675
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
      Top             =   2100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "E-Mail"
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
      TabIndex        =   20
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   " ·ËÂÛÙ˛Ú ÷.–.¡."
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
      TabIndex        =   18
      Top             =   4875
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   " ˘‰. √ÂÌ. ÀÔ„ÈÛÙÈÍﬁÚ"
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
      TabIndex        =   17
      Top             =   5925
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "œÈÍÔÌÔÏÈÍﬁ ıÁÒÂÛﬂ·"
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
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "¡.÷.Ã."
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
      TabIndex        =   15
      Top             =   1725
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "ƒÈÂ˝ËıÌÛÁ"
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
      Top             =   3300
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "ƒÒ·ÛÙÁÒÈ¸ÙÁÙ·"
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
      Top             =   2775
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "‘ÁÎ›ˆ˘Ì·"
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
      TabIndex        =   12
      Top             =   3825
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "’Â˝ËıÌÔÚ ÂÈÍÔÈÌ˘Ìﬂ·Ú"
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
      TabIndex        =   11
      Top             =   4350
      Width           =   1740
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "≈˘ÌıÏﬂ·"
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
      Top             =   1200
      Width           =   1740
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
      Begin VB.Menu mnu¡ÔËﬁÍÂıÛÁ–Î‹ÙÔıÚ”ÙÁÎ˛Ì 
         Caption         =   "¡ÔËﬁÍÂıÛÁ Î‹ÙÔıÚ ÛÙÁÎ˛Ì"
      End
   End
End
Attribute VB_Name = "Persons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean

Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtPersonID, txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode
            DisableFields txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode, cmdIndex(0), cmdIndex(1), cmdIndex(2)
            UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Public Function SeekRecord(myID)

    Dim blnEnableDelete As Boolean
    Dim tmpRecordset As Recordset
    Dim tmpTableData As typTableData
    
    ClearFields txtPersonID, txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode
    DisableFields txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode, cmdIndex(0), cmdIndex(1), cmdIndex(2)
    
    SeekRecord = False
    
    blnEnableDelete = SimpleSeek("Invoices", "InvoicePersonIDAndInvoiceMasterRefersTo", myID, txtInvoiceMasterRefersTo.text)
    If blnEnableDelete Then blnEnableDelete = SimpleSeek("Invoices", "InvoicePersonIDAndInvoiceMasterRefersTo", myID, Trim(Str(Val(txtInvoiceMasterRefersTo.text + 2))))
    
    If MainSeekRecord("CommonDB", txtCustomersOrSuppliers.text, "ID", myID, True, txtPersonID, txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo.text, txtPersonTaxOfficeID, txtPersonVATStateID, txtAccountCode) Then
        'œÈÍÔÌÔÏÈÍﬁ ıÁÒÂÛﬂ·
        Set tmpRecordset = CheckForMatch("CommonDB", "TaxOffices", "TaxOfficeID", "Numeric", txtPersonTaxOfficeID.text)
        txtPersonTaxOfficeID.text = tmpRecordset.Fields(0)
        txtTaxOfficeDescription.text = tmpRecordset.Fields(1)
        ' ·ËÂÛÙ˛Ú ÷.–.¡.
        Set tmpRecordset = CheckForMatch("CommonDB", "VATStates", "VATStateID", "Numeric", txtPersonVATStateID.text)
        txtPersonVATStateID.text = tmpRecordset.Fields(0)
        txtVATStateDescription.text = tmpRecordset.Fields(1)
        '
        EnableFields txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode, cmdIndex(0), cmdIndex(1), cmdIndex(2)
        UpdateButtons Me, 5, 0, 1, IIf(blnEnableDelete, 1, 0), 0, 1, 0
        blnStatus = False
        SeekRecord = True
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", txtCustomersOrSuppliers.text, strApplicationName, "ID", txtPersonID.text, "True") Then
        ClearFields txtPersonID, txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode
        DisableFields txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode, cmdIndex(0), cmdIndex(1), cmdIndex(2)
        UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
    End If

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtPersonID, txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode
    EnableFields txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode, cmdIndex(0), cmdIndex(1), cmdIndex(2)
    UpdateButtons Me, 5, 0, 1, 0, 0, 1, 0
    txtDescription.SetFocus

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", txtCustomersOrSuppliers.text, blnStatus, strApplicationName, "ID", txtPersonID.text, txtDescription.text, txtProfession.text, txtAddress.text, txtPhones.text, txtPersonInCharge.text, txtEmail.text, txtTaxNo.text, txtPersonTaxOfficeID.text, txtPersonVATStateID.text, txtAccountCode.text, 1, strCurrentUser) <> 0 Then
        ClearFields txtPersonID, txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode
        DisableFields txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode, cmdIndex(0), cmdIndex(1), cmdIndex(2)
        UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    '≈˘ÌıÏﬂ·
    If Len(Trim(txtDescription.text)) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtDescription.SetFocus
        Exit Function
    End If
    
    '≈ÎÂ„˜ÔÚ ¡.÷.Ã.
    If blnCustomerCheckTaxNo Then
        If Len(txtTaxNo.text) = 0 Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
            End If
            txtTaxNo.SetFocus
            Exit Function
        End If
    End If
    
    'ƒ.œ.’.
    If Len(txtPersonTaxOfficeID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtTaxOfficeDescription.SetFocus
        Exit Function
    End If
    
    ' ˘‰. √ÂÌ. ÀÔ„ÈÛÙÈÍﬁÚ
    If Len(txtAccountCode.text) <> 0 Then
        If Len(txtAccountCode.text) <> intAccountsCodeLength Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
            End If
            txtAccountCode.SetFocus
            Exit Function
        End If
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
            ShowIndex
        Case 4
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select

End Sub

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'œÈÍÔÌÔÏÈÍﬁ ’ÁÒÂÛﬂ·
            Set tmpRecordset = CheckForMatch("CommonDB", "TaxOffices", "TaxOfficeDescription", "String", txtTaxOfficeDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "œÌÔÏ·Ûﬂ·", 0, 40, 1, 0)
                txtPersonTaxOfficeID.text = tmpTableData.strCode
                txtTaxOfficeDescription.text = tmpTableData.strFirstField
            End If
        Case 1
            'œÈÍÔÌÔÏÈÍﬁ ’ÁÒÂÛﬂ·
            With TablesTaxOffices
                .Tag = "True"
                .Show 1, Me
            End With
        Case 2
            ' ·ËÂÛÙ˛Ú ÷.–.¡.
            Set tmpRecordset = CheckForMatch("CommonDB", "VATStates", "VATStateDescription", "String", txtVATStateDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "œÌÔÏ·Ûﬂ·", 0, 40, 1, 0)
                txtPersonVATStateID.text = tmpTableData.strCode
                txtVATStateDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)
    
End Sub

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
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()
    
    PositionControls Me, False
    ColorizeControls Me, False, False
    ClearFields txtPersonID, txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode
    DisableFields txtDescription, txtProfession, txtAddress, txtPhones, txtPersonInCharge, txtEmail, txtTaxNo, txtPersonTaxOfficeID, txtTaxOfficeDescription, txtPersonVATStateID, txtVATStateDescription, txtAccountCode, cmdIndex(0), cmdIndex(1), cmdIndex(2)
    UpdateButtons Me, 5, 1, 0, 0, 1, 0, 1

End Sub

Private Sub txtTaxOfficeDescription_Change()

    If txtTaxOfficeDescription.text = "" Then
        ClearFields txtPersonTaxOfficeID
    End If
    
End Sub

Private Sub txtTaxOfficeDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub

Private Sub txtTaxOfficeDescription_Validate(Cancel As Boolean)

    If txtPersonTaxOfficeID.text = "" And txtTaxOfficeDescription.text <> "" Then cmdIndex_Click 0

End Sub

Private Sub txtVATStateDescription_Change()

    If txtVATStateDescription.text = "" Then
        ClearFields txtPersonVATStateID
    End If
    
End Sub

Private Sub txtVATStateDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtVATStateDescription_Validate(Cancel As Boolean)

    If txtPersonVATStateID.text = "" And txtVATStateDescription.text <> "" Then cmdIndex_Click 2

End Sub

Private Function ShowIndex()

    With PersonsIndex
        .Tag = "True"
        .txtCustomersOrSuppliers.text = txtCustomersOrSuppliers.text
        .lblTitle.Caption = "≈ıÒÂÙﬁÒÈÔ " & IIf(txtCustomersOrSuppliers.text = "Customers", "ÂÎ·Ù˛Ì", "ÒÔÏÁËÂıÙ˛Ì")
        .txtInvoiceMasterRefersTo.text = txtInvoiceMasterRefersTo.text
        .Show 1, Me
    End With

End Function
