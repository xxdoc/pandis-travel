VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form InvoicesIn 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14430
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   14430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   61
      Top             =   8100
      Width           =   8940
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         TabIndex        =   67
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
   Begin VB.Frame frmProgress 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   990
      Left            =   2250
      TabIndex        =   23
      Top             =   5700
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   465
         Left            =   150
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   820
         Picture         =   "InvoicesIn.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "InvoicesIn.frx":001C
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
         TabIndex        =   25
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   9600
      TabIndex        =   14
      Top             =   3000
      Width           =   4515
      Begin VB.TextBox Text14 
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
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "Codes.CodeHandID"
         Top             =   4575
         Width           =   3540
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
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "Codes.CodeLastDate"
         Top             =   3825
         Width           =   3540
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
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "Codes.CodeLastNo"
         Top             =   3450
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
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "Codes.CodeSuppliers"
         Top             =   4200
         Width           =   3540
      End
      Begin VB.TextBox txtCodePersonsPlusOrMinus 
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
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   4200
         Width           =   780
      End
      Begin VB.TextBox txtCodeLastNo 
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
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3450
         Width           =   780
      End
      Begin VB.TextBox txtCodeLastDate 
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
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3825
         Width           =   780
      End
      Begin VB.CheckBox chkCodeHandID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   4575
         Width           =   780
      End
      Begin VB.TextBox txtInvoicePersonID 
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1950
         Width           =   780
      End
      Begin VB.TextBox Text16 
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
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Invoices.InvoicePersonID"
         Top             =   1950
         Width           =   3540
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
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceCodeID"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceID"
         Top             =   825
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceID 
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
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin VB.TextBox txtInvoiceCodeID 
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
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
      Begin VB.TextBox txtInvoiceDateIn 
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2325
         Width           =   780
      End
      Begin VB.TextBox Text9 
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
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceDateIn"
         Top             =   2325
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceTrnID 
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
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox Text11 
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
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceTrnID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Text            =   "InvoicesIn.InvoiceInPaymentTermID"
         Top             =   3075
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceInPaymentTermID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Text            =   "7"
         Top             =   3075
         Width           =   780
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Text            =   "InvoicesIn.InvoiceInExpenseCategoryID"
         Top             =   2700
         Width           =   3540
      End
      Begin VB.TextBox txtInvoiceInExpenseCategoryID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "7"
         Top             =   2700
         Width           =   780
      End
      Begin VB.TextBox txtInvoiceSecondaryRefersTo 
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
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox Text8 
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
         Text            =   "Invoices.InvoiceSecondaryRefersTo"
         Top             =   450
         Width           =   3540
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
         TabIndex        =   33
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
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   75
         Width           =   780
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   4950
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "InvoicesIn.frx":0038
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin UserControls.newDate mskDateIssue 
      Height          =   465
      Left            =   2175
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
   Begin UserControls.newText txtCodeShortDescriptionA 
      Height          =   465
      Left            =   2175
      TabIndex        =   1
      Top             =   1650
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
   Begin UserControls.newText txtSupplierDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   3
      Top             =   2700
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
   Begin UserControls.newText txtExpenseDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   4
      Top             =   3225
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
   Begin UserControls.newText txtPaymentTermDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   5
      Top             =   3750
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   2925
      TabIndex        =   15
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
      PicNormal       =   "InvoicesIn.frx":1248
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   5
      Left            =   3375
      TabIndex        =   16
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
      PicNormal       =   "InvoicesIn.frx":17E2
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7200
      TabIndex        =   17
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
      PicNormal       =   "InvoicesIn.frx":1D7C
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   7200
      TabIndex        =   18
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
      PicNormal       =   "InvoicesIn.frx":2316
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   6
      Left            =   7200
      TabIndex        =   19
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
      PicNormal       =   "InvoicesIn.frx":28B0
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   7650
      TabIndex        =   20
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
      PicNormal       =   "InvoicesIn.frx":2E4A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   7650
      TabIndex        =   21
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
      PicNormal       =   "InvoicesIn.frx":33E4
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   7
      Left            =   7650
      TabIndex        =   22
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
      PicNormal       =   "InvoicesIn.frx":397E
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtInvoiceNo 
      Height          =   465
      Left            =   2175
      TabIndex        =   2
      Top             =   2175
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   10
      Text            =   "9999999999"
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
   Begin iGrid300_10Tec.iGrid grdInvoicesIn 
      Height          =   2490
      Left            =   2175
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4275
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   4392
      Appearance      =   0
      BackColor       =   14737632
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
   Begin UserControls.newText mskTotalNet 
      Height          =   465
      Left            =   3825
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      Enabled         =   0   'False
      Alignment       =   1
      ForeColor       =   0
      Text            =   "9.999.999,99"
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
   Begin UserControls.newText mskTotalGross 
      Height          =   465
      Left            =   6075
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      Enabled         =   0   'False
      Alignment       =   1
      ForeColor       =   0
      Text            =   "9.999.999,99"
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
   Begin UserControls.newText mskTotalVAT 
      Height          =   465
      Left            =   4950
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      Enabled         =   0   'False
      Alignment       =   1
      ForeColor       =   0
      Text            =   "9.999.999,99"
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
   Begin VB.Label lblCodeDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   3825
      TabIndex        =   38
      Top             =   1725
      Width           =   4215
   End
   Begin VB.Label lblCodeBatch 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ΣΕΙΡΑ Ω"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   8100
      TabIndex        =   37
      Top             =   1725
      Width           =   615
   End
   Begin VB.Label lblCodeHand 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ΧΕΙΡΟΓΡΑΦΟ"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   8775
      TabIndex        =   36
      Top             =   1725
      Width           =   990
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Συνολική αξία"
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
      Left            =   6075
      TabIndex        =   31
      Top             =   6825
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Αξία Φ.Π.Α."
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
      Left            =   4950
      TabIndex        =   30
      Top             =   6825
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Καθαρή αξία"
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
      Left            =   3825
      TabIndex        =   27
      Top             =   6825
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   4200
      Top             =   7575
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
      Index           =   13
      Left            =   2850
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   4950
      Top             =   8775
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   9825
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
      Index           =   0
      Left            =   1725
      Top             =   1575
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
      Top             =   8100
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Αγορές"
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
      Width           =   1635
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Παραστατικό"
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
      TabIndex        =   12
      Top             =   1725
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Νο παραστατικού"
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
      TabIndex        =   11
      Top             =   2250
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Κατηγορία εξόδου"
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
      TabIndex        =   10
      Top             =   3300
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Όρος πληρωμής"
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
      TabIndex        =   9
      Top             =   3825
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Προμηθευτής"
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
      TabIndex        =   8
      Top             =   2775
      Width           =   1290
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
      Index           =   2
      Left            =   450
      TabIndex        =   7
      Top             =   1200
      Width           =   1290
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
Attribute VB_Name = "InvoicesIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim blnCancel As Boolean
Dim lngTrnID As Long
Dim IsError As Boolean

Dim IsGridEditInProgress As Boolean
Dim strGridFocus As String


Private Function DeleteExpensesPerVAT()

    Dim lngRow As Long
    
    If IsError Then Exit Function
    
    With grdInvoicesIn
        For lngRow = 1 To .RowCount
             If Not MainDeleteRecord("CommonDB", "ExpensesPerVAT", strAppTitle, "ExpensePerVATID", .CellValue(lngRow, "ID"), False) Then
                IsError = True
                Exit For
             End If
        Next lngRow
    End With

End Function

Private Function DisplayTotals()

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    Dim curTotalNet As Currency
    Dim curTotalVAT As Currency
    Dim curTotalGross As Currency
    
    For lngRow = 1 To grdInvoicesIn.RowCount
        curTotalNet = curTotalNet + grdInvoicesIn.CellValue(lngRow, 4)
        curTotalVAT = curTotalVAT + grdInvoicesIn.CellValue(lngRow, 5)
        curTotalGross = curTotalGross + grdInvoicesIn.CellValue(lngRow, 6)
    Next lngRow
    
    mskTotalNet.text = Format(curTotalNet, "#,##0.00")
    mskTotalVAT.text = Format(curTotalVAT, "#,##0.00")
    mskTotalGross.text = Format(curTotalGross, "#,##0.00")
    
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function

Private Function DoCalculations(lngRow As Long, lngCol As Long, blnCalculateVATAmount)
    
        On Error GoTo ErrTrap
        
        'Local μεταβλητές
        Dim curNetAmount As Currency
        Dim curVATPercent As Currency
        Dim curVATAmount As Currency
        Dim curGrossAmount As Currency
        
        'Υπολογίζω την αξία του ΦΠΑ και το σύνολο της γραμμής
        curNetAmount = grdInvoicesIn.CellValue(lngRow, 4)
        curVATPercent = IIf(blnCalculateVATAmount, curNetAmount * (grdInvoicesIn.CellValue(lngRow, 3) / 100), grdInvoicesIn.CellValue(lngRow, 5))
        curGrossAmount = curNetAmount + curVATPercent
        
        'Εμφανίζω την αξία του ΦΠΑ και το σύνολο της γραμμής
        grdInvoicesIn.CellValue(lngRow, 5) = curVATPercent
        grdInvoicesIn.CellValue(lngRow, 6) = curGrossAmount
        
        'Βγαίνω
        Exit Function
        
ErrTrap:
        If Err.Number = 13 Or Err.Number = 6 Then Resume Next
        
End Function
   
Private Function AbortProcedure(blnStatus)
    
    If IsGridEditInProgress Then
        IsGridEditInProgress = False
        grdInvoicesIn.CancelEdit
        Exit Function
    End If
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
            ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
            ClearFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
            ClearFields grdInvoicesIn
            ClearFields mskTotalNet, mskTotalVAT, mskTotalGross
            DisableFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
            UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("BuysIndex"), 0, 1), 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()

    IsError = False
    
    BeginTrans
    
    DeleteInvoice
    DeleteInvoiceTrn
    DeleteExpensesPerVAT
    
    If Not IsError Then
        CommitTrans
        ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
        ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
        ClearFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
        ClearFields mskTotalNet, mskTotalVAT, mskTotalGross
        ClearFields grdInvoicesIn
        DisableFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
        UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("InvoicesInIndex"), 0, 1), 0, 1
    Else
        Rollback
    End If
    
End Function

Public Function FindExpensesPerVAT(lngTrnID)

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

    'Πλέγμα
    With grdInvoicesIn
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT ExpensePerVATID, VATPercentDescription,ExpensePerVATPercentID, ExpensePerVATNetAmount, ExpensePerVATVATAmount, ExpensePerVATGrossAmount " _
        & "FROM VATPercents " _
        & "INNER JOIN ExpensesPerVAT ON VATPercents.VATPercentID = ExpensesPerVAT.ExpensePerVATPercentID " _
        
    'InvoiceTrnID
    strThisParameter = "lngInvoiceTrnID long"
    strThisQuery = "ExpensesPerVAT.ExpensePerVATTrnID = lngInvoiceTrnID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = lngTrnID
        
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    TempQuery.SQL = strSQL
    
    strOrder = " ORDER BY VATPercentDescription "
    
    strParameters = "PARAMETERS " & strParameters & "; "
    strParFields = "WHERE " & strParFields
    strSQL = strParameters & strSQL & strParFields
    TempQuery.SQL = strSQL & strOrder
    
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
    Set FindExpensesPerVAT = rstRecordset
    
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
    DisplayErrorMessage True, Err.Description

End Function

Private Function InitializeGridWithZeroValues()

    Dim intLoop As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    Dim curTotalNet As Currency
    Dim curTotalVAT As Currency
    Dim curTotalGross As Currency
    
    For lngRow = 1 To grdInvoicesIn.RowCount
        grdInvoicesIn.CellValue(lngRow, 4) = 0
        grdInvoicesIn.CellValue(lngRow, 5) = 0
        grdInvoicesIn.CellValue(lngRow, 6) = 0
    Next lngRow
    
    mskTotalNet.text = Format(curTotalNet, "#,##0.00")
    mskTotalVAT.text = Format(curTotalVAT, "#,##0.00")
    mskTotalGross.text = Format(curTotalGross, "#,##0.00")

End Function

Private Function NewRecord()
    
    Dim tmpRecordset As Recordset
    
    blnStatus = True
    blnCancel = False
    
    ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
    ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
    ClearFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
    ClearFields grdInvoicesIn
    ClearFields mskTotalNet, mskTotalVAT, mskTotalGross
    EnableFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
    
    If RefreshListWithVATPercents Then
        InitializeGridWithZeroValues
        EditableFields grdInvoicesIn
        EnableTabStop grdInvoicesIn
        UpdateButtons Me, 5, 0, 1, 0, 0, 1, 0
    End If
    
    mskDateIssue.text = Date
    txtInvoiceDateIn.text = Date
    
    mskDateIssue.SetFocus
    
End Function

Private Function FindChildTransactions(lngTrnID)

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
    
    'Local μεταβλητές
    Dim lngIndex As Long
    Dim lngRow As Long
    
    'Recordsets
    Dim rstRecordset As Recordset
    Dim tmpRecordset As Recordset
    
    Set TempQuery = CommonDB.CreateQueryDef("")
    
    'Κύριο SQL
    strSQL = "SELECT ID, TrnID, Transactions.VATPercentID, NetAmount, VATAmount, GrossAmount, VATPercentDescription " _
        & "FROM Transactions " _
        & "INNER JOIN VATPercents ON " _
        & "Transactions.VATPercentID = VATPercents.VATPercentID "

    'TrnID κινήσεων
    strThisParameter = "lngTrnID Long"
    strThisQuery = "TrnID = lngTrnID "
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = lngTrnID
        
    'Ταξινόμηση
    strOrder = " ORDER BY VATPercentDescription "
        
    'Προσθέτω τα κριτήρια
    If strThisParameter <> "" Then
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = "WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
    End If
    
    'SQL
    TempQuery.SQL = strSQL & strOrder
    
    'Κριτήρια
    If strThisParameter <> "" Then
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
    End If
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then IsError = False: FindChildTransactions = True: Exit Function
    
    'Βάζω γραμμές στο πλέγμα
    grdInvoicesIn.AddRow , , , , , , , rstRecordset.RecordCount
    
    'Πλέγμα
    With grdInvoicesIn
        .Editable = True
        .Redraw = True
        .RowMode = False
        .TabStop = True
    End With
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        While Not .EOF
            With grdInvoicesIn
                lngRow = lngRow + 1
                .CellValue(lngRow, "ID") = rstRecordset!ID
                .CellValue(lngRow, "VATPercentID") = rstRecordset!VATPercentID
                .CellValue(lngRow, "Description") = rstRecordset!VATPercentDescription
                .CellValue(lngRow, 4) = rstRecordset!NetAmount
                .CellValue(lngRow, 5) = rstRecordset!VATAmount
                .CellValue(lngRow, 6) = rstRecordset!GrossAmount
            End With
            .MoveNext
        Wend
    End With
    
    'Τελικές ενέργειες
    FindChildTransactions = True
    
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
    FindChildTransactions = False
    DisplayErrorMessage True, Err.Description

End Function

Private Function PopulateGrid(rstExpensesPerVAT As Recordset)

    Dim lngRow As Long
    
    grdInvoicesIn.AddRow , , , , , , , rstExpensesPerVAT.RecordCount
    
    With grdInvoicesIn
        .Editable = True
        .Redraw = True
        .RowMode = False
        .TabStop = True
    End With
    
    With rstExpensesPerVAT
        While Not .EOF
            With grdInvoicesIn
                lngRow = lngRow + 1
                .CellValue(lngRow, "ID") = rstExpensesPerVAT!ExpensePerVATID
                .CellValue(lngRow, "VATPercentID") = rstExpensesPerVAT!ExpensePerVATPercentID
                .CellValue(lngRow, "Description") = rstExpensesPerVAT!VATPercentDescription
                .CellValue(lngRow, 4) = rstExpensesPerVAT!ExpensePerVATNetAmount
                .CellValue(lngRow, 5) = rstExpensesPerVAT!ExpensePerVATVATAmount
                .CellValue(lngRow, 6) = rstExpensesPerVAT!ExpensePerVATGrossAmount
            End With
            .MoveNext
        Wend
    End With

End Function

Private Function RefreshListWithVATPercents()

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

    'Πλέγμα
    With grdInvoicesIn
        .Clear
        .Editable = False
        .Redraw = False
        .RowMode = False
    End With
    
    'Κυρίως διαδικασία
    strSQL = "SELECT * FROM VATPercents ORDER BY VATPercentDescription"

    Set TempQuery = CommonDB.CreateQueryDef("")
    
    TempQuery.SQL = strSQL
    
    'Ανοίγω το recordset
    Set rstRecordset = TempQuery.OpenRecordset()
    
    'Αν δεν έχω εγγραφές, βγαίνω
    If rstRecordset.RecordCount = 0 Then Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strAppTitle, rstRecordset
    
    'Γεμίζω το πλέγμα
    With rstRecordset
        grdInvoicesIn.AddRow , , , , , , , rstRecordset.RecordCount
        Do While Not .EOF
            lngRow = lngRow + 1
            UpdateProgressBar Me
            grdInvoicesIn.CellValue(lngRow, "VATPercentID") = !VATPercentID
            grdInvoicesIn.CellValue(lngRow, "Description") = !VATPercentDescription
            .MoveNext
        Loop
        grdInvoicesIn.Redraw = True
    End With
    
    'Τελικές ενέργειες
    frmProgress.Visible = False
    RefreshListWithVATPercents = True
    
    Exit Function
    
ErrTrap:
    RefreshListWithVATPercents = False
    ClearFields grdInvoicesIn, frmProgress
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function SaveExpensesPerVAT()

    Dim lngRow As Long
    
    If IsError Then Exit Function
    
    With grdInvoicesIn
        For lngRow = 1 To .RowCount
            If MainSaveRecord("CommonDB", "ExpensesPerVAT", blnStatus, strAppTitle, "ExpensePerVATID", .CellValue(lngRow, "ID"), txtInvoiceTrnID.text, grdInvoicesIn.CellValue(lngRow, "VATPercentID"), grdInvoicesIn.CellValue(lngRow, "Net"), grdInvoicesIn.CellValue(lngRow, "Tax"), grdInvoicesIn.CellValue(lngRow, "Gross")) <> 0 Then
                IsError = False
            Else
                IsError = True
                Exit Function
            End If
        Next lngRow
    End With

    IsError = False

End Function

Private Function SaveInvoice()

    If blnStatus Then txtInvoiceTrnID.text = AddOneToTheLastRecord("Invoices")

    If MainSaveRecord("CommonDB", "Invoices", blnStatus, strAppTitle, "InvoiceID", txtInvoiceID.text, txtInvoiceTrnID.text, txtInvoiceMasterRefersTo.text, txtInvoiceSecondaryRefersTo.text, mskDateIssue.text, txtInvoiceDateIn.text, txtInvoiceCodeID.text, txtInvoiceNo.text, txtInvoicePersonID.text, strCurrentUser) <> 0 Then
        IsError = False
    Else
        IsError = True
    End If
    
End Function

Private Function SaveInvoiceIn()

    If IsError Then Exit Function
    
    If MainSaveRecord("CommonDB", "InvoicesIn", blnStatus, strAppTitle, "InvoiceInTrnID", txtInvoiceTrnID.text, txtInvoiceTrnID.text, _
        txtInvoiceInExpenseCategoryID.text, _
        txtInvoiceInPaymentTermID.text, _
        mskTotalGross.text) <> 0 Then
        IsError = False
    Else
        IsError = True
    End If
    
End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    IsError = False
    
    BeginTrans
    
    SaveInvoice
    SaveInvoiceIn
    SaveExpensesPerVAT
    
    If Not IsError Then
        CommitTrans
        ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
        ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
        ClearFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
        ClearFields mskTotalNet, mskTotalVAT, mskTotalGross
        ClearFields grdInvoicesIn
        DisableFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
        UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("InvoicesInIndex"), 0, 1), 0, 1
    Else
        Rollback
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Ημερομηνία
    If Not CheckDate(mskDateIssue.text, strAppTitle) Then
        mskDateIssue.SetFocus
        Exit Function
    End If
    
    'Καταχώρηση σε ημερομηνία μεγαλύτερη από σήμερα
    If CDate(mskDateIssue.text) > Date Then
        If MyMsgBox(4, strAppTitle, strAppMessages(5), 1) Then
        End If
        mskDateIssue.SetFocus
        Exit Function
    End If
    
    'Μήκος ημερομηνίας
    If Len(mskDateIssue.text) <> 10 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskDateIssue.SetFocus
        Exit Function
    End If
    
    'Στοιχείο
    If Len(txtInvoiceCodeID.text) = 0 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtCodeShortDescriptionA.SetFocus
        Exit Function
    End If
    
    'Νο παραστατικού
    If txtInvoiceNo.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtInvoiceNo.SetFocus
        Exit Function
    End If
    
    'Νο παραστατικού = αριθμός
    If Not IsNumeric(txtInvoiceNo.text) Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        txtInvoiceNo.SetFocus
        Exit Function
    End If
    
    'Στοιχείο ήδη καταχωρημένο: Ελέγχω αν το νούμερο του στοιχείου υπάρχει ήδη στην χρήση
    'If CheckForDuplicateInvoice(mskDateIssue.text, txtSupplierID.text, txtCodeID.text, txtInvoiceNo.text) Then
    '    If MyMsgBox(4, strAppTitle, strStandardMessages(22), 1) Then
    '    End If
    '    txtSupplierDescription.SetFocus
    '    Exit Function
    'End If
    
    'Προμηθευτής
    If Len(txtInvoicePersonID.text) = 0 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtSupplierDescription.SetFocus
        Exit Function
    End If
    
    'Όρος πληρωμής
    If Len(txtInvoiceInPaymentTermID.text) = 0 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtPaymentTermDescription.SetFocus
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
            FindRecords
        Case 4
            AbortProcedure False
        Case 5
            AbortProcedure True
    End Select

End Sub

Private Function FindRecords()

    With InvoicesInIndex
        .Tag = "True"
        .txtInvoiceMasterRefersTo.text = txtInvoiceMasterRefersTo.text
        .txtInvoiceSecondaryRefersTo.text = txtInvoiceSecondaryRefersTo.text
        .Show 1, Me
    End With

End Function

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Κατηγορία εξόδου - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "ExpensesCategories", "ExpenseCategoryDescription", "String", txtExpenseDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtInvoiceInExpenseCategoryID.text = tmpTableData.strCode
            txtExpenseDescription.text = tmpTableData.strFirstField
        Case 1
            'Κατηγορία εξόδου - F5
            With TablesExpenseCategories
                .Tag = "True"
                .Show 1, Me
            End With
        Case 2
            'Προμηθευτής - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Suppliers", "Description", "String", txtSupplierDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
            txtInvoicePersonID.text = tmpTableData.strCode
            txtSupplierDescription.text = tmpTableData.strFirstField
        Case 3
            'Προμηθευτής - F5
            With persons
                .Tag = "True"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .lblTitle.Caption = "Προμηθευτές"
                .Show 1, Me
            End With
        Case 4
            'Παραστατικό - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Codes", "CodeShortDescriptionA, CodeMasterRefersTo", "String, String", txtCodeShortDescriptionA.text, txtInvoiceMasterRefersTo.text)
            tmpTableData = DisplayIndex(tmpRecordset, 3, True, 8, 0, 3, 5, 6, 7, 9, 10, 11, "ID", "Συντ. Α'", "Περιγραφή", "Σειρά", "Χειρόγραφο", "Προμηθευτές", "Τελευταίο Νο", "Ημερομηνία", 0, 6, 40, 6, 0, 0, 0, 0, 1, 1, 0, 1, 1, 1, 1, 1)
            txtInvoiceCodeID.text = tmpTableData.strCode
            txtCodeShortDescriptionA.text = tmpTableData.strFirstField
            lblCodeDescription.Caption = tmpTableData.strSecondField
            'txtCodePersonsPlusOrMinus.text = tmpTableData.strFifthField
        Case 5
            'Παραστατικό - F5
            With TablesCodes
                .Tag = "True"
                .txtCodeMasterRefersTo.text = txtInvoiceMasterRefersTo.text
                .txtCodeSecondaryRefersTo.text = "0"
                .Show 1, Me
            End With
        Case 6
            'Όρος πληρωμής - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "PaymentTerms", "PaymentTermDescription", "String", txtPaymentTermDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 2, "ID", "Περιγραφή", "Πίστωση", 0, 40, 0, 1, 0, 0)
            txtInvoiceInPaymentTermID.text = tmpTableData.strCode
            txtPaymentTermDescription.text = tmpTableData.strFirstField
        Case 7
            'Όρος πληρωμής - F5
            With TablesPaymentTerms
                .Tag = "True"
                .Show 1, Me
            End With
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
    End If
    
    'AddDummyLines grdInvoicesIn, "99999", "99999", "-9.999.999,99", "-9.999.999,99", "-9.999.999,99"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Public Function SeekRecord(lngTrnID)

    Dim intIndex As Byte
    Dim strThisQuery As String
    Dim strParameters As String
    Dim strParFields As String
    Dim strThisParameter As String
    Dim strOrder As String
    Dim strLogic As String
    Dim arrQuery() As Variant
    Dim strSQL As String
    
    Dim rstRecordset As Recordset
    
    strSQL = "SELECT " _
        & "Invoices.InvoiceID, Invoices.InvoiceTrnID, Invoices.InvoiceMasterRefersTo, Invoices.InvoiceSecondaryRefersTo, Invoices.InvoiceDateIssue, Invoices.InvoiceDateIn, Invoices.InvoiceCodeID, Invoices.InvoiceNo, Invoices.InvoicePersonID, Invoices.InvoiceDateIssue, Invoices.InvoiceNo, " _
        & "InvoicesIn.InvoiceInAmount, " _
        & "Codes.CodeShortDescriptionA, Codes.CodeDescription, Codes.CodeBatch, Codes.CodeHandID, Codes.CodeSuppliers, Codes.CodeLastNo, Codes.CodeLastDate, " _
        & "Suppliers.Description, " _
        & "ExpensesCategories.ExpenseCategoryID, ExpensesCategories.ExpenseCategoryDescription, " _
        & "PaymentTerms.PaymentTermID, PaymentTerms.PaymentTermDescription " _
        & "FROM ((((Invoices " _
        & "INNER JOIN InvoicesIn ON Invoices.InvoiceTrnID = InvoicesIn.InvoiceInTrnID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN Suppliers ON Invoices.InvoicePersonID = Suppliers.ID) " _
        & "INNER JOIN ExpensesCategories ON InvoicesIn.InvoiceInExpenseCategoryID = ExpensesCategories.ExpenseCategoryID) " _
        & "INNER JOIN PaymentTerms ON InvoicesIn.InvoiceInPaymentTermID = PaymentTerms.PaymentTermID "
        
    'InvoiceTrnID
    strThisParameter = "lngInvoiceTrnID long"
    strThisQuery = "Invoices.InvoiceTrnID = lngInvoiceTrnID"
    strLogic = " AND "
    GoSub UpdateSQLString
    arrQuery(intIndex) = lngTrnID

    Set TempQuery = CommonDB.CreateQueryDef("")
    
    strParameters = "PARAMETERS " & strParameters & "; "
    strParFields = "WHERE " & strParFields
    strSQL = strParameters & strSQL & strParFields
    TempQuery.SQL = strSQL & strOrder
    
    For intIndex = 1 To UBound(arrQuery)
        TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
    Next intIndex
    
    Set rstRecordset = TempQuery.OpenRecordset()
    
    Set SeekRecord = rstRecordset
    
    Exit Function

UpdateSQLString:
    intIndex = intIndex + 1
    strParameters = IIf(intIndex > 1, strParameters & ", ", strParameters)
    strParFields = IIf(intIndex > 1, strParFields & strLogic, strParFields)
    strParameters = strParameters & strThisParameter
    strParFields = strParFields & strThisQuery
    ReDim Preserve arrQuery(intIndex)
    
    Return
    
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
    AddColumnsToGrid grdInvoicesIn, 44, GetSetting(strAppTitle, "Layout Strings", "grdInvoicesIn"), "04NCNID,04NCNVATPercentID,04NRFDescription,10YRFNet,04NRFTax,04NRFGross", "ID,VATPercentID,Ποσοστό,Καθαρή αξία,Αξία Φ.Π.Α.,Συνολική αξία"
    SetUpGrid lstIconList, grdInvoicesIn
    ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
    ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
    ClearFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
    ClearFields mskTotalNet, mskTotalVAT, mskTotalGross
    ClearFields grdInvoicesIn
    DisableFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
    UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("BuysIndex"), 0, 1), 0, 1

End Sub

Private Sub grdInvoicesIn_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)

    Dim blnCalculateVATAmount As Boolean
    
    Select Case lCol
        Case 4
            'Καθαρή αξία
            If grdInvoicesIn.CellValue(lRow, lCol) <> "" Then MoveToNextColumn grdInvoicesIn, lRow, lCol: blnCalculateVATAmount = True
        Case 5
            'Αξία ΦΠΑ
            If grdInvoicesIn.CellValue(lRow, lCol) <> "" Then MoveToNextColumn grdInvoicesIn, lRow, lCol: blnCalculateVATAmount = False
    End Select
    
    DoCalculations lRow, lCol, blnCalculateVATAmount
    
    DisplayTotals
    
    IsGridEditInProgress = False

End Sub

Private Sub grdInvoicesIn_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid300_10Tec.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)

    If lCol = 4 Or lCol = 5 Then
        vNewValue = Replace(sNewText, ".", ",")
        If vNewValue = "," Then
            vNewValue = "0,00"
        End If
    End If

End Sub

Private Sub grdInvoicesIn_GotFocus()

    If grdInvoicesIn.RowCount = 0 Or Not grdInvoicesIn.Enabled Then Exit Sub
    
    Select Case strGridFocus
        Case Is = "FromTop"
            grdInvoicesIn.SetCurCell 1, 4
            strGridFocus = ""
    End Select

End Sub

Private Sub grdInvoicesIn_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdInvoicesIn_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    'Πάνω βελάκι
    If KeyCode = 38 Then
        If grdInvoicesIn.CurRow = 1 Then
            grdInvoicesIn.CurCol = 0
            txtPaymentTermDescription.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub grdInvoicesIn_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)
    
    If lCol = 1 Or lCol = 2 Or lCol = 3 Or lCol = 6 Then bCancel = True
    
    IsGridEditInProgress = True
    
    sText = ClearNumberFormat(sText)
    
    If lCol = 4 Or lCol = 5 Then
        If CheckForAcceptableKey(iKeyAscii) Then
            CaptureNumbers grdInvoicesIn.TextEditText, lRow, lCol, iKeyAscii, True
        Else
            bCancel = True
        End If
    End If

End Sub

Private Sub grdInvoicesIn_TextEditKeyPress(ByVal lRow As Long, ByVal lCol As Long, KeyAscii As Integer)

    If lCol = 4 Or lCol = 5 Then
        If CheckForAcceptableKey(KeyAscii) Then
            CaptureNumbers grdInvoicesIn.TextEditText, lRow, lCol, KeyAscii, True
        Else
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdInvoicesIn", grdInvoicesIn.LayoutCol
    
End Sub

Private Sub txtCodeShortDescriptionA_Change()

    If txtCodeShortDescriptionA.text = "" Then
        ClearFields txtInvoiceCodeID, txtCodeShortDescriptionA, lblCodeDescription, txtCodePersonsPlusOrMinus
    End If
    
End Sub

Private Sub txtCodeShortDescriptionA_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    If KeyCode = vbKeyF5 Then cmdIndex_Click 5

End Sub

Private Sub txtCodeShortDescriptionA_Validate(Cancel As Boolean)

    If txtInvoiceCodeID.text = "" And txtCodeShortDescriptionA.text <> "" Then cmdIndex_Click 4

End Sub

Private Sub txtExpenseDescription_Change()

    If txtExpenseDescription.text = "" Then
        ClearFields txtInvoiceInExpenseCategoryID
    End If

End Sub

Private Sub txtExpenseDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub

Private Sub txtExpenseDescription_Validate(Cancel As Boolean)

    If txtInvoiceInExpenseCategoryID.text = "" And txtExpenseDescription.text <> "" Then cmdIndex_Click 0

End Sub

Private Sub txtPaymentTermDescription_Change()

    If txtPaymentTermDescription.text = "" Then
        ClearFields txtInvoiceInPaymentTermID
    End If

End Sub

Private Sub txtPaymentTermDescription_GotFocus()

    strGridFocus = "FromTop"

End Sub

Private Sub txtPaymentTermDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 6
    If KeyCode = vbKeyF5 Then cmdIndex_Click 7

End Sub

Private Sub txtPaymentTermDescription_Validate(Cancel As Boolean)

    If txtInvoiceInPaymentTermID.text = "" And txtPaymentTermDescription.text <> "" Then cmdIndex_Click 6

End Sub

Private Sub txtSupplierDescription_Change()

    If txtSupplierDescription.text = "" Then
        ClearFields txtInvoicePersonID
    End If

End Sub

Private Sub txtSupplierDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    If KeyCode = vbKeyF5 Then cmdIndex_Click 3

End Sub

Private Sub txtSupplierDescription_Validate(Cancel As Boolean)

    If txtInvoicePersonID.text = "" And txtSupplierDescription.text <> "" Then cmdIndex_Click 2

End Sub

Private Function DeleteInvoice()

    If Not MainDeleteRecord("CommonDB", "Invoices", strAppTitle, "InvoiceID", txtInvoiceID.text, True) Then
        IsError = True
    End If

End Function

Private Function DeleteInvoiceTrn()

    Dim lngRow As Long
    
    If IsError Then Exit Function
    
    If Not MainDeleteRecord("CommonDB", "InvoicesIn", strAppTitle, "InvoiceInTrnID", txtInvoiceTrnID.text, False) Then
        IsError = True
    End If
    
End Function

Public Function DoPostFoundJobs(rstRecordset As Recordset, rstExpensesPerVAT As Recordset)

    On Error GoTo ErrTrap

    blnStatus = False
    
    PopulateFields rstRecordset
    PopulateGrid rstExpensesPerVAT
    DisplayTotals
    EnableOrDisableFields
    UpdateButtons Me, 5, 0, 1, 1, 0, 1, 0
        
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function


Private Function EnableOrDisableFields()

    EnableFields mskDateIssue, txtExpenseDescription, txtSupplierDescription, txtCodeShortDescriptionA, txtInvoiceNo, txtPaymentTermDescription
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)

End Function

Private Function CheckForTheLastInvoice()

    'Αν έχω φέρει το τελευταίο παραστατικό ή αν είναι χειρόγραφο
    If txtCodeLastNo.text = Int(txtInvoiceNo.text) Or chkCodeHandID.Value = 1 Then
        CheckForTheLastInvoice = True 'Μπορώ να διαγράψω
    Else
        CheckForTheLastInvoice = False 'Δεν μπορώ να διαγράψω
    End If

End Function


Private Function PopulateFields(rstRecordset As Recordset)

    With rstRecordset
    
        txtInvoiceMasterRefersTo.text = !InvoiceMasterRefersTo
        txtInvoiceSecondaryRefersTo.text = !InvoiceSecondaryRefersTo
        txtInvoiceID.text = !InvoiceID
        txtInvoiceTrnID.text = !InvoiceTrnID
        txtInvoiceCodeID.text = !InvoiceCodeID
        txtInvoicePersonID.text = !InvoicePersonID
        txtInvoiceDateIn.text = !InvoiceDateIn
        txtInvoiceInExpenseCategoryID.text = !ExpenseCategoryID
        txtInvoiceInPaymentTermID.text = !PaymentTermID
        txtCodeLastNo.text = !CodeLastNo
        txtCodeLastDate.text = !CodeLastDate
        'txtCodePersonsPlusOrMinus.text = !CodeSuppliers
        chkCodeHandID.Value = !CodeHandID
        
        mskDateIssue.text = Format(!InvoiceDateIssue, "dd/mm/yyyy")
        txtCodeShortDescriptionA.text = !CodeShortDescriptionA
        lblCodeDescription.Caption = !CodeDescription
        lblCodeBatch.Caption = !CodeBatch
        lblCodeHand.Caption = IIf(!CodeHandID, "ΧΕΙΡΟΓΡΑΦΟ", "")
        txtInvoiceNo.text = !InvoiceNo
        txtSupplierDescription.text = !Description
        txtExpenseDescription.text = !ExpenseCategoryDescription
        txtPaymentTermDescription.text = !PaymentTermDescription
        
    End With

End Function


