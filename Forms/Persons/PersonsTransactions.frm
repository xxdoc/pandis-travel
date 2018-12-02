VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form PersonsTransactions 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   10680
   ClientLeft      =   1725
   ClientTop       =   795
   ClientWidth     =   15285
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00404040&
   Icon            =   "PersonsTransactions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   10050
      TabIndex        =   12
      Top             =   3150
      Width           =   4515
      Begin VB.TextBox Text13 
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
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "CustomersOrSuppliers"
         Top             =   5325
         Width           =   3540
      End
      Begin VB.TextBox txtCustomersOrSuppliers 
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
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   5325
         Width           =   780
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "PaymentInOrPaymentOut"
         Top             =   4950
         Width           =   3540
      End
      Begin VB.TextBox txtPaymentInOrPaymentOut 
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
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   4950
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
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "6"
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
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "Invoices.InvoicePersonID"
         Top             =   1950
         Width           =   3540
      End
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
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceCodeID"
         Top             =   1575
         Width           =   3540
      End
      Begin VB.TextBox Text10 
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
         TabIndex        =   52
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
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "3"
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
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "5"
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
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "7"
         Top             =   2325
         Width           =   780
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
         TabIndex        =   48
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
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "4"
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
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceTrnID"
         Top             =   1200
         Width           =   3540
      End
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
         TabIndex        =   45
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
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "Codes.CodeLastDate"
         Top             =   3825
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
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Codes.CodeLastNo"
         Top             =   3450
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
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Codes.CodeCustomers"
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
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "12"
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
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "10"
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
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "11"
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4575
         Width           =   780
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
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "1"
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
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceMasterRefersTo"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox Text5 
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
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Invoices.InvoiceSecondaryRefersTo"
         Top             =   450
         Width           =   3540
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
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox txtPaymentBankID 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "9"
         Top             =   3075
         Width           =   780
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Payments.PaymentBankID"
         Top             =   3075
         Width           =   3540
      End
      Begin VB.TextBox txtPaymentPaymentWayID 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         Text            =   "8"
         Top             =   2700
         Width           =   780
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         Text            =   "Payments.PaymentWayID"
         Top             =   2700
         Width           =   3540
      End
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   375
      TabIndex        =   60
      Top             =   5775
      Width           =   10365
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   61
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
         TabIndex        =   62
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
         TabIndex        =   63
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
         Index           =   3
         Left            =   4500
         TabIndex        =   64
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
         Index           =   4
         Left            =   5925
         TabIndex        =   65
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
         Index           =   5
         Left            =   7350
         TabIndex        =   66
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
         Index           =   2
         Left            =   3075
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1217
         BackColor       =   12640511
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
         PicOpacity      =   0
      End
   End
   Begin UserControls.newFloat mskAmount 
      Height          =   465
      Left            =   2175
      TabIndex        =   8
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
      Alignment       =   1
      ForeColor       =   0
      Text            =   "-999.999,99"
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
   Begin UserControls.newText txtPersonDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   4
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
   Begin UserControls.newDate mskDateIssue 
      Height          =   465
      Left            =   2175
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
   Begin UserControls.newText txtCodeShortDescriptionA 
      Height          =   465
      Left            =   2175
      TabIndex        =   2
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
   Begin UserControls.newText txtInvoiceNo 
      Height          =   465
      Left            =   2175
      TabIndex        =   3
      Top             =   2175
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   0
      MaxLength       =   6
      Text            =   "999999"
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
   Begin UserControls.newText txtReason 
      Height          =   465
      Left            =   2175
      TabIndex        =   5
      Top             =   3225
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
   Begin UserControls.newText txtPaymentWayDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   6
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
      Index           =   0
      Left            =   2925
      TabIndex        =   20
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
      PicNormal       =   "PersonsTransactions.frx":0442
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   3375
      TabIndex        =   21
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
      PicNormal       =   "PersonsTransactions.frx":09DC
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7200
      TabIndex        =   22
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
      PicNormal       =   "PersonsTransactions.frx":0F76
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   7200
      TabIndex        =   23
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
      PicNormal       =   "PersonsTransactions.frx":1510
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtBankDescription 
      Height          =   465
      Left            =   2175
      TabIndex        =   7
      Top             =   4275
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
      Index           =   6
      Left            =   7200
      TabIndex        =   25
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
      PicNormal       =   "PersonsTransactions.frx":1AAA
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   7650
      TabIndex        =   30
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
      PicNormal       =   "PersonsTransactions.frx":2044
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   5
      Left            =   7650
      TabIndex        =   31
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
      PicNormal       =   "PersonsTransactions.frx":25DE
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   7
      Left            =   7650
      TabIndex        =   32
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
      PicNormal       =   "PersonsTransactions.frx":2B78
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
      Left            =   0
      Top             =   5550
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblFullNumber 
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
      Left            =   3675
      TabIndex        =   33
      Top             =   4875
      Visible         =   0   'False
      Width           =   7065
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Τράπεζα"
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
      TabIndex        =   24
      Top             =   4350
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Τρόπος είσπραξης"
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
      TabIndex        =   19
      Top             =   3825
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Αιτιολογία"
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
      Top             =   3300
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
      Index           =   4
      Left            =   450
      TabIndex        =   17
      Top             =   2250
      Width           =   1290
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
      TabIndex        =   16
      Top             =   1725
      Width           =   1290
   End
   Begin VB.Label lblCodeDescription 
      AutoSize        =   -1  'True
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
      Height          =   255
      Left            =   3825
      TabIndex        =   15
      Top             =   1650
      Width           =   4200
   End
   Begin VB.Label lblCodeBatch 
      AutoSize        =   -1  'True
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
      Height          =   255
      Left            =   5250
      TabIndex        =   14
      Top             =   1875
      Width           =   585
   End
   Begin VB.Label lblCodeHand 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ΜΗΧΑΝΟΓΡΑΦΙΚΟ"
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
      Height          =   255
      Left            =   3825
      TabIndex        =   13
      Top             =   1875
      Width           =   1350
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   900
      Top             =   6450
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   10725
      Top             =   5250
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   540
      Index           =   1
      Left            =   2700
      Top             =   5250
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
      Top             =   2250
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
      Left            =   1725
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Ποσό"
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
      Top             =   4875
      Width           =   840
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Συναλλασόμενος"
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
      Top             =   2775
      Width           =   1215
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
      TabIndex        =   9
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Κινήσεις συναλλασόμενων"
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
      Width           =   6150
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
      Left            =   3375
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "PersonsTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnStatus As Boolean
Dim blnCancel As Boolean
Dim blnPrinterHasBeenSelected As Boolean
Dim lngTrnID As Long
Dim IsError As Boolean

Private Function AbortProcedure(blnStatus)

    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtPaymentPaymentWayID, txtPaymentBankID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
            ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
            ClearFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
            DisableFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
            UpdateButtons Me, 6, 1, 0, 0, 0, 1, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()

    If MainDeleteRecord("CommonDB", "Invoices", strApplicationName, "InvoiceID", Val(txtInvoiceID.text), "True") Then
        ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtPaymentPaymentWayID, txtPaymentBankID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
        ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
        ClearFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
        DisableFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
        UpdateButtons Me, 6, 1, 0, 0, 0, 1, 0, 1
    End If
    
End Function

Private Function FindRecords()

    With PersonsTransactionsIndex
        .Tag = "True"
        .txtInvoiceMasterRefersTo.text = txtInvoiceMasterRefersTo.text
        .txtInvoiceSecondaryRefersTo.text = txtInvoiceSecondaryRefersTo.text
        .txtPaymentInOrPaymentOut.text = txtPaymentInOrPaymentOut.text
        .txtCustomersOrSuppliers.text = txtCustomersOrSuppliers.text
        .Show 1, Me
    End With

End Function

Private Function NewRecord()

    blnStatus = True
    ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtPaymentPaymentWayID, txtPaymentBankID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
    ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
    ClearFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
    EnableFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
    UpdateButtons Me, 6, 0, 1, 0, 0, 0, 1, 0
    mskDateIssue.SetFocus
    
    InitializeFields mskDateIssue, mskAmount, txtPaymentBankID
    
    mskDateIssue.text = Date
    txtInvoiceDateIn.text = Date
    
End Function

Private Function PopulateHelperFields(strPaymentInOrPaymentOut, strCustomersOrSuppliers)

    txtPaymentInOrPaymentOut.text = strPaymentInOrPaymentOut
    txtCustomersOrSuppliers.text = strCustomersOrSuppliers

End Function

Private Function SavePayment()

    If IsError Then Exit Function
    
    If MainSaveRecord("CommonDB", txtPaymentInOrPaymentOut.text, blnStatus, strApplicationName, "TrnID", txtInvoiceTrnID.text, txtInvoiceTrnID.text, _
        txtReason.text, _
        txtPaymentPaymentWayID.text, _
        txtPaymentBankID.text, _
        mskAmount.text) <> 0 Then
        IsError = False
    Else
        IsError = True
    End If

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    txtPaymentBankID.text = IIf(txtPaymentBankID.text = "", 0, txtPaymentBankID.text)
    
    BeginTrans
    
    SaveInvoice
    SavePayment
    UpdateCodes
    
    If IsError Then
        DisplayErrorMessage True, strStandardMessages(5)
        Rollback
        Exit Function
    Else
        CommitTrans
        blnCancel = True
        AskToPrintReceipt
        ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtPaymentPaymentWayID, txtPaymentBankID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
        ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
        ClearFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
        DisableFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
        UpdateButtons Me, 6, 1, 0, 0, 0, IIf(CheckForLoadedForm("InvoicesOutIndex"), 0, 1), 0, 1
    End If

End Function

Private Function AskToPrintReceipt()

    Dim arrDummy()
    
    'Ερώτηση για εκτύπωση αν τυπώνεται και αν είμαι σε νέα εγγραφή
    If chkCodeHandID.Value = 0 And blnStatus Then
        If MyMsgBox(2, strApplicationName, strAppMessages(7), 2) Then
            ProcessSelectedReceiptsForPrinting txtInvoiceTrnID.text, arrDummy
        End If
    End If

End Function

Private Function SaveInvoice()

    If blnStatus Then txtInvoiceTrnID.text = AddOneToTheLastRecord("Invoices")
    
    If MainSaveRecord("CommonDB", "Invoices", blnStatus, strApplicationName, "InvoiceID", txtInvoiceID.text, txtInvoiceTrnID.text, txtInvoiceMasterRefersTo.text, txtInvoiceSecondaryRefersTo.text, mskDateIssue.text, txtInvoiceDateIn.text, txtInvoiceCodeID.text, txtInvoiceNo.text, txtInvoicePersonID.text, strCurrentUser) <> 0 Then
        IsError = False
    Else
        IsError = True
    End If

End Function

Private Function UpdateCodes()

    If IsError Then Exit Function
    
    Dim rsCodes As Recordset
    
    If blnStatus And chkCodeHandID.Value = 0 Then
        Set rsCodes = CommonDB.OpenRecordset("Codes")
        With rsCodes
            .index = "CodeID"
            .Seek "=", txtInvoiceCodeID.text
            If Not .NoMatch Then
                .Edit
                !CodeLastNo = Int(txtInvoiceNo.text)
                !CodeLastDate = mskDateIssue.text
                .Update
            End If
        End With
    End If

End Function

Public Function SeekRecord(lngTrnID, strPaymentInOrPaymentOut, strCustomersOrSuppliers)

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
    
    Dim rstRecordset As Recordset
    
    strSQL = "SELECT " _
        & "Invoices.InvoiceID, Invoices.InvoiceTrnID, Invoices.InvoiceMasterRefersTo, Invoices.InvoiceSecondaryRefersTo, Invoices.InvoiceDateIssue, Invoices.InvoiceDateIn, Invoices.InvoiceCodeID, Invoices.InvoiceNo, Invoices.InvoicePersonID, Invoices.InvoiceDateIssue, Invoices.InvoiceNo, " _
        & strPaymentInOrPaymentOut & ".Reason, " & strPaymentInOrPaymentOut & ".PaymentWayID, " & strPaymentInOrPaymentOut & ".BankID, " & strPaymentInOrPaymentOut & ".Amount, " _
        & "PaymentWays.PaymentWayDescription, " _
        & "Codes.CodeShortDescriptionA, Codes.CodeDescription, Codes.CodeBatch, Codes.CodeHandID, Codes.Code" & strCustomersOrSuppliers & " , Codes.CodeLastNo, Codes.CodeLastDate, " _
        & strCustomersOrSuppliers & ".Description, " _
        & "Banks.BankDescription " _
        & "FROM (((((Invoices " _
        & "INNER JOIN " & strPaymentInOrPaymentOut & " ON Invoices.InvoiceTrnID = " & strPaymentInOrPaymentOut & ".TrnID) " _
        & "INNER JOIN Codes ON Invoices.InvoiceCodeID = Codes.CodeID) " _
        & "INNER JOIN " & strCustomersOrSuppliers & " ON Invoices.InvoicePersonID = " & strCustomersOrSuppliers & ".ID) " _
        & "INNER JOIN PaymentWays ON " & strPaymentInOrPaymentOut & ".PaymentWayID = PaymentWays.PaymentWayID) " _
        & "LEFT JOIN Banks ON " & strPaymentInOrPaymentOut & ".BankID = Banks.BankID) "
        
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
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Ημερομηνία
    If Not CheckDate(mskDateIssue.text, strApplicationName) Then
        mskDateIssue.SetFocus
        Exit Function
    End If
    
    'Καταχώρηση σε ημερομηνία μεγαλύτερη από σήμερα
    If CDate(mskDateIssue.text) > Date Then
        If MyMsgBox(4, strApplicationName, strAppMessages(5), 1) Then
        End If
        mskDateIssue.SetFocus
        Exit Function
    End If
    
    'Παραστατικό
    If txtInvoiceCodeID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtCodeShortDescriptionA.SetFocus
        Exit Function
    End If
    
    'Νο παραστατικού
    If txtInvoiceNo.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtInvoiceNo.SetFocus
        Exit Function
    End If
    
    'Νο παραστατικού = αριθμός
    If Not IsNumeric(txtInvoiceNo.text) Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtInvoiceNo.SetFocus
        Exit Function
    End If
    
    'Μηχανογραφικό στοιχείο ήδη καταχωρημένο: Ελέγχω αν το νούμερο του στοιχείου υπάρχει ήδη στην χρήση
    If chkCodeHandID.Value = 0 Then
        If CheckForDuplicateInvoice(mskDateIssue.text, txtInvoiceCodeID.text, txtInvoiceNo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(22), 1) Then
            End If
            txtCodeShortDescriptionA.SetFocus
            Exit Function
        End If
    End If
    
    'Μηχανογραφικό στοιχείο: Εχω καταχωρήσει σε μεταγενέστερη ημερομηνία
    If chkCodeHandID.Value = 0 Then
        If CDate(mskDateIssue.text) < CDate(txtCodeLastDate) Then
            If MyMsgBox(4, strApplicationName, strAppMessages(4) & txtCodeLastDate.text, 1) Then
            End If
            mskDateIssue.SetFocus
            Exit Function
        End If
    End If
    
    'Πελάτης
    If txtInvoicePersonID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPersonDescription.SetFocus
        Exit Function
    End If
    
    'Τρόπος είσπραξης
    If txtPaymentPaymentWayID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPaymentWayDescription.SetFocus
        Exit Function
    End If
    
    'Τράπεζα
    If txtPaymentBankID.text = "" Then
        txtPaymentBankID.text = "0"
    End If
    
    'Ποσό
    If mskAmount.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskAmount.SetFocus
        Exit Function
    End If
    
    'Πολύ μεγάλο ποσό
    If Val(mskAmount.text) > 9999999.99 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskAmount.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Sub cmdButton_Click(index As Integer)

    Dim arrDummy()

    Select Case index
        Case 0
            NewRecord
        Case 1
            SaveRecord
        Case 2
            ProcessSelectedReceiptsForPrinting txtInvoiceTrnID.text, arrDummy 'Called when the print button is clicked
            'PrintReceipt True, False, txtInvoiceNo.text, txtCustomersOrSuppliers.text
        Case 3
            DeleteRecord
        Case 4
            FindRecords
        Case 5
            AbortProcedure False
        Case 6
            AbortProcedure True
    End Select

End Sub

Public Function ProcessSelectedReceiptsForPrinting(strInvoiceTrnID, arrInvoicesTrnID())

    Dim intLoop As Integer
    Dim rstRecordset As Recordset
    
    If Not SelectPrinter("PrinterPrintsReports") Then Exit Function
    If Not PrinterExists(strPrinterName) Then Exit Function
    
    If strInvoiceTrnID <> "" Then
        ReDim arrInvoicesTrnID(0)
        arrInvoicesTrnID(0) = strInvoiceTrnID
    End If
    
    For intLoop = 0 To UBound(arrInvoicesTrnID)
        Set rstRecordset = SeekRecord(arrInvoicesTrnID(intLoop), txtPaymentInOrPaymentOut.text, txtCustomersOrSuppliers.text)
        If rstRecordset.RecordCount = 0 Then MyMsgBox 4, strApplicationName, strStandardMessages(9), 1: Exit Function
        PrintThisReceipt blnPreviewInvoices, False, rstRecordset!InvoiceNo 'False = Do not preview, True = Create PDF instead of print
    Next intLoop

End Function

Private Function PrintThisReceipt(blnPreview As Boolean, blnExportPDF As Boolean, strInvoiceNo As String)

    Dim intLoop As Integer
    Dim pdf As New ARExportPDF
    
    If blnExportPDF Then
        rptReceipt.Run False
        pdf.AcrobatVersion = 2
        pdf.SemiDelimitedNeverEmbedFonts = ""
        pdf.fileName = strReportsPathName & "Invoice" & Right("00000" & strInvoiceNo, 5) & ".pdf"
        pdf.Export rptReceipt.Pages
    Else
        For intLoop = 1 To intInvoiceCopies + 1
            rptReceipt.Restart
            If blnPreview Then
                rptReceipt.Zoom = -2
                rptReceipt.Printer.ColorMode = vbPRCMMonochrome
                rptReceipt.WindowState = vbMaximized
                rptReceipt.Show 1
                Exit For
            Else
                rptReceipt.Printer.DeviceName = strPrinterName
                rptReceipt.PrintReport False
                rptReceipt.Run True
            End If
        Next intLoop
    End If

End Function

Private Sub cmdIndex_Click(index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Παραστατικό - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Codes", "CodeShortDescriptionA, CodeMasterRefersTo", "String, String", txtCodeShortDescriptionA.text, txtInvoiceMasterRefersTo.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 3, True, 8, 0, 3, 5, 6, 7, 8, 10, 11, "ID", "Συντ. Α'", "Περιγραφή", "Σειρά", "Χειρόγραφο", "Πελάτες", "Τελευταίο Νο", "Ημερομηνία", 0, 6, 40, 6, 0, 0, 0, 0, 1, 1, 0, 1, 1, 1, 1, 1)
                txtInvoiceCodeID.text = tmpTableData.strCode
                txtCodeShortDescriptionA.text = tmpTableData.strFirstField
                lblCodeDescription.Caption = tmpTableData.strSecondField
                lblCodeBatch.Caption = IIf(txtInvoiceCodeID.text <> "" And tmpTableData.strThirdField <> "", " ΣΕΙΡΑ " & tmpTableData.strThirdField, "")
                chkCodeHandID.Value = IIf(tmpTableData.strFourthField = "1", 1, 0)
                lblCodeHand.Caption = IIf(tmpTableData.strFourthField = "1", "ΧΕΙΡΟΓΡΑΦΟ", "ΜΗΧΑΝΟΓΡΑΦΙΚΟ")
                txtInvoiceNo.Locked = IIf(chkCodeHandID.Value = 1, False, True)
                txtCodeLastNo.text = tmpTableData.strSixthField
                txtCodeLastDate.text = format(tmpTableData.strSeventhField, "dd/mm/yyyy")
                If txtInvoiceCodeID.text <> "" And chkCodeHandID.Value = 0 Then txtInvoiceNo.text = txtCodeLastNo.text + 1
            End If
        Case 1
            'Παραστατικό - F5
            With TablesCodes
                .Tag = "True"
                .txtCodeMasterRefersTo.text = txtInvoiceMasterRefersTo.text
                .txtCodeSecondaryRefersTo.text = txtInvoiceSecondaryRefersTo.text
                .Show 1, Me
            End With
        Case 2
            'Πελάτης - F2
            Set tmpRecordset = CheckForMatch("CommonDB", txtCustomersOrSuppliers.text, "Description", "String", txtPersonDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 7, "ID", "Επωνυμία", "Α.Φ.Μ.", 0, 40, 15, 1, 0, 1)
                txtInvoicePersonID.text = tmpTableData.strCode
                txtPersonDescription.text = tmpTableData.strFirstField
            End If
        Case 3
            'Πελάτης - F5
            With persons
                .Tag = "True"
                .Show 1, Me
            End With
        Case 4
            'Τρόπος είσπραξης / πληρωμής - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "PaymentWays", "PaymentWayDescription", "String", txtPaymentWayDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtPaymentPaymentWayID.text = tmpTableData.strCode
                txtPaymentWayDescription.text = tmpTableData.strFirstField
            End If
        Case 5
            'Τρόπος είσπραξης / πληρωμής - F5
            With TablesPaymentWays
                .Tag = "True"
                .Show 1, Me
            End With
        Case 6
            'Τράπεζα
            Set tmpRecordset = CheckForMatch("CommonDB", "Banks", "BankDescription", "String", txtBankDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtPaymentBankID.text = tmpTableData.strCode
                txtBankDescription.text = tmpTableData.strFirstField
            End If
        Case 7
            'Τράπεζα - F5
            With TablesBanks
                .Tag = "True"
                .Show 1, Me
            End With
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
        Case vbKeyP And CtrlDown = 4 And cmdButton(2).Enabled
            cmdButton_Click 2
        Case vbKeyF3 And cmdButton(3).Enabled, vbKeyD And CtrlDown = 4 And cmdButton(3).Enabled
            cmdButton_Click 3
        Case vbKeyF7 And cmdButton(4).Enabled, vbKeyF And CtrlDown = 4 And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyEscape
            If cmdButton(5).Enabled Then cmdButton_Click 5: Exit Function
            If cmdButton(6).Enabled Then cmdButton_Click 6
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    PositionControls Me, False
    ColorizeControls Me, False, False
    ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, txtInvoiceDateIn, txtPaymentPaymentWayID, txtPaymentBankID, txtCodeLastNo, txtCodeLastDate, txtCodePersonsPlusOrMinus, chkCodeHandID
    ClearFields lblCodeDescription, lblCodeBatch, lblCodeHand
    ClearFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
    DisableFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
    UpdateButtons Me, 6, 1, 0, 0, 0, 1, 0, 1

End Sub

Private Sub mskAmount_Change()

    lblFullNumber.Caption = FullNumber(mskAmount.text)

End Sub

Private Sub txtBankDescription_Change()

    If txtBankDescription.text = "" Then txtPaymentBankID.text = ""

End Sub

Private Sub txtBankDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 6
    If KeyCode = vbKeyF5 Then cmdIndex_Click 7

End Sub

Private Sub txtBankDescription_Validate(Cancel As Boolean)

    If txtBankDescription.text <> "" And txtPaymentBankID.text = "" Then cmdIndex_Click 6

End Sub

Private Sub txtCodeShortDescriptionA_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 1

End Sub

Private Sub txtPersonDescription_Change()

    If txtPersonDescription.text = "" Then txtInvoicePersonID.text = ""

End Sub

Private Sub txtPersonDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    If KeyCode = vbKeyF5 Then cmdIndex_Click 3
    
End Sub

Private Sub txtPersonDescription_Validate(Cancel As Boolean)

    If txtPersonDescription.text <> "" And txtInvoicePersonID.text = "" Then cmdIndex_Click 2

End Sub

Private Sub txtCodeShortDescriptionA_Change()

    If txtCodeShortDescriptionA.text = "" Then
        ClearFields txtInvoiceID, txtInvoiceTrnID, txtInvoiceCodeID, txtInvoicePersonID, lblCodeDescription, lblCodeBatch, lblCodeHand
    End If

End Sub

Private Sub txtCodeShortDescriptionA_Validate(Cancel As Boolean)

    If txtCodeShortDescriptionA.text <> "" And txtInvoiceCodeID.text = "" Then cmdIndex_Click 0

End Sub

Private Sub txtPaymentWayDescription_Change()

    If txtPaymentWayDescription.text = "" Then txtPaymentPaymentWayID.text = ""

End Sub

Private Sub txtPaymentWayDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4
    If KeyCode = vbKeyF5 Then cmdIndex_Click 5

End Sub

Private Sub txtPaymentWayDescription_Validate(Cancel As Boolean)

    If txtPaymentWayDescription.text <> "" And txtPaymentPaymentWayID.text = "" Then cmdIndex_Click 4

End Sub

Public Function DoPostFoundJobs(rstRecordset As Recordset, strPaymentInOrPaymentOut, strCustomersOrSuppliers)

    On Error GoTo ErrTrap

    blnStatus = False
    
    PopulateFields rstRecordset
    PopulateHelperFields strPaymentInOrPaymentOut, strCustomersOrSuppliers
    EnableOrDisableFields
    UpdateButtons Me, 6, IIf(chkCodeHandID.Value = 1, 0, 1), IIf(chkCodeHandID.Value = 1, 1, 0), IIf(chkCodeHandID.Value = 1, 0, 1), IIf(CheckForTheLastInvoice = 0, 0, 1), IIf(CheckForLoadedForm("InvoicesOutIndex"), 0, 1), IIf(chkCodeHandID.Value = 1, 1, 0), IIf(chkCodeHandID.Value = 1, 0, 1)
        
    Exit Function
    
ErrTrap:
    DisplayErrorMessage True, Err.Description

End Function



Private Function EnableOrDisableFields()

    If chkCodeHandID.Value = 1 Then
        EnableFields mskDateIssue, txtCodeShortDescriptionA, txtInvoiceNo, txtPersonDescription, txtReason, txtPaymentWayDescription, txtBankDescription, mskAmount
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7)
    End If

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
        txtPaymentPaymentWayID.text = !PaymentWayID
        txtPaymentBankID.text = IIf(IsNull(!BankID), "", !BankID)
        txtCodeLastNo.text = !CodeLastNo
        txtCodeLastDate.text = !CodeLastDate
        chkCodeHandID.Value = !CodeHandID
        
        mskDateIssue.text = format(!InvoiceDateIssue, "dd/mm/yyyy")
        txtCodeShortDescriptionA.text = !CodeShortDescriptionA
        lblCodeDescription.Caption = !CodeDescription
        lblCodeBatch.Caption = IIf(!CodeBatch <> "", "ΣΕΙΡΑ " & !CodeBatch, "")
        lblCodeHand.Caption = IIf(!CodeHandID, "ΧΕΙΡΟΓΡΑΦΟ", "ΜΗΧΑΝΟΓΡΑΦΙΚΟ")
        txtInvoiceNo.text = !InvoiceNo
        txtPersonDescription.text = !Description
        txtReason.text = !Reason
        txtPaymentWayDescription.text = !PaymentWayDescription
        txtBankDescription.text = IIf(IsNull(!BankDescription), "", !BankDescription)
        mskAmount.text = format(!Amount, "#,##0.00")
        
    End With

End Function

