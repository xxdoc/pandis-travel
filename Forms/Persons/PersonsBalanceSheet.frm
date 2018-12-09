VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "ProgressBar.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form PersonsBalanceSheet 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   10875
   ClientLeft      =   15
   ClientTop       =   15
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
      Left            =   12825
      TabIndex        =   28
      Top             =   7650
      Width           =   4065
      Begin vbalProgBarLib6.vbalProgressBar prgProgressBar 
         Height          =   615
         Left            =   150
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1085
         Picture         =   "PersonsBalanceSheet.frx":0000
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "PersonsBalanceSheet.frx":001C
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
         TabIndex        =   30
         Top             =   75
         Width           =   3765
      End
   End
   Begin VB.Frame frmContainer 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
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
         TabIndex        =   17
         Top             =   8850
         Width           =   10290
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   0
            Left            =   225
            TabIndex        =   18
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
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   6
            Left            =   8775
            TabIndex        =   19
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
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   1
            Left            =   1650
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Καρτέλα"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   5
            Left            =   7350
            TabIndex        =   21
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
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   2
            Left            =   3080
            TabIndex        =   22
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
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   3
            Left            =   4500
            TabIndex        =   23
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
            ForeColor       =   8388736
            PicOpacity      =   0
         End
         Begin Dacara_dcButton.dcButton cmdButton 
            Height          =   690
            Index           =   4
            Left            =   5925
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1217
            ButtonShape     =   3
            ButtonStyle     =   4
            Caption         =   "Δημιουργία αρχείου XLS"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ubuntu Condensed"
               Size            =   9.75
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388736
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
         Left            =   7650
         TabIndex        =   10
         Top             =   6525
         Width           =   5040
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
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   75
            Width           =   1305
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
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "3"
            Top             =   825
            Width           =   1305
         End
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
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "CustomersOrSuppliers"
            Top             =   825
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
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   450
            Width           =   1305
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
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "PaymentInOrPaymentOut"
            Top             =   450
            Width           =   3540
         End
         Begin VB.TextBox txtFilterID 
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
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1305
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
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "RecordFilter.RecordFilterID"
            Top             =   1200
            Width           =   3540
         End
         Begin vbalIml6.vbalImageList lstIconList 
            Left            =   75
            Top             =   1575
            _ExtentX        =   953
            _ExtentY        =   953
            Size            =   4592
            Images          =   "PersonsBalanceSheet.frx":0038
            Version         =   131072
            KeyCount        =   4
            Keys            =   "ÿÿÿ"
         End
      End
      Begin VB.Frame frmCriteria 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   2640
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   6075
         Width           =   7440
         Begin UserControls.newDate mskInvoiceDateIssueFrom 
            Height          =   465
            Left            =   1575
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
         Begin UserControls.newDate mskInvoiceDateIssueTo 
            Height          =   465
            Left            =   3075
            TabIndex        =   2
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
         Begin UserControls.newText txtFilterDescription 
            Height          =   465
            Left            =   1575
            TabIndex        =   3
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   6600
            TabIndex        =   14
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
            PicNormal       =   "PersonsBalanceSheet.frx":1248
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   3
            Left            =   2175
            Top             =   1800
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   315
            Index           =   4
            Left            =   2400
            Top             =   525
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Εγγραφές"
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
            TabIndex        =   13
            Top             =   1425
            Width           =   690
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   2
            Left            =   6975
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
            Index           =   1
            Left            =   1125
            Top             =   675
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
            Top             =   675
            Visible         =   0   'False
            Width           =   465
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
            Left            =   4350
            TabIndex        =   12
            Top             =   75
            Width           =   2940
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
            TabIndex        =   11
            Top             =   0
            Width           =   7440
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
            TabIndex        =   8
            Top             =   2100
            Width           =   7440
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Εκδοση"
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
            TabIndex        =   7
            Top             =   900
            Width           =   690
         End
      End
      Begin iGrid300_10Tec.iGrid grdPersonsBalanceSheet 
         Height          =   7290
         Left            =   75
         TabIndex        =   4
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
         Left            =   3975
         TabIndex        =   27
         Top             =   1125
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
         TabIndex        =   26
         Top             =   525
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
         TabIndex        =   25
         Top             =   825
         Width           =   14940
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
         TabIndex        =   24
         Top             =   1125
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ισοζύγιο συναλλασόμενων"
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
         TabIndex        =   5
         Top             =   75
         Width           =   5850
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
Attribute VB_Name = "PersonsBalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngRowCount As Long
Dim blnError As Boolean
Dim blnProcessing As Boolean

'Προηγούμενη περίοδος
    'Γραμμή κάθε πελάτη
    Dim curDebitSoFar As Currency
    Dim curCreditSoFar As Currency
    Dim curBalanceSoFar As Currency
    'Σύνολα (Τελευταία γραμμή)
    Dim curDebitTotalSoFar As Currency
    Dim curCreditTotalSoFar As Currency
    Dim curBalanceTotalSoFar As Currency
'Ζητούμενη περίοδος
    'Ποσά κάθε πελάτη
    Dim curDebitPeriod As Currency
    Dim curCreditPeriod As Currency
    'Σύνολα (Τελευταία γραμμή)
    Dim curDebitTotalPeriod As Currency
    Dim curCreditTotalPeriod As Currency
'Υπόλοιπα
    'Γραμμή κάθε πελάτη
    Dim curBalance As Currency
    'Γενικά σύνολα (Τελευταία γραμμή)
    Dim curBalanceTotal As Currency

Private Function AddCurrentLineToGrid(rstPersons As Recordset, lngLastRow As Long)

    'Τι έχω δώσει στα κριτήρια
    If txtFilterID.text = "1" Or (txtFilterID.text = "2" And curBalanceSoFar + curBalance <> 0) Then
        'Προσθέτω μία γραμμή
        grdPersonsBalanceSheet.AddRow
        lngRowCount = lngRowCount + 1
        'Ενημερώνω τη γραμμή
        grdPersonsBalanceSheet.CellValue(lngLastRow, "ID") = rstPersons!ID
        grdPersonsBalanceSheet.CellValue(lngLastRow, "Description") = rstPersons!Description
        'Προηγούμενη περιόδος
        grdPersonsBalanceSheet.CellValue(lngLastRow, "DebitSoFar") = curDebitSoFar
        grdPersonsBalanceSheet.CellValue(lngLastRow, "CreditSoFar") = curCreditSoFar
        grdPersonsBalanceSheet.CellValue(lngLastRow, "BalanceSoFar") = curBalanceSoFar
        'Ζητούμενη περίοδος
        grdPersonsBalanceSheet.CellValue(lngLastRow, "DebitPeriod") = curDebitPeriod
        grdPersonsBalanceSheet.CellValue(lngLastRow, "CreditPeriod") = curCreditPeriod
        'Υπόλοιπο
        grdPersonsBalanceSheet.CellValue(lngLastRow, "Balance") = curBalanceSoFar + curBalance
        'Υπολογίζω τα γενικά σύνολα (τελευταία γραμμή)
        curDebitTotalSoFar = curDebitTotalSoFar + curDebitSoFar
        curCreditTotalSoFar = curCreditTotalSoFar + curCreditSoFar
        curBalanceTotalSoFar = curBalanceTotalSoFar + curBalanceSoFar
        curDebitTotalPeriod = curDebitTotalPeriod + curDebitPeriod
        curCreditTotalPeriod = curCreditTotalPeriod + curCreditPeriod
        curBalanceTotal = curBalanceTotal + curBalanceSoFar + curBalance
        'Χρωματίζω
        InvertColorForNegativeNumbers grdPersonsBalanceSheet, grdPersonsBalanceSheet.RowCount
    End If

End Function

Private Function AddGrandTotalsToGrid(lngLastRow As Long)

    With grdPersonsBalanceSheet
        If .RowCount > 0 Then
            'Προσθέτω τα γενικά σύνολα (τελευταία γραμμή)
            .AddRow
            .AddRow
            .CellValue(lngLastRow, "Description") = "ΓΕΝΙΚΑ ΣΥΝΟΛΑ"
            'Προηγούμενη περιόδος
            .CellValue(lngLastRow, "DebitSoFar") = curDebitTotalSoFar
            .CellValue(lngLastRow, "CreditSoFar") = curCreditTotalSoFar
            .CellValue(lngLastRow, "BalanceSoFar") = curBalanceTotalSoFar
            'Ζητούμενη περίοδος
            .CellValue(lngLastRow, "DebitPeriod") = curDebitTotalPeriod
            .CellValue(lngLastRow, "CreditPeriod") = curCreditTotalPeriod
            'Υπόλοιπο
            .CellValue(lngLastRow, "Balance") = curBalanceTotal
            'Χρωματίζω
            InvertColorForNegativeNumbers grdPersonsBalanceSheet, grdPersonsBalanceSheet.RowCount
        End If
    End With

End Function

Private Function CalculateAskedPeriodForCustomers(rstTransactions As Recordset)

    'Helper
    Dim curTotals As Currency
    
    With rstTransactions
        'Πώληση (Χρεωστική ή πιστωτική) - Στήλη χρέωσης
        If !InvoiceMasterRefersTo = "2" Then
            'Helper
            curTotals = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsAmountWithTransfer", "InvoiceOutAdultsAmountWithoutTransfer", "InvoiceOutKidsAmountWithTransfer", "InvoiceOutKidsAmountWithoutTransfer", "InvoiceOutDirectAmount")
            'Adults, Kids και Direct
            curDebitPeriod = curDebitPeriod + curTotals
            'Αν η κίνηση είναι μετρητοίς βάζω το ποσό και στην πίστωση
            curCreditPeriod = IIf(!PaymentTermCreditID = 0, curCreditPeriod + curTotals, curCreditPeriod)
        End If
        'Είσπραξη ή πληρωμή - Στήλη πίστωσης
        If !InvoiceMasterRefersTo = "4" Then
            If !CodeCustomers = "+" Then
                curDebitPeriod = curDebitPeriod + CalculateFields(rstTransactions, !CodeCustomers, "Amount")
            End If
            If !CodeCustomers = "-" Then
                curCreditPeriod = curCreditPeriod + Abs(CalculateFields(rstTransactions, !CodeCustomers, "Amount"))
            End If
        End If
    End With
    
    'Υπόλοιπο
    curBalance = curDebitPeriod - curCreditPeriod

End Function

Private Function CalculateAskedPeriodForSuppliers(rstTransactions As Recordset)

    'Helper
    Dim curTotals As Currency
    
    With rstTransactions
        'Εξοδο (Χρεωστική ή πιστωτική) - Στήλη πίστωσης
        If !InvoiceMasterRefersTo = "1" Then
            'Helper
            curTotals = CalculateFields(rstTransactions, !CodeSuppliers, "InvoiceInAmount")
            'Πίστωση
            curCreditPeriod = curCreditPeriod + curTotals
            'Χρέωση
            curDebitPeriod = IIf(!PaymentTermCreditID = 0, curDebitPeriod + curTotals, curDebitPeriod)
        End If
        'Είσπραξη ή πληρωμή - Στήλη χρέωσης
        If !InvoiceMasterRefersTo = "3" Then
            'Helper
            curTotals = CalculateFields(rstTransactions, !CodeSuppliers, "Amount")
            'Χρέωση
            If !CodeSuppliers = "-" Then
                curDebitPeriod = curDebitPeriod + Abs(curTotals)
            End If
            'Πίστωση
            If !CodeSuppliers = "+" Then
                curCreditPeriod = curCreditPeriod + Abs(curTotals)
            End If
        End If
    End With
    
    'Υπόλοιπο
    curBalance = curDebitPeriod - curCreditPeriod

End Function


Private Function CalculateSoFarTotalsForSuppliers(rstTransactions As Recordset)

    'Helper
    Dim curTotals As Currency
    
    With rstTransactions
        Do Until !InvoiceDateIssue >= CDate(mskInvoiceDateIssueFrom.text)
            'Εξοδο (Χρεωστικό ή πιστωτικό) - Στήλη πίστωσης
            If !InvoiceMasterRefersTo = "1" Then
                'Helper
                curTotals = CalculateFields(rstTransactions, !CodeSuppliers, "InvoiceInAmount")
                'Πίστωση
                curCreditSoFar = curCreditSoFar + curTotals
                'Χρέωση
                curDebitSoFar = IIf(!PaymentTermCreditID = 0, curDebitSoFar + curTotals, curDebitSoFar)
            End If
            'Είσπραξη ή πληρωμή - Στήλη χρέωσης
            If !InvoiceMasterRefersTo = "3" Then
                'Helper
                curTotals = CalculateFields(rstTransactions, !CodeSuppliers, "Amount")
                'Χρέωση
                If !CodeSuppliers = "-" Then
                    curDebitSoFar = curDebitSoFar + Abs(curTotals)
                End If
                'Πίστωση
                If !CodeSuppliers = "+" Then
                    curCreditSoFar = curCreditSoFar + Abs(curTotals)
                End If
            End If
            'Υπόλοιπο
            curBalanceSoFar = curDebitSoFar - curCreditSoFar
            rstTransactions.MoveNext
            'Async!
            DoEvents
            If rstTransactions.EOF Then Exit Do
        Loop
    End With

End Function

Private Function CalculateSoFarTotals(rstTransactions As Recordset)

    If txtInvoiceMasterRefersTo.text = "1" Then CalculateSoFarTotalsForSuppliers rstTransactions
    If txtInvoiceMasterRefersTo.text = "2" Then CalculateSoFarTotalsForCustomers rstTransactions

End Function
Private Function CalculateSoFarTotalsForCustomers(rstTransactions As Recordset)

    'Helper
    Dim curTotals As Currency
    
    With rstTransactions
        Do Until !InvoiceDateIssue >= CDate(mskInvoiceDateIssueFrom.text)
            'Πώληση (Χρεωστική ή πιστωτική) - Στήλη χρέωσης
            If !InvoiceMasterRefersTo = "2" Then
                'Helper
                curTotals = CalculateFields(rstTransactions, !CodeCustomers, "InvoiceOutAdultsAmountWithTransfer", "InvoiceOutAdultsAmountWithoutTransfer", "InvoiceOutKidsAmountWithTransfer", "InvoiceOutKidsAmountWithoutTransfer", "InvoiceOutDirectAmount")
                'Adults, Kids και Direct
                curDebitSoFar = curDebitSoFar + curTotals
                'Αν η κίνηση είναι μετρητοίς βάζω το ποσό και στην πίστωση
                curCreditSoFar = IIf(!PaymentTermCreditID = 0, curCreditSoFar + curTotals, curCreditSoFar)
            End If
            'Είσπραξη ή πληρωμή - Στήλη πίστωσης
            If !InvoiceMasterRefersTo = "4" Then
                If !CodeCustomers = "+" Then
                    curDebitSoFar = curDebitSoFar + CalculateFields(rstTransactions, !CodeCustomers, "Amount")
                End If
                If !CodeCustomers = "-" Then
                    curCreditSoFar = curCreditSoFar + Abs(CalculateFields(rstTransactions, !CodeCustomers, "Amount"))
                End If
            End If
            'Υπόλοιπο
            curBalanceSoFar = curDebitSoFar - curCreditSoFar
            rstTransactions.MoveNext
            'Async!
            DoEvents
            If rstTransactions.EOF Then Exit Do
        Loop
    End With

End Function
Private Function FindRecordsAndPopulateGrid()

    If ValidateFields Then
        If RefreshList > 0 Then
            UpdateRecordCount lblRecordCount, lngRowCount
            UpdateCriteriaLabels mskInvoiceDateIssueFrom.text, mskInvoiceDateIssueTo.text, txtFilterDescription.text
            EnableGrid grdPersonsBalanceSheet, False
            HighlightRow grdPersonsBalanceSheet, 1, 1, "", True
            UpdateButtons Me, 6, 0, 0, 1, 1, 1, 1, 0
            Exit Function
        Else
            UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
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
            mskInvoiceDateIssueFrom.SetFocus
        End If
    End If

End Function

Private Function GetProcessID(programName)

    Dim sTemp As String
    
    sTemp = UCase(Trim$(programName))
    If sTemp = "" Then Exit Function
    
    glPid = Shell(sTemp, vbNormalFocus)
    
    If glPid = 0 Then
        MsgBox "Could not start process", vbExclamation, "Error"
    End If

End Function

Private Function UpdateCriteriaLabels(InvoiceDateIssueFrom, InvoiceDateIssueTo, FilterDescription)

    Dim strCriteriaA As String

    strCriteriaA = IIf(InvoiceDateIssueFrom = "", "Από [ ΟΛΑ ] ", "Από [ " & InvoiceDateIssueFrom & " ] ")
    strCriteriaA = strCriteriaA & IIf(InvoiceDateIssueTo = "", "Εως [ ΟΛΑ ] ", "Εως [ " & InvoiceDateIssueTo & " ] ")
    strCriteriaA = strCriteriaA & "Εγγραφές [ " & FilterDescription & " ]"
    
    lblCriteria.Caption = strCriteriaA
    
End Function

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            FindRecordsAndPopulateGrid
        Case 1
            'EditRecord
        Case 2
            DoReport "Print"
        Case 3
            DoReport "CreatePDF"
        Case 4
            ExportToExcel
        Case 5
            AbortProcedure False
        Case 6
            AbortProcedure True
    End Select
   
End Sub
Private Function ExportToExcel()

    On Error GoTo ErrTrap
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim xlsRowOffsetFromTop As Long
    Dim xlsColCount As Long
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object

    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    xlsRowOffsetFromTop = 10
    xlsColCount = 7
    
    With oSheet
    
        SetFontNameAndSize oSheet, "Ubuntu Condensed", 11
        AddCompanyData oSheet, xlsColCount
        AddTitle oSheet, lblTitle.Caption, xlsColCount
        AddCriteria oSheet, lblCriteria.Caption, xlsColCount
        AddHeaders oSheet, grdPersonsBalanceSheet, xlsColCount, "A", "Description", "B", "DebitSoFar", "C", "CreditSoFar", "D", "BalanceSoFar", "E", "DebitPeriod", "F", "CreditPeriod", "G", "Balance"
        AdjustColumnWidths oSheet, "A", 60, "B", 15, "C", 15, "D", 15, "E", 15, "F", 15, "G", 15
                
        For lngRow = 1 To grdPersonsBalanceSheet.RowCount
            .Range("A" & lngRow + xlsRowOffsetFromTop) = grdPersonsBalanceSheet.CellValue(lngRow, "Description")
            .Range("B" & lngRow + xlsRowOffsetFromTop) = grdPersonsBalanceSheet.CellValue(lngRow, "DebitSoFar")
            .Range("C" & lngRow + xlsRowOffsetFromTop) = grdPersonsBalanceSheet.CellValue(lngRow, "CreditSoFar")
            .Range("D" & lngRow + xlsRowOffsetFromTop) = grdPersonsBalanceSheet.CellValue(lngRow, "BalanceSoFar")
            .Range("E" & lngRow + xlsRowOffsetFromTop) = grdPersonsBalanceSheet.CellValue(lngRow, "DebitPeriod")
            .Range("F" & lngRow + xlsRowOffsetFromTop) = grdPersonsBalanceSheet.CellValue(lngRow, "CreditPeriod")
            .Range("G" & lngRow + xlsRowOffsetFromTop) = grdPersonsBalanceSheet.CellValue(lngRow, "Balance")
        Next lngRow
        
        AddNumberFormats oSheet, grdPersonsBalanceSheet, "Floats", 10, "B", "C", "D", "E", "F", "G"
    
    End With
    
    oBook.SaveAs strReportsPathName & lblTitle.Caption & ".xls"
    
    oExcel.Quit
    
    grdPersonsBalanceSheet.SetFocus
    
    MyMsgBox 1, strApplicationName, strStandardMessages(8), 1
    
    Exit Function
    
ErrTrap:
    oBook.Close False
    oExcel.Quit

    grdPersonsBalanceSheet.SetFocus
    
    If Err.Number = 1004 Then
        MyMsgBox 4, strApplicationName, strStandardMessages(27), 1
    Else
        DisplayErrorMessage True, Err.Description
    End If
    
    Exit Function
    
End Function

Private Function DoReport(action As String)
    
    On Error GoTo ErrTrap
    
    If action = "Print" Then
        If SelectPrinter("PrinterPrintsReports") Then
            CreateUnicodeFile lblTitle.Caption, " από " & mskInvoiceDateIssueFrom.text & " έως " & mskInvoiceDateIssueTo.text, "", intPrinterReportDetailLines
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
    
    If action = "CreatePDF" Then
        CreateUnicodeFile lblTitle.Caption, " από " & mskInvoiceDateIssueFrom.text & " έως " & mskInvoiceDateIssueTo.text, "", GetSetting(strApplicationName, "Settings", "Export Report Height")
        CreateUnisexPDF lblTitle.Caption & " από " & mskInvoiceDateIssueFrom.text & " έως " & mskInvoiceDateIssueTo.text
        If MyMsgBox(1, strApplicationName, strStandardMessages(8), 1) Then
        End If
    End If
    
    Exit Function
    
ErrTrap:
    Close #1
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function CreateUnicodeFile(strReportTitle, strReportSubTitle1, strReportSubTitle2, intReportDetailLines)

    'Εκτυπωτής
    Dim lngRow As Long
    Dim intProcessedDetailLines As Integer
    Dim intPageNo As Integer
    
    'Μετρητές
    Dim curDebitSoFar As Currency
    Dim curCreditSoFar As Currency
    Dim curBalanceSoFar As Currency
    Dim curDebitPeriod As Currency
    Dim curCreditPeriod As Currency
    Dim curBalance As Currency

    'Αρχικές τιμές
    curDebitSoFar = 0
    curCreditSoFar = 0
    curBalanceSoFar = 0
    curDebitPeriod = 0
    curCreditPeriod = 0
    curBalance = 0
    intPageNo = 1
    
    intPageNo = 0
    intProcessedDetailLines = 0
    
    Open strUnicodeFile For Output As #1
    GoSub Headers
    
    'Εγγραφές
    With grdPersonsBalanceSheet
        For lngRow = 1 To grdPersonsBalanceSheet.RowCount
            
            'Εκτυπώνω τη γραμμή
            Print #1, .CellText(lngRow, "Description"); _
                Tab(55 - Len((format(.CellText(lngRow, "DebitSoFar"), "#,##0.00")))); format(.CellText(lngRow, "DebitSoFar"), "#,##0.00"); _
                Tab(69 - Len((format(.CellText(lngRow, "CreditSoFar"), "#,##0.00")))); format(.CellText(lngRow, "CreditSoFar"), "#,##0.00"); _
                Tab(83 - Len((format(.CellText(lngRow, "BalanceSoFar"), "#,##0.00")))); format(.CellText(lngRow, "BalanceSoFar"), "#,##0.00"); _
                Tab(97 - Len((format(.CellText(lngRow, "DebitPeriod"), "#,##0.00")))); format(.CellText(lngRow, "DebitPeriod"), "#,##0.00"); _
                Tab(111 - Len((format(.CellText(lngRow, "CreditPeriod"), "#,##0.00")))); format(.CellText(lngRow, "CreditPeriod"), "#,##0.00"); _
                Tab(125 - Len((format(.CellText(lngRow, "Balance"), "#,##0.00")))); format(.CellText(lngRow, "Balance"), "#,##0.00")
            
            intProcessedDetailLines = intProcessedDetailLines + 1
            
            'Eject
            If intProcessedDetailLines > Val(intReportDetailLines) Then
                Print #1, ""
                Print #1, "ΣΕ ΜΕΤΑΦΟΡΑ"; _
                    Tab(55 - Len(format(curDebitSoFar, "#,##0.00"))); format(curDebitSoFar, "#,##0.00"); _
                    Tab(69 - Len(format(curCreditSoFar, "#,##0.00"))); format(curCreditSoFar, "#,##0.00"); _
                    Tab(83 - Len(format(curBalanceSoFar, "#,##0.00"))); format(curBalanceSoFar, "#,##0.00"); _
                    Tab(97 - Len(format(curDebitPeriod, "#,##0.00"))); format(curDebitPeriod, "#,##0.00"); _
                    Tab(111 - Len(format(curCreditPeriod, "#,##0.00"))); format(curCreditPeriod, "#,##0.00"); _
                    Tab(125 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                    
                GoSub Headers

                Print #1, "ΑΠΟ ΜΕΤΑΦΟΡΑ"; _
                    Tab(55 - Len(format(curDebitSoFar, "#,##0.00"))); format(curDebitSoFar, "#,##0.00"); _
                    Tab(69 - Len(format(curCreditSoFar, "#,##0.00"))); format(curCreditSoFar, "#,##0.00"); _
                    Tab(83 - Len(format(curBalanceSoFar, "#,##0.00"))); format(curBalanceSoFar, "#,##0.00"); _
                    Tab(97 - Len(format(curDebitPeriod, "#,##0.00"))); format(curDebitPeriod, "#,##0.00"); _
                    Tab(111 - Len(format(curCreditPeriod, "#,##0.00"))); format(curCreditPeriod, "#,##0.00"); _
                    Tab(125 - Len(format(curBalance, "#,##0.00"))); format(curBalance, "#,##0.00")
                Print #1, ""
                intProcessedDetailLines = intProcessedDetailLines + 2
            End If
            
            'Σύνολα
            If .CellText(lngRow, "ID") <> "" Then
                curDebitSoFar = curDebitSoFar + .CellValue(lngRow, "DebitSoFar")
                curCreditSoFar = curCreditSoFar + .CellValue(lngRow, "CreditSoFar")
                curBalanceSoFar = curBalanceSoFar + .CellValue(lngRow, "BalanceSoFar")
                curDebitPeriod = curDebitPeriod + .CellValue(lngRow, "DebitPeriod")
                curCreditPeriod = curCreditPeriod + .CellValue(lngRow, "CreditPeriod")
                curBalance = curBalance + .CellValue(lngRow, "Balance")
            End If
            
        Next lngRow
    End With
    
    Close #1
    
    Exit Function
    
Headers:
    intPageNo = intPageNo + 1
    PrintHeadings 124, intPageNo, strReportTitle, strReportSubTitle1, strReportSubTitle2
    PrintColumnHeadings 42, "--------- ΠΡΟΗΓΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ ---------- ---------- ΖΗΤΟΥΜΕΝΗ ΠΕΡΙΟΔΟΣ -----------"
    PrintColumnHeadings 1, "ΕΠΩΝΥΜΙΑ", 49, "ΧΡΕΩΣΗ       ΠΙΣΤΩΣΗ      ΥΠΟΛΟΙΠΟ        ΧΡΕΩΣΗ       ΠΙΣΤΩΣΗ      ΥΠΟΛΟΙΠΟ"
    Print #1, ""
    intProcessedDetailLines = 7
    
    Return
    
End Function

Private Function ValidateFields()

    'Αρχικές τιμές
    ValidateFields = False
    
    'Από
    If mskInvoiceDateIssueFrom.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskInvoiceDateIssueFrom.SetFocus
        Exit Function
    End If
    
    'Εως
    If mskInvoiceDateIssueTo.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        mskInvoiceDateIssueTo.SetFocus
        Exit Function
    End If
    
    'Σωστό διάστημα
    If IsDate(mskInvoiceDateIssueFrom.text) And IsDate(mskInvoiceDateIssueTo.text) Then
        If CDate(mskInvoiceDateIssueFrom.text) > CDate(mskInvoiceDateIssueTo.text) Then
            If MyMsgBox(4, strApplicationName, strStandardMessages(10), 1) Then
            End If
            mskInvoiceDateIssueFrom.SetFocus
            Exit Function
        End If
    End If

    'Εγγραφές
    If txtFilterID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtFilterDescription.SetFocus
        Exit Function
    End If

    ValidateFields = True
    
End Function

Private Function AbortProcedure(blnStatus)

    If blnProcessing Then blnProcessing = False: Exit Function

    If Not blnStatus Then
        ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
        ClearFields grdPersonsBalanceSheet
        frmCriteria(0).Visible = True
        mskInvoiceDateIssueFrom.SetFocus
        UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
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
    Dim rstPersons As Recordset
    Dim rstTransactions As Recordset

    'Αρχικές τιμές
    intIndex = 0
    lngRow = 0
    lngRowCount = 0
    frmCriteria(0).Visible = False
    
    curDebitTotalSoFar = 0
    curCreditTotalSoFar = 0
    curBalanceTotalSoFar = 0
    curDebitTotalPeriod = 0
    curCreditTotalPeriod = 0
    curBalanceTotal = 0
    
    'Πλέγμα
    With grdPersonsBalanceSheet
        .Clear
        .TabStop = False
        .ColHeaderText("BalanceSoFar") = "Υπόλοιπο " & Chr(13) & " έως " & CDate(mskInvoiceDateIssueFrom.text) - 1
        .ColHeaderText("Balance") = "Υπόλοιπο " & Chr(13) & " έως " & CDate(mskInvoiceDateIssueTo.text)
        .ColHeaderTextFlags(8) = 32789
        .Redraw = False
    End With
    
    'Ολοι οι πελάτες
    strSQL = "SELECT ID, Description FROM " & txtCustomersOrSuppliers.text & " ORDER BY Description "
    Set rstPersons = CommonDB.OpenRecordset(strSQL)
    
    Set TempQuery = CommonDB.CreateQueryDef("")

    'Αν δεν έχω εγγραφές, βγαίνω
    If rstPersons.RecordCount = 0 Then blnError = False: RefreshList = False: Exit Function
    
    'Προετοιμάζω τη μπάρα προόδου
    InitializeProgressBar Me, strApplicationName, rstPersons
    
    'Προσωρινά
    UpdateButtons Me, 6, 0, 0, 0, 0, 0, 1, 0
    cmdButton(5).Caption = "Διακοπή επεξεργασίας"
    blnProcessing = True
    
    'Επεξεργασία για κάθε πελάτη
    Do While Not rstPersons.EOF
        If Not blnProcessing Then Exit Do
        UpdateProgressBar Me
        'Κινήσεις
        strSQL = CreateSELECTStatement(txtInvoiceMasterRefersTo.text)
        'Καθαρίζω τις μεταβλητές για το loop
        intIndex = 0
        strParameters = ""
        strParFields = ""
        'Προηγούμενη περίοδος
        curDebitSoFar = 0
        curCreditSoFar = 0
        curBalanceSoFar = 0
        'Ζητούμενη περίοδος
        curDebitPeriod = 0
        curCreditPeriod = 0
        'Υπόλοιπα
        curBalance = 0
        'Πελάτης
        strThisParameter = "intPerson Integer"
        strThisQuery = "Invoices.InvoicePersonID = intPerson"
        strLogic = " AND "
        GoSub UpdateSQLString
        arrQuery(intIndex) = rstPersons!ID
        'Εως
        If IsDate(mskInvoiceDateIssueTo.text) Then
            strThisParameter = "datTo Date"
            strThisQuery = "Invoices.InvoiceDateIssue <= datTo "
            strLogic = " AND "
            GoSub UpdateSQLString
            arrQuery(intIndex) = CDate(mskInvoiceDateIssueTo.text)
        End If
        'Ταξινόμηση
        strOrder = " ORDER BY Invoices.InvoiceDateIssue "
        'Ελέγχω αν έχω κριτήρια
        strParameters = "PARAMETERS " & strParameters & "; "
        strParFields = " WHERE " & strParFields
        strSQL = strParameters & strSQL & strParFields
        TempQuery.SQL = strSQL & strOrder
        For intIndex = 1 To UBound(arrQuery)
            TempQuery.Parameters(intIndex - 1) = arrQuery(intIndex)
        Next intIndex
        'Ανοίγω το recordset
        Set rstTransactions = TempQuery.OpenRecordset()
        'Διαβάζω τις εγγραφές των κινήσεων - Ολες οι κινήσεις του πελάτη εως την ημερομηνία ΕΩΣ
        With rstTransactions
            'Αν έχω βρει εγγραφές
            If .EOF = False Then
                CalculateSoFarTotals rstTransactions 'Υπολογίζω έως το "Από" της ζητούμενης περιόδου
                Do While Not .EOF
                    If Not blnProcessing Then Exit Do 'Async!
                    CalculateAskedPeriod rstTransactions 'Υπολογίζω τη ζητούμενη περίοδο
                    rstTransactions.MoveNext 'Επόμενη εγγραφή
                Loop
            End If
            AddCurrentLineToGrid rstPersons, grdPersonsBalanceSheet.RowCount + 1 'Εμφανίζω την τρέχουσα γραμμή
        End With
        rstPersons.MoveNext
    Loop
    
    'Ακύρωση επεξεργασίας
    If Not blnProcessing Then
        blnProcessing = True
        ClearFields grdPersonsBalanceSheet
        RefreshList = 0
    Else
        RefreshList = lngRowCount
        blnProcessing = False
    End If
    
    'Σύνολα
    If Not blnProcessing Then
        AddGrandTotalsToGrid grdPersonsBalanceSheet.RowCount + 2
    End If
    
    'Τελικές ενέργειες
    cmdButton(5).Caption = "Νέα αναζήτηση"
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
    ClearFields grdPersonsBalanceSheet, frmProgress
    DisplayErrorMessage True, Err.Description
    
    Exit Function
    
End Function

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Εγγραφές - F2
            Set tmpRecordset = CheckForMatch("CommonDB", "Options", "OptionDescription", "String", txtFilterDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 0, 0)
                txtFilterID.text = tmpTableData.strCode
                txtFilterDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()
                
    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdPersonsBalanceSheet, 62, GetSetting(strApplicationName, "Layout Strings", "grdPersonsBalanceSheet"), _
            "05NCNID,40NLNDescription,10NRFXDebitSoFar,10NRFXCreditSoFar,10NRFXBalanceSoFar,10NRFXDebitPeriod,10NRFXCreditPeriod,10NRFBalance,04NCNSelected", _
            "ID, Επωνυμία,Χρέωση προηγούμενης περιόδου,Πίστωση προηγούμενης περιόδου,Υπόλοιπο προηγούμενης περιόδου,Χρέωση ζητούμενης περιόδου,Πίστωση ζητούμενης περιόδου,Υπόλοιπο,E"
        Me.Refresh
        frmCriteria(0).Visible = True
        mskInvoiceDateIssueFrom.SetFocus
    End If
            
    'AddDummyLines grdPersonsBalanceSheet, "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "-9999999", "-99999999", "-99999999", "-99999999", "-99999999", "-99999999"
    
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
        Case vbKeyX And CtrlDown = 5 And cmdButton(4).Enabled
            cmdButton_Click 4
        Case vbKeyEscape
            If cmdButton(5).Enabled Then cmdButton_Click 5: Exit Function
            If cmdButton(6).Enabled Then cmdButton_Click 6
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Sub Form_Load()

    SetUpGrid lstIconList, grdPersonsBalanceSheet
    PositionControls Me, True, grdPersonsBalanceSheet
    ColorizeControls Me, True
    ClearFields lblSelectedGridTotals, lblSelectedGridLines, lblCriteria, lblRecordCount
    ClearFields txtFilterID
    ClearFields mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, txtFilterDescription
    ClearFields grdPersonsBalanceSheet
    EnableFields mskInvoiceDateIssueFrom, mskInvoiceDateIssueTo, txtFilterDescription
    UpdateButtons Me, 6, 1, 0, 0, 0, 0, 0, 1
    
End Sub

Private Sub grdPersonsBalanceSheet_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    bDoDefault = False

End Sub

Private Sub grdPersonsBalanceSheet_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdPersonsBalanceSheet_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeySpace And grdPersonsBalanceSheet.RowCount > 0 Then
        grdPersonsBalanceSheet.CellIcon(grdPersonsBalanceSheet.CurRow, "Selected") = lstIconList.ItemIndex(SelectRow(grdPersonsBalanceSheet, KeyCode, grdPersonsBalanceSheet.CurRow, "ID"))
        lblSelectedGridLines.Caption = CountSelected(grdPersonsBalanceSheet)
        lblSelectedGridTotals.Caption = SumSelectedGridRows(grdPersonsBalanceSheet, False, "BalanceSoFar", "DebitPeriod", "CreditPeriod", "Balance")
    End If

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdPersonsBalanceSheet", grdPersonsBalanceSheet.LayoutCol

End Sub

Private Sub txtFilterDescription_Change()

    If txtFilterDescription.text = "" Then
        ClearFields txtFilterID
    End If

End Sub

Private Sub txtFilterDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtFilterDescription_Validate(Cancel As Boolean)

    If txtFilterID = "" And txtFilterDescription.text <> "" Then cmdIndex_Click 0

End Sub

Private Function CalculateAskedPeriod(rstTransactions As Recordset)

    If txtInvoiceMasterRefersTo.text = "1" Then CalculateAskedPeriodForSuppliers rstTransactions
    If txtInvoiceMasterRefersTo.text = "2" Then CalculateAskedPeriodForCustomers rstTransactions

End Function


