VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesCodes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   11430
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   15870
   ControlBox      =   0   'False
   ForeColor       =   &H00400000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11430
   ScaleWidth      =   15870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   450
      TabIndex        =   34
      Top             =   8250
      Width           =   7515
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
         Index           =   4
         Left            =   5925
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
         Index           =   3
         Left            =   4500
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
   End
   Begin UserControls.newText txtCodeLastNo 
      Height          =   465
      Left            =   2175
      TabIndex        =   8
      Top             =   6375
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   820
      Alignment       =   2
      MaxLength       =   6
      Text            =   "99999"
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
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2190
      Left            =   7350
      TabIndex        =   18
      Top             =   5475
      Width           =   4515
      Begin VB.TextBox txtCodeSecondaryRefersTo 
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Codes.CodeSecondaryRefersTo"
         Top             =   450
         Width           =   3540
      End
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
         Text            =   "Codes.CodeMasterRefersTo"
         Top             =   75
         Width           =   3540
      End
      Begin VB.TextBox txtCodeMasterRefersTo 
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
         Top             =   75
         Width           =   780
      End
      Begin VB.TextBox txtCodeHandID 
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
         Top             =   1200
         Width           =   780
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
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Codes.CodeHandID"
         Top             =   1200
         Width           =   3540
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
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Codes.CodeID"
         Top             =   825
         Width           =   3540
      End
      Begin VB.TextBox txtCodeID 
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   1575
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "TablesCodes.frx":0000
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin UserControls.newDate mskCodeLastDate 
      Height          =   465
      Left            =   2175
      TabIndex        =   9
      Top             =   6900
      Width           =   1455
      _ExtentX        =   2672
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
   Begin iGrid300_10Tec.iGrid grdCodes 
      Height          =   6615
      Left            =   7200
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1125
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   11668
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
   Begin UserControls.newText txtCodeBatch 
      Height          =   465
      Left            =   1800
      TabIndex        =   4
      Top             =   2700
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   1
      Text            =   "Α"
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
   Begin UserControls.newText txtCodeShortDescriptionB 
      Height          =   465
      Left            =   1800
      TabIndex        =   2
      Top             =   1650
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin UserControls.newText txtCodeShortDescriptionA 
      Height          =   465
      Left            =   1800
      TabIndex        =   1
      Top             =   1125
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin UserControls.newText txtCodeDescription 
      Height          =   465
      Left            =   1800
      TabIndex        =   3
      Top             =   2175
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
      BackColor       =   &H00C0FFFF&
      Caption         =   " Τελευταίο παραστατικό που εκδόθηκε "
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
      Height          =   1890
      Index           =   1
      Left            =   450
      TabIndex        =   11
      Tag             =   "SameColorAsBackground"
      Top             =   5850
      Width           =   3615
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   9
         Left            =   3150
         Top             =   975
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
         Left            =   1275
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
         Index           =   5
         Left            =   0
         Top             =   450
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
         Left            =   1800
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
         Left            =   2025
         Top             =   1500
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Νο"
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
         TabIndex        =   17
         Top             =   600
         Width           =   840
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
         Index           =   10
         Left            =   450
         TabIndex        =   16
         Top             =   1125
         Width           =   840
      End
   End
   Begin UserControls.newText txtCodeHandDescription 
      Height          =   465
      Left            =   1800
      TabIndex        =   5
      Top             =   3225
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
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
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   0
      Left            =   2475
      TabIndex        =   24
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
      PicNormal       =   "TablesCodes.frx":1210
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin UserControls.newText txtCodeCustomers 
      Height          =   465
      Left            =   1050
      TabIndex        =   6
      Top             =   4875
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   1
      Text            =   "+"
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
   Begin UserControls.newText txtCodeSuppliers 
      Height          =   465
      Left            =   2025
      TabIndex        =   7
      Top             =   4875
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   1
      Text            =   "+"
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
      BackColor       =   &H00C0FFFF&
      Caption         =   " Ενημερώνει"
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
      Height          =   2040
      Index           =   3
      Left            =   450
      TabIndex        =   25
      Tag             =   "SameColorAsBackground"
      Top             =   3750
      Width           =   3615
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   12
         Left            =   1575
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Προμηθευτές"
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
         Index           =   16
         Left            =   1425
         TabIndex        =   29
         Top             =   525
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Πελάτες"
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
         Index           =   17
         Left            =   450
         TabIndex        =   28
         Top             =   525
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "(+/-)"
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
         Index           =   20
         Left            =   1425
         TabIndex        =   27
         Top             =   825
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "(+/-)"
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
         Index           =   21
         Left            =   450
         TabIndex        =   26
         Top             =   825
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   11
         Left            =   0
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   10
         Left            =   2325
         Top             =   675
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   13
      Left            =   0
      Top             =   8175
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   6900
      Top             =   7725
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   13050
      Top             =   3675
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   4575
      Top             =   8925
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   2
      Left            =   6750
      Top             =   1950
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
      Left            =   1350
      Top             =   1650
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
      Top             =   1650
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
      Left            =   7650
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Χειρόγραφο"
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
      TabIndex        =   21
      Top             =   3300
      Width           =   915
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Τύποι παραστατικών"
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
      TabIndex        =   15
      Top             =   75
      Width           =   4815
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Συντ. Β'"
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
      TabIndex        =   13
      Top             =   1725
      Width           =   915
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Συντ. A'"
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
      TabIndex        =   0
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Περιγραφή"
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
      Top             =   2250
      Width           =   915
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Σειρά"
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
      Top             =   2775
      Width           =   915
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
      Index           =   4
      Left            =   1800
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Menu mnuHdrPopUp 
      Caption         =   "mnuHdrPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuΑποθήκευσηΠλάτουςΣτηλών 
         Caption         =   "Αποθήκευση πλάτους στηλών"
      End
   End
End
Attribute VB_Name = "TablesCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim lngSelectedRow As Long
    
Private Function AbortProcedure(blnStatus)
    
    If Not blnStatus Then
        If MyMsgBox(3, strAppTitle, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtCodeID, txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeHandID, txtCodeHandDescription, txtCodeCustomers, txtCodeSuppliers, txtCodeLastNo, mskCodeLastDate
            DisableFields txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeCustomers, txtCodeSuppliers, txtCodeHandDescription, txtCodeLastNo, mskCodeLastDate, cmdIndex(0)
            grdCodes.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", "Codes", strAppTitle, "CodeID", Val(txtCodeID.text), "True") Then
        PopulateGrid
        HighlightRow grdCodes, lngSelectedRow, 2, "", True
        ClearFields txtCodeID, txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeHandID, txtCodeHandDescription, txtCodeCustomers, txtCodeSuppliers, txtCodeLastNo, mskCodeLastDate
        DisableFields txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeCustomers, txtCodeSuppliers, txtCodeHandDescription, txtCodeLastNo, mskCodeLastDate, cmdIndex(0)
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    End If

End Function

Private Function PopulateGrid()
        
    If FillGridFromDB("CommonDB", grdCodes, "Codes", "", "", "CodeMasterRefersTo = '" & txtCodeMasterRefersTo.text & "' AND CodeSecondaryRefersTo = '" & txtCodeSecondaryRefersTo.text & "'", 3, 0, 3, 5, 6) Then
        grdCodes.SetFocus
        grdCodes.SetCurCell 1, 1
    End If

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtCodeID, txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeHandID, txtCodeHandDescription, txtCodeCustomers, txtCodeSuppliers, txtCodeLastNo, mskCodeLastDate
    EnableFields txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeCustomers, txtCodeSuppliers, txtCodeHandDescription, txtCodeLastNo, mskCodeLastDate, cmdIndex(0)
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    txtCodeShortDescriptionA.SetFocus

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "Codes", blnStatus, strAppTitle, "CodeID", Val(txtCodeID.text), txtCodeMasterRefersTo.text, txtCodeSecondaryRefersTo.text, txtCodeShortDescriptionA.text, txtCodeShortDescriptionB.text, txtCodeDescription.text, txtCodeBatch.text, Val(txtCodeHandID.text), txtCodeCustomers.text, txtCodeSuppliers.text, txtCodeLastNo.text, mskCodeLastDate.text, 1, strCurrentUser) <> 0 Then
        PopulateGrid
        HighlightRow grdCodes, lngSelectedRow, 3, txtCodeDescription.text, True
        lngSelectedRow = 0
        ClearFields txtCodeID, txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeHandID, txtCodeHandDescription, txtCodeCustomers, txtCodeSuppliers, txtCodeLastNo, mskCodeLastDate, 1
        DisableFields txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeCustomers, txtCodeSuppliers, txtCodeHandDescription, txtCodeLastNo, mskCodeLastDate, cmdIndex(0)
        UpdateButtons Me, 4, 1, 0, 0, 0, 1
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If

End Function

Private Function SeekRecord()
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Dim blnEnableDelete As Boolean
    
    If grdCodes.RowCount = 0 Then Exit Function
    
    ClearFields txtCodeID, txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeHandID, txtCodeHandDescription, txtCodeCustomers, txtCodeSuppliers, txtCodeLastNo, mskCodeLastDate
    DisableFields txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeCustomers, txtCodeSuppliers, txtCodeHandDescription, txtCodeLastNo, mskCodeLastDate, cmdIndex(0)
    
    blnEnableDelete = SimpleSeek("Invoices", "InvoiceCodeID", grdCodes.CellValue(grdCodes.CurRow, 1))
    
    If MainSeekRecord("CommonDB", "Codes", "CodeID", grdCodes.CellValue(grdCodes.CurRow, 1), True, txtCodeID, txtCodeMasterRefersTo, txtCodeSecondaryRefersTo, txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeHandID, txtCodeCustomers, txtCodeSuppliers, txtCodeLastNo, mskCodeLastDate) Then
        'Χειρόγραφο
        Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtCodeHandID.text)
        txtCodeHandID.text = tmpRecordset.fields(0)
        txtCodeHandDescription.text = tmpRecordset.fields(1)
        '
        blnStatus = False
        lngSelectedRow = grdCodes.CurRow
        EnableFields txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeCustomers, txtCodeSuppliers, txtCodeHandDescription, txtCodeLastNo, mskCodeLastDate, cmdIndex(0)
        UpdateButtons Me, 4, 0, 1, IIf(blnEnableDelete, 1, 0), 1, 0
        txtCodeShortDescriptionA.SetFocus
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'ID
    If Len(txtCodeShortDescriptionA.text) = 0 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtCodeShortDescriptionA.SetFocus
        Exit Function
    End If
    
    'Συντ. Α'
    If Len(txtCodeShortDescriptionA.text) = 0 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtCodeShortDescriptionA.SetFocus
        Exit Function
    End If
    
    'Συντ. Β'
    If Len(txtCodeShortDescriptionB.text) = 0 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtCodeShortDescriptionB.SetFocus
        Exit Function
    End If
    If Len(txtCodeShortDescriptionB.text) <> 3 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        txtCodeShortDescriptionB.SetFocus
        Exit Function
    End If
        
    'Περιγραφή
    If Len(txtCodeDescription.text) = 0 Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtCodeDescription.SetFocus
        Exit Function
    End If
    
    'Πελάτες
    If txtCodeCustomers.text <> "+" And txtCodeCustomers.text <> "-" And txtCodeCustomers.text <> "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        txtCodeCustomers.SetFocus
        Exit Function
    End If
    
    'Προμηθευτές
    If txtCodeSuppliers.text <> "+" And txtCodeSuppliers.text <> "-" And txtCodeSuppliers.text <> "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        txtCodeSuppliers.SetFocus
        Exit Function
    End If
    
    'Χειρόγραφο
    If txtCodeHandID.text = "" Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(1), 1) Then
        End If
        txtCodeHandDescription.SetFocus
        Exit Function
    End If
    
    'Τελευταίο παραστατικό
    If Not IsNumeric(txtCodeLastNo.text) Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        txtCodeLastNo.SetFocus
        Exit Function
    End If
    
    'Τελευταία ημερομηνία
    If Not IsDate(mskCodeLastDate.text) Then
        If MyMsgBox(4, strAppTitle, strStandardMessages(2), 1) Then
        End If
        mskCodeLastDate.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

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
            AbortProcedure False
        Case 4
            AbortProcedure True
    End Select

End Sub

Private Sub cmdIndex_Click(Index As Integer)
    
    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case Index
        Case 0
            'Χειρόγραφο
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtCodeHandDescription.text)
            tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "CodeID", "Περιγραφή", 0, 40, 1, 0)
            txtCodeHandID.text = tmpTableData.strCode
            txtCodeHandDescription.text = tmpTableData.strFirstField
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdCodes, 25, GetSetting(strAppTitle, "Layout Strings", "grdCodes"), "04NCIID,04NCNShortDescription,40NLNDescription,05NCNBatch", "ID,Συντ. Α,Περιγραφή,Σειρά"
        Me.Refresh
        PopulateGrid
    End If

    'AddDummyLines grdCodes, "99999", "AAAAA", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", "A"

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
    SetUpGrid lstIconList, grdCodes
    ClearFields txtCodeID, txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeHandID, txtCodeHandDescription, txtCodeCustomers, txtCodeSuppliers, txtCodeLastNo, mskCodeLastDate
    DisableFields txtCodeShortDescriptionA, txtCodeShortDescriptionB, txtCodeDescription, txtCodeBatch, txtCodeCustomers, txtCodeSuppliers, txtCodeHandDescription, txtCodeLastNo, mskCodeLastDate, cmdIndex(0)
    UpdateButtons Me, 4, 1, 0, 0, 0, 1
    
End Sub

Private Sub grdCodes_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub

Private Sub grdCodes_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdCodes_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord

End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strAppTitle, "Layout Strings", "grdCodes", grdCodes.LayoutCol

End Sub

Private Sub txtCodeHandDescription_Change()

    If txtCodeHandDescription.text = "" Then
        ClearFields txtCodeHandID
    End If

End Sub


Private Sub txtCodeHandDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    
End Sub


Private Sub txtCodeHandDescription_Validate(Cancel As Boolean)

    If txtCodeHandID.text = "" And txtCodeHandDescription.text <> "" Then cmdIndex_Click 0
    
End Sub


