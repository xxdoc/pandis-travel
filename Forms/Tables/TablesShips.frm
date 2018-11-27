VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "ImageList.ocx"
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesShips 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   11430
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   17760
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "TablesShips.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11430
   ScaleWidth      =   17760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFrame 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   6465
      Index           =   0
      Left            =   15975
      TabIndex        =   20
      Top             =   4500
      Width           =   12540
      Begin UserControls.newText txtShipDescription 
         Height          =   465
         Left            =   2025
         TabIndex        =   0
         Top             =   525
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
      Begin UserControls.newText txtShipFlag 
         Height          =   465
         Left            =   2025
         TabIndex        =   2
         Top             =   1575
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
      Begin UserControls.newText txtShipRegistryNo 
         Height          =   465
         Left            =   2025
         TabIndex        =   3
         Top             =   2100
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
      Begin UserControls.newText txtShipIMO 
         Height          =   465
         Left            =   2025
         TabIndex        =   4
         Top             =   2625
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
      Begin UserControls.newInteger mskShipPersons 
         Height          =   465
         Left            =   2025
         TabIndex        =   1
         Top             =   1050
         Width           =   540
         _ExtentX        =   953
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
      Begin iGrid300_10Tec.iGrid grdShips 
         Height          =   5415
         Left            =   7425
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   525
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   9551
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
      Begin UserControls.newText txtShipSalesCode 
         Height          =   465
         Left            =   2025
         TabIndex        =   5
         Top             =   3150
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   820
         Alignment       =   2
         ForeColor       =   0
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
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Κωδ. πωλήσεων"
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
         TabIndex        =   61
         Top             =   3225
         Width           =   1140
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   15
         Left            =   8175
         Top             =   5925
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
         Left            =   3450
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
         Index           =   5
         Left            =   12075
         Top             =   1725
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
         Left            =   6975
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   1140
         Index           =   2
         Left            =   1575
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
         Index           =   1
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Ονομασία"
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
         TabIndex        =   26
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Χωρητικότητα"
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
         TabIndex        =   25
         Top             =   1125
         Width           =   1140
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Σημαία"
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
         TabIndex        =   24
         Top             =   1650
         Width           =   1140
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Αρ. Νηολογίου"
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
         TabIndex        =   23
         Top             =   2175
         Width           =   1140
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "I.M.O."
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
         TabIndex        =   22
         Top             =   2700
         Width           =   1140
      End
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   62
      Top             =   8025
      Width           =   7515
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
   Begin VB.Frame frmFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6465
      Index           =   2
      Left            =   16125
      TabIndex        =   34
      Top             =   675
      Width           =   12465
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Υπεύθυνοι καταγραφής επιβατών "
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
         Height          =   3765
         Index           =   3
         Left            =   450
         TabIndex        =   35
         Top             =   2175
         Width           =   11565
         Begin UserControls.newText txtShipManagerAPersonName 
            Height          =   465
            Left            =   2100
            TabIndex        =   9
            Top             =   825
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerBPersonName 
            Height          =   465
            Left            =   5925
            TabIndex        =   14
            Top             =   825
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerBPersonPhones 
            Height          =   465
            Left            =   5925
            TabIndex        =   15
            Top             =   1350
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerAPersonPhones 
            Height          =   465
            Left            =   2100
            TabIndex        =   10
            Top             =   1350
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerAPersonEmail 
            Height          =   465
            Left            =   2100
            TabIndex        =   11
            Top             =   1875
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerBPersonEmail 
            Height          =   465
            Left            =   5925
            TabIndex        =   16
            Top             =   1875
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerAPersonFax 
            Height          =   465
            Left            =   2100
            TabIndex        =   12
            Top             =   2400
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerBPersonFax 
            Height          =   465
            Left            =   5925
            TabIndex        =   17
            Top             =   2400
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerAPersonAddress 
            Height          =   465
            Left            =   2100
            TabIndex        =   13
            Top             =   2925
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerBPersonAddress 
            Height          =   465
            Left            =   5925
            TabIndex        =   18
            Top             =   2925
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Ονοματεπώνυμο"
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
            Left            =   450
            TabIndex        =   42
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Τηλέφωνα"
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
            Index           =   21
            Left            =   450
            TabIndex        =   41
            Top             =   1425
            Width           =   1215
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Email"
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
            Index           =   22
            Left            =   450
            TabIndex        =   40
            Top             =   1950
            Width           =   1215
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Fax"
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
            Index           =   23
            Left            =   450
            TabIndex        =   39
            Top             =   2475
            Width           =   1215
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Διεύθυνση"
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
            Index           =   24
            Left            =   450
            TabIndex        =   38
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Κύριος υπεύθυνος"
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
            Index           =   25
            Left            =   2100
            TabIndex        =   37
            Top             =   525
            Width           =   3765
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Αντικαταστάτης κύριου υπεύθυνου"
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
            Index           =   26
            Left            =   5925
            TabIndex        =   36
            Top             =   525
            Width           =   3765
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   24
            Left            =   0
            Top             =   1275
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   25
            Left            =   3375
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
            Index           =   26
            Left            =   3150
            Top             =   3375
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   27
            Left            =   1650
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
            Index           =   28
            Left            =   9675
            Top             =   1275
            Visible         =   0   'False
            Width           =   465
         End
      End
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Στοιχεία εταιρίας για καταγραφή επιβατών "
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
         Height          =   1665
         Index           =   2
         Left            =   450
         TabIndex        =   43
         Top             =   450
         Width           =   11565
         Begin UserControls.newText txtShipManagerName 
            Height          =   465
            Left            =   450
            TabIndex        =   6
            Top             =   825
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerAgent 
            Height          =   465
            Left            =   6675
            TabIndex        =   8
            Top             =   825
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin UserControls.newText txtShipManagerNameInGreece 
            Height          =   465
            Left            =   3525
            TabIndex        =   7
            Top             =   825
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   30
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
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
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Υπεύθυνος διαχειριστής"
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
            TabIndex        =   46
            Top             =   525
            Width           =   1740
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Διαχειριστής στην Ελλάδα"
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
            Index           =   18
            Left            =   3525
            TabIndex        =   45
            Top             =   525
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Υπεύθυνοι ναυτικοί πράκτορες"
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
            Index           =   19
            Left            =   6675
            TabIndex        =   44
            Top             =   525
            Width           =   2190
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   20
            Left            =   0
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
            Index           =   21
            Left            =   3750
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
            Index           =   22
            Left            =   1800
            Top             =   1275
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   23
            Left            =   9675
            Top             =   825
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
         Index           =   16
         Left            =   12000
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   19
         Left            =   0
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
         Index           =   29
         Left            =   12000
         Top             =   3975
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2190
      Left            =   9525
      TabIndex        =   31
      Top             =   8475
      Width           =   4515
      Begin VB.TextBox txtShipSaveAndNewID 
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
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
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Ships.ShipSaveAndNewID"
         Top             =   1200
         Width           =   3540
      End
      Begin VB.TextBox txtShipRepeatedEntriesID 
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
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   825
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
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "Ships.ShipRepeatedEntriesID"
         Top             =   825
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
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "ShipsManagers.ManagerID"
         Top             =   450
         Width           =   3540
      End
      Begin VB.TextBox txtShipManagerID 
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
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox txtShipID 
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
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Ships.ShipID"
         Top             =   75
         Width           =   3540
      End
      Begin vbalIml6.vbalImageList lstIconList 
         Left            =   75
         Top             =   1575
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   4592
         Images          =   "TablesShips.frx":000C
         Version         =   131072
         KeyCount        =   4
         Keys            =   "ÿÿÿ"
      End
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   6465
      Index           =   3
      Left            =   15150
      TabIndex        =   53
      Top             =   2475
      Width           =   12465
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Καταχώρηση επιβατών "
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
         TabIndex        =   54
         Top             =   450
         Width           =   11565
         Begin UserControls.newText txtShipRepeatedEntriesDescription 
            Height          =   465
            Left            =   3975
            TabIndex        =   55
            Top             =   525
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   0
            Left            =   4650
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   525
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
            PicNormal       =   "TablesShips.frx":121C
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin UserControls.newText txtShipSaveAndNewDescription 
            Height          =   465
            Left            =   3975
            TabIndex        =   57
            Top             =   1050
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
         Begin Dacara_dcButton.dcButton cmdIndex 
            Height          =   465
            Index           =   1
            Left            =   4650
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1050
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
            PicNormal       =   "TablesShips.frx":17B6
            PicSizeH        =   16
            PicSizeW        =   16
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   6
            Left            =   0
            Top             =   600
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
            Left            =   4200
            Top             =   0
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Επαναλαμβανόμενη καταχώρηση"
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
            TabIndex        =   60
            Top             =   600
            Width           =   3090
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   8
            Left            =   3525
            Top             =   675
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "Αποθήκευση και δημιουργία με μία εντολή"
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
            TabIndex        =   59
            Top             =   1125
            Width           =   3090
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   390
            Index           =   9
            Left            =   4125
            Top             =   1500
            Visible         =   0   'False
            Width           =   465
         End
      End
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   6465
      Index           =   1
      Left            =   1875
      TabIndex        =   21
      Top             =   1125
      Width           =   12465
      Begin iGrid300_10Tec.iGrid grdCrew 
         Height          =   5190
         Left            =   450
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   525
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   9155
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
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   465
         Index           =   5
         Left            =   3525
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   5850
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   820
         BackColor       =   8421376
         ButtonShape     =   3
         ButtonStyle     =   2
         Caption         =   "Δημιουργία μέλους πληρώματος"
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
         Height          =   465
         Index           =   6
         Left            =   6300
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   5850
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   820
         BackColor       =   8421376
         ButtonShape     =   3
         ButtonStyle     =   2
         Caption         =   "Διαγραφή μέλους πληρώματος"
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
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   1140
         Index           =   13
         Left            =   12000
         Top             =   1950
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   1140
         Index           =   12
         Left            =   0
         Top             =   2175
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   11
         Left            =   3750
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   0
      Left            =   450
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Γενικά στοιχεία πλοίου"
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
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   1
      Left            =   450
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Στοιχεία πληρώματος"
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
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   2
      Left            =   450
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Υπεύθυνοι καταγραφής επιβατών"
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
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   3
      Left            =   450
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Αρχείο καταχώρησης επιβατών"
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
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   3
      Left            =   1725
      Top             =   4275
      Width           =   1890
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   2
      Left            =   1725
      Top             =   3225
      Width           =   1890
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   14325
      Top             =   2025
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   4575
      Top             =   8700
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   7350
      Top             =   7500
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   1140
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
      Height          =   1140
      Index           =   4
      Left            =   5700
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   0
      Left            =   1725
      Top             =   1125
      Width           =   1890
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   1
      Left            =   1725
      Top             =   2175
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Πλοία"
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   30
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   720
      Left            =   225
      TabIndex        =   19
      Top             =   75
      Width           =   1380
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
Attribute VB_Name = "TablesShips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim blnStatus As Boolean
Dim lngSelectedRow As Long

Private Function AddGridLine()

    With grdCrew
        .Enabled = True
        .AddRow
        .CellIcon(.RowCount, "Status") = lstIconList.ItemIndex(2)
        .SetCurCell .RowCount, 2
        .SetFocus
    End With
    
End Function

Private Function DeleteCrew(shipID As Integer)

    On Error GoTo ErrTrap
    
    Dim lngID As Long
    Dim lngRow As Long
    
    DeleteCrew = True
    
    With grdCrew
        For lngRow = 1 To .RowCount
            If Not MainDeleteRecord("CommonDB", "ShipsCrew", strApplicationName, "ID", .CellValue(lngRow, "ID"), False) Then
                DeleteCrew = False
                Exit For
            End If
        Next lngRow
    End With
    
    Exit Function

ErrTrap:
    DeleteCrew = False
    DisplayErrorMessage True, Err.Description
    
    Exit Function

End Function

Private Function PopulateCrewGrid()

    If FillGridFromDB("CommonDB", grdCrew, "ShipsCrew", "", "", "CrewShipID = " & txtShipID.text, 2, 0, 1, 2, 3, 4, 5) Then
        PopulateCrewGrid = True
    Else
        PopulateCrewGrid = False
    End If
    
    'AddDummyLines grdCrew, "99999", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "ΑΑΑΑΑΑΑΑΑΑ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ", "ΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑΑ"

End Function

Private Function ValidateFields()

    Dim lngRow As Long
    Dim lngCol As Long
    
    ValidateFields = False
    
    'Ονομα
    If Len(txtShipDescription.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 0
        txtShipDescription.SetFocus
        Exit Function
    End If
    
    'Χωρητικότητα
    If Len(mskShipPersons.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 0
        mskShipPersons.SetFocus
        Exit Function
    End If
    
    'Επαναλαμβανόμενη καταχώρηση
    If txtShipRepeatedEntriesID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 3
        txtShipRepeatedEntriesDescription.SetFocus
        Exit Function
    End If
                
    'Αποθήκευση και δημιουργία με ένα κλικ
    If txtShipSaveAndNewID.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 3
        txtShipSaveAndNewDescription.SetFocus
        Exit Function
    End If
    
    'Πλήρωμα
    For lngRow = 1 To grdCrew.RowCount
        'Αν δεν έχω επιλέξει τη γραμμή για διαγραφή
        If grdCrew.CellIcon(lngRow, "Deleted") <> 2 Then
            'Επίθετο
            If Len(grdCrew.CellValue(lngRow, "LastName")) = 0 Then
                If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
                End If
                btnPanel_Click 1
                grdCrew.SetCurCell lngRow, "LastName"
                grdCrew.SetFocus
                Exit Function
            End If
            'Ιδιότητα
            If IsNull(grdCrew.CellValue(lngRow, "PropertyDescription")) Or IsEmpty(grdCrew.CellValue(lngRow, "PropertyDescription")) Then
                If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
                End If
                btnPanel_Click 1
                grdCrew.SetCurCell lngRow, "PropertyDescription"
                grdCrew.SetFocus
                Exit Function
            End If
            'Φύλο
            If IsNull(grdCrew.CellValue(lngRow, "GenderDescription")) Or IsEmpty(grdCrew.CellValue(lngRow, "GenderDescription")) Then
                If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
                End If
                btnPanel_Click 1
                grdCrew.SetCurCell lngRow, "GenderDescription"
                grdCrew.SetFocus
                Exit Function
            End If
            'Ηλικία
            If IsNull(grdCrew.CellValue(lngRow, "AgeDescription")) Or IsEmpty(grdCrew.CellValue(lngRow, "AgeDescription")) Then
                If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
                End If
                btnPanel_Click 1
                grdCrew.SetCurCell lngRow, "AgeDescription"
                grdCrew.SetFocus
                Exit Function
            End If
        End If
    Next lngRow
    
    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)
    
    If grdCrew.TextEditText <> "" Then
        grdCrew.CancelEdit
        Exit Function
    End If

    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            blnStatus = False
            ClearFields txtShipID, txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerID, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipRepeatedEntriesID, txtShipSaveAndNewDescription, txtShipSaveAndNewID
            DisableFields txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipSaveAndNewDescription, btnPanel(1), btnPanel(2), btnPanel(3)
            btnPanel_Click 0
            grdShips.SetFocus
            UpdateButtons Me, 4, 1, 0, 0, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    With wrkCurrent
        
        .BeginTrans
    
        If MainDeleteRecord("CommonDB", "Ships", strApplicationName, "ID", txtShipID.text, "True") Then
            If MainDeleteRecord("CommonDB", "ShipsManagers", strApplicationName, "ManagerID", txtShipManagerID.text, False) Then
                If DeleteCrew(txtShipID.text) Then
                    btnPanel_Click 0
                    PopulateGrid
                    HighlightRow grdShips, lngSelectedRow, 1, "", True
                    ClearFields txtShipID, txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerID, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipRepeatedEntriesID, txtShipSaveAndNewDescription, txtShipSaveAndNewID
                    DisableFields txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipSaveAndNewDescription, btnPanel(1), btnPanel(2), btnPanel(3)
                    UpdateButtons Me, 4, 1, 0, 0, 0, 1
                    .CommitTrans
                    Exit Function
                End If
            End If
        End If
    
        .Rollback
        
    End With

End Function

Private Function PopulateGrid()

    If FillGridFromDB("CommonDB", grdShips, "Ships", "", "", "", 2, 0, 1) Then
        grdShips.SetFocus
        grdShips.SetCurCell 1, 1
    End If

End Function

Private Function NewRecord()
    
    blnStatus = True
    ClearFields txtShipID, txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerID, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipRepeatedEntriesID, txtShipSaveAndNewDescription, txtShipSaveAndNewID
    EnableFields txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipSaveAndNewDescription, btnPanel(1), btnPanel(2), btnPanel(3)
    UpdateButtons Me, 4, 0, 1, 0, 1, 0
    txtShipDescription.SetFocus

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    With wrkCurrent
        
        .BeginTrans
        
        txtShipID.text = MainSaveRecord("CommonDB", "Ships", blnStatus, strApplicationName, "ShipID", txtShipID.text, txtShipDescription.text, mskShipPersons.text, txtShipFlag.text, txtShipRegistryNo.text, txtShipIMO.text, txtShipSalesCode.text, txtShipRepeatedEntriesID.text, txtShipSaveAndNewID.text, 1, strCurrentUser)
        
        If txtShipID.text <> "0" Then
            txtShipManagerID.text = MainSaveRecord("CommonDB", "ShipsManagers", blnStatus, strApplicationName, "ShipManagerID", txtShipManagerID.text, txtShipID.text, txtShipManagerName.text, txtShipManagerNameInGreece.text, txtShipManagerAgent.text, txtShipManagerAPersonName.text, txtShipManagerAPersonPhones.text, txtShipManagerAPersonEmail.text, txtShipManagerAPersonFax.text, txtShipManagerAPersonAddress.text, txtShipManagerBPersonName.text, txtShipManagerBPersonPhones.text, txtShipManagerBPersonEmail.text, txtShipManagerBPersonFax.text, txtShipManagerBPersonAddress.text)
            If txtShipManagerID.text <> "0" Then
                If SaveCrew(txtShipID.text) Then
                    btnPanel_Click 0
                    PopulateGrid
                    HighlightRow grdShips, lngSelectedRow, 2, txtShipDescription.text, True
                    lngSelectedRow = 0
                    ClearFields txtShipID, txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerID, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipRepeatedEntriesID, txtShipSaveAndNewDescription, txtShipSaveAndNewID
                    DisableFields txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipSaveAndNewDescription, btnPanel(1), btnPanel(2), btnPanel(3)
                    UpdateButtons Me, 4, 1, 0, 0, 0, 1
                    .CommitTrans
                    Exit Function
                Else
                    DisplayErrorMessage True, strStandardMessages(5)
                End If
            Else
                DisplayErrorMessage True, strStandardMessages(5)
            End If
        Else
            DisplayErrorMessage True, strStandardMessages(5)
        End If
    
        .Rollback
        
    End With
    
End Function

Private Function SeekRecord()

    Dim blnEnableDelete As Boolean
    Dim tmpRecordset As Recordset
    
    If grdShips.RowCount = 0 Then Exit Function
    
    ClearFields txtShipID, txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerID, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipRepeatedEntriesID, txtShipSaveAndNewDescription, txtShipSaveAndNewID
    DisableFields txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipSaveAndNewDescription, btnPanel(1), btnPanel(2), btnPanel(3)
    blnEnableDelete = SimpleSeek("InvoicesOut", "InvoiceShipID", grdShips.CellValue(grdShips.CurRow, 1))
    If blnEnableDelete Then blnEnableDelete = SimpleSeek("Manifest", "TripShipID", grdShips.CellValue(grdShips.CurRow, 1))
    If MainSeekRecord("CommonDB", "Ships", "ShipID", grdShips.CellValue(grdShips.CurRow, 1), True, txtShipID, txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, txtShipRepeatedEntriesID, txtShipSaveAndNewID) Then
        If MainSeekRecord("CommonDB", "ShipsManagers", "ShipManagerShipID", txtShipID.text, True, txtShipManagerID, txtShipID, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress) Then
            If Not PopulateCrewGrid Then
                ClearFields grdCrew
                grdShips.SetFocus
            End If
            blnStatus = False
            lngSelectedRow = grdShips.CurRow
            EnableFields txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipSaveAndNewDescription, btnPanel(1), btnPanel(2), btnPanel(3)
            'Επαναλαμβανόμενη καταχώρηση (Εύρεση τελευταίας εγγραφής)
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtShipRepeatedEntriesID.text)
            txtShipRepeatedEntriesID.text = tmpRecordset.Fields(0)
            txtShipRepeatedEntriesDescription.text = tmpRecordset.Fields(1)
            'Αποθήκευση και δημιουργία με ένα κλικ
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtShipSaveAndNewID.text)
            txtShipSaveAndNewID.text = tmpRecordset.Fields(0)
            txtShipSaveAndNewDescription.text = tmpRecordset.Fields(1)
            '
            UpdateButtons Me, 4, 0, 1, IIf(blnEnableDelete, 1, 0), 1, 0, 0
            txtShipDescription.SetFocus
        End If
    End If
    
End Function

Private Sub btnPanel_Click(index As Integer)

    Dim intLoop As Integer
    
    For intLoop = 0 To 3
        btnPanel(intLoop).Enabled = True
        frmFrame(intLoop).Visible = False
        shpBridge(intLoop).Visible = False
    Next intLoop
    
    btnPanel(index).Enabled = False
    frmFrame(index).Visible = True
    shpBridge(index).Visible = True
    
    Select Case index
        'Στοιχεία πλοίου
        Case 0
            If txtShipDescription.Enabled Then txtShipDescription.SetFocus
        'Στοιχεία πληρώματος
        Case 1
            With grdCrew
                If .Enabled And .RowCount >= 1 Then
                    .SetFocus
                    .SetCurCell 1, 2
                End If
            End With
        'Καταγραφή επιβατών
        Case 2
            If txtShipManagerName.Enabled Then txtShipManagerName.SetFocus
        'Καταχώρηση επιβατών
        Case 3
            If txtShipRepeatedEntriesDescription.Enabled Then txtShipRepeatedEntriesDescription.SetFocus
    End Select

End Sub

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
        Case 5
            AddGridLine
        Case 6
            ToggleGridLineToDelete
    End Select

End Sub
   
Private Sub cmdIndex_Click(index As Integer)

    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Επαναλαμβανόμενη καταχώρηση (Εύρεση τελευταίας εγγραφής)
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtShipRepeatedEntriesDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtShipRepeatedEntriesID.text = tmpTableData.strCode
                txtShipRepeatedEntriesDescription.text = tmpTableData.strFirstField
            End If
        Case 1
            'Αποθήκευση και δημιουργία με ένα κλικ
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtShipSaveAndNewDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtShipSaveAndNewID.text = tmpTableData.strCode
                txtShipSaveAndNewDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    If Me.Tag = "True" Then
        Me.Tag = "False"
        AddColumnsToGrid grdShips, 25, GetSetting(strApplicationName, "Layout Strings", "grdShips"), "04LNID,40LNDescription", "ID,Ονομασία"
        AddColumnsAndCombosToCrewGrid
        Me.Refresh
        PopulateGrid
    End If

    'AddDummyLines grdShips, "99999", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAΑΑΑΑΑΑΑΑΑΑ"

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
        Case vbKeyF3 And cmdButton(2).Enabled, vbKeyD And CtrlDown = 4 And cmdButton(2).Enabled And Not btnPanel(0).Enabled 'Διαγραφή πλοίου, πληρώματος, διαχειριστών
            cmdButton_Click 2
        Case vbKeyInsert And cmdButton(5).Enabled, vbKeyN And CtrlDown = 4 And cmdButton(5).Enabled
            cmdButton_Click 5
        Case vbKeyF3 And cmdButton(6).Enabled, vbKeyD And CtrlDown = 4 And cmdButton(6).Enabled And Not btnPanel(1).Enabled 'Διαγραφή μέλους πληρώματος
            cmdButton_Click 6
        Case vbKeyEscape
            If cmdButton(3).Enabled Then cmdButton_Click 3: Exit Function
            If cmdButton(4).Enabled Then cmdButton_Click 4
        Case vbKeyPageUp
            GotoPreviousPanel Me, btnPanel.Count
        Case vbKeyPageDown
            GotoNextPanel Me, btnPanel.Count
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Function GotoNextPanel(formName, ParamArray panels())

    Dim intLoop As Integer
    
    For intLoop = 0 To btnPanel.Count - 1
    
        If Not btnPanel(intLoop).Enabled Then
            If intLoop + 1 <= btnPanel.Count - 1 Then
                If btnPanel(intLoop + 1).Enabled Then
                    btnPanel_Click intLoop + 1
                    Exit Function
                End If
            End If
        End If
    
    Next intLoop

End Function

Private Function GotoPreviousPanel(formName, intPanelCount)

    Dim intLoop As Integer
    
    For intLoop = 0 To btnPanel.Count - 1
    
        If Not btnPanel(intLoop).Enabled Then
            If intLoop - 1 >= 0 Then
                If btnPanel(intLoop - 1).Enabled Then
                    btnPanel_Click intLoop - 1
                    Exit Function
                End If
            End If
        End If
    
    Next intLoop

End Function

Private Sub Form_Load()
    
    UpdateColors Me, False, , False
    PositionPanels
    SetUpGrid lstIconList, grdShips
    SetUpGrid lstIconList, grdCrew
    ClearFields txtShipID, txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerID, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipRepeatedEntriesID, txtShipSaveAndNewDescription, txtShipSaveAndNewID
    DisableFields txtShipDescription, mskShipPersons, txtShipFlag, txtShipRegistryNo, txtShipIMO, txtShipSalesCode, grdCrew, txtShipManagerName, txtShipManagerNameInGreece, txtShipManagerAgent, txtShipManagerAPersonName, txtShipManagerAPersonPhones, txtShipManagerAPersonEmail, txtShipManagerAPersonFax, txtShipManagerAPersonAddress, txtShipManagerBPersonName, txtShipManagerBPersonPhones, txtShipManagerBPersonEmail, txtShipManagerBPersonFax, txtShipManagerBPersonAddress, txtShipRepeatedEntriesDescription, txtShipSaveAndNewDescription, btnPanel(1), btnPanel(2), btnPanel(3)
    UpdateButtons Me, 4, 1, 0, 0, 0, 1
    
End Sub

Private Function PositionPanels()

    Dim intLoop As Integer
    
    For intLoop = 0 To 3
        frmFrame(intLoop).Visible = False
    Next intLoop
        
    For intLoop = 0 To 3
        btnPanel(intLoop).Enabled = True
        shpBridge(intLoop).Visible = False
        With frmFrame(intLoop)
            .Height = 6465
            .Width = 12465
            .Left = 1875
            .Top = 1125
            .BackColor = &HE0E0E0
        End With
    Next intLoop
    
    btnPanel(0).Enabled = False
    frmFrame(0).Visible = True
    shpBridge(0).Visible = True

End Function

Private Sub grdCrew_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)

    MoveToNextColumn grdCrew, lRow, lCol + 1
    
End Sub

Private Sub grdCrew_ColHeaderMouseEnter(ByVal lCol As Long)

    grdCrew.Header.Buttons = True

End Sub

Private Sub grdCrew_ColHeaderMouseLeave(ByVal lCol As Long)

    grdCrew.Header.Buttons = False

End Sub

Private Sub grdCrew_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp
    
End Sub

Private Sub grdCrew_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid300_10Tec.ETextEditFlags)

    If lCol = 1 Or lCol >= 7 Then bCancel = True

End Sub

Private Sub grdShips_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)

    SeekRecord

End Sub

Private Sub grdShips_HeaderRightClick(ByVal lCol As Long, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    PopupMenu mnuHdrPopUp

End Sub

Private Sub grdShips_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then SeekRecord
    
End Sub

Private Sub mnuΑποθήκευσηΠλάτουςΣτηλών_Click()

    SaveSetting strApplicationName, "Layout Strings", "grdShips", grdShips.LayoutCol
    SaveSetting strApplicationName, "Layout Strings", "grdCrew", grdCrew.LayoutCol

End Sub

Private Function AddColumnsAndCombosToCrewGrid()

    Dim strSavedLayout As String
    
    'Local recordsets
    Dim TableOccupantsDescriptions As TableDef
    Dim TableAges As TableDef
    Dim TableGenders As TableDef
    
    Dim rsOccupantsDescriptions As Recordset
    Dim rsAges As Recordset
    Dim rsGenders As Recordset
    
    'Αρχικές τιμές
    Set TableOccupantsDescriptions = dBaseTables("OccupantsDescriptions")
    Set TableGenders = dBaseTables("Genders")
    Set TableAges = dBaseTables("Ages")
    
    Set rsAges = TableAges.OpenRecordset()
    Set rsGenders = TableGenders.OpenRecordset()
    Set rsOccupantsDescriptions = TableOccupantsDescriptions.OpenRecordset()
    
    With grdCrew
        
        'Ιδιότητα
        With .Combos.Add("PropertyDescription")
            While Not rsOccupantsDescriptions.EOF
                .AddItem sItemText:=rsOccupantsDescriptions!OccupantDescriptionDescription, vItemValue:=CLng(rsOccupantsDescriptions!OccupantDescriptionID)
                rsOccupantsDescriptions.MoveNext
            Wend
            rsOccupantsDescriptions.Close
        End With
        'Ηλικία
        With .Combos.Add("AgeDescription")
            While Not rsAges.EOF
                .AddItem sItemText:=rsAges!AgeDescription, vItemValue:=CLng(rsAges!AgeID)
                rsAges.MoveNext
            Wend
            rsAges.Close
        End With
        'Φύλο
        With .Combos.Add("GenderDescription")
            While Not rsGenders.EOF
                .AddItem sItemText:=rsGenders!GenderDescription, vItemValue:=CLng(rsGenders!GenderID)
                rsGenders.MoveNext
            Wend
            rsGenders.Close
        End With
        With .AddCol(sKey:="ID", sHeader:="ID", lWidth:=175, eHdrTextFlags:=igTextCenter)
            .eTextFlags = igTextCenter
        End With
        With .AddCol(sKey:="LastName", sHeader:="Επίθετο", lWidth:=75, eHdrTextFlags:=igTextCenter)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="FirstName", sHeader:="Ονομα", lWidth:=75, eHdrTextFlags:=igTextCenter)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="PropertyDescription", sHeader:="Ιδιότητα", lWidth:=32, lMinWidth:=24, eHdrTextFlags:=igTextCenter)
           .eType = igCellCombo
           .sCtrlKey = "PropertyDescription"
        End With
        With .AddCol(sKey:="GenderDescription", sHeader:="Φύλο", lWidth:=32, lMinWidth:=24, eHdrTextFlags:=igTextCenter)
           .eType = igCellCombo
           .sCtrlKey = "GenderDescription"
        End With
        With .AddCol(sKey:="AgeDescription", sHeader:="Ηλικία", lWidth:=32, lMinWidth:=24, eHdrTextFlags:=igTextCenter)
           .eType = igCellCombo
           .sCtrlKey = "AgeDescription"
        End With
        With .AddCol(sKey:="Status", sHeader:="", lWidth:=202, eHdrTextFlags:=igTextCenter) 'Blank = Edit, Blue = New
            .eTextFlags = igTextCenter
        End With
        With .AddCol(sKey:="Deleted", sHeader:="", lWidth:=202, eHdrTextFlags:=igTextCenter) 'Red = Mark To Delete
            .eTextFlags = igTextCenter
        End With
        
        .RowMode = False
        .Editable = True
        .ImageList = lstIconList
        
    End With
    
    strSavedLayout = GetSetting(strApplicationName, "Layout Strings", "grdCrew"): If strSavedLayout <> "" Then grdCrew.LayoutCol = strSavedLayout

End Function

Private Function ToggleGridLineToDelete()

    With grdCrew
        If .RowCount > 0 Then
            .CellIcon(.CurRow, "Deleted") = IIf(.CellIcon(.CurRow, "Deleted") <= 0, lstIconList.ItemIndex(3), lstIconList.ItemIndex(1))
        End If
    End With

End Function

Private Function SaveCrew(shipID As Integer)

    On Error GoTo ErrTrap
    
    Dim lngID As Long
    Dim lngRow As Long
    
    With grdCrew
        For lngRow = 1 To .RowCount
        '    'Add Record when Status = Blue and Deleted = Blank
            If (.CellIcon(lngRow, "Status") = 1) And (.CellIcon(lngRow, "Deleted") = -1) Then
                lngID = MainSaveRecord("CommonDB", "ShipsCrew", True, strApplicationName, "ID", 0, Trim(Left(.CellValue(lngRow, "LastName"), 30)), Trim(Left(.CellValue(lngRow, "FirstName"), 10)), .CellValue(lngRow, "PropertyDescription"), .CellValue(lngRow, "GenderDescription"), .CellValue(lngRow, "AgeDescription"), txtShipID.text, txtShipID.text, strCurrentUser)
            End If
        '    'Delete Existing Record when Status = Blank and Deleted = Red
            If (.CellIcon(lngRow, "Status") = -1) And (.CellIcon(lngRow, "Deleted") = 2) Then
                lngID = MainDeleteRecord("CommonDB", "ShipsCrew", strApplicationName, "ID", .CellValue(lngRow, "ID"), False)
            End If
        '    'Update Existing Record when Status = Blank and Deleted = Blank
            If (.CellIcon(lngRow, "Status") = -1) And (.CellIcon(lngRow, "Deleted") = -1) Then
                lngID = MainSaveRecord("CommonDB", "ShipsCrew", False, strApplicationName, "ID", .CellValue(lngRow, "ID"), Trim(Left(.CellValue(lngRow, "LastName"), 30)), Trim(Left(.CellValue(lngRow, "FirstName"), 10)), .CellValue(lngRow, "PropertyDescription"), .CellValue(lngRow, "GenderDescription"), .CellValue(lngRow, "AgeDescription"), txtShipID.text, txtShipID.text, strCurrentUser)
            End If
        Next lngRow
    End With
    
    SaveCrew = shipID
    
    Exit Function

ErrTrap:
    SaveCrew = shipID
    DisplayErrorMessage True, Err.Description

End Function

Private Sub txtShipRepeatedEntriesDescription_Change()
    
    If txtShipRepeatedEntriesDescription.text = "" Then
        ClearFields txtShipRepeatedEntriesID
    End If

End Sub

Private Sub txtShipRepeatedEntriesDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtShipRepeatedEntriesDescription_Validate(Cancel As Boolean)

    If txtShipRepeatedEntriesID.text = "" And txtShipRepeatedEntriesDescription.text <> "" Then cmdIndex_Click 0

End Sub

Private Sub txtShipSaveAndNewDescription_Change()

    If txtShipSaveAndNewDescription.text = "" Then
        ClearFields txtShipSaveAndNewID
    End If

End Sub

Private Sub txtShipSaveAndNewDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtShipSaveAndNewDescription_Validate(Cancel As Boolean)

    If txtShipSaveAndNewID.text = "" And txtShipSaveAndNewDescription.text <> "" Then cmdIndex_Click 1

End Sub

