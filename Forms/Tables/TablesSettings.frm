VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form TablesSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   13335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13335
   ScaleWidth      =   17880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFrame 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   6090
      Index           =   0
      Left            =   10800
      TabIndex        =   21
      Top             =   8850
      Width           =   9165
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Επικεφαλίδες παραστατικών "
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
         Height          =   3990
         Index           =   0
         Left            =   450
         TabIndex        =   22
         Top             =   525
         Width           =   8265
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   1
            Left            =   1650
            TabIndex        =   0
            Top             =   525
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
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   2
            Left            =   1650
            TabIndex        =   1
            Top             =   1050
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   3
            Left            =   1650
            TabIndex        =   2
            Top             =   1575
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   4
            Left            =   1650
            TabIndex        =   3
            Top             =   2100
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   5
            Left            =   1650
            TabIndex        =   4
            Top             =   2625
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   6
            Left            =   1650
            TabIndex        =   5
            Top             =   3150
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   13
            Left            =   7800
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
            Index           =   12
            Left            =   1200
            Top             =   750
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   390
            Index           =   11
            Left            =   2475
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
            Index           =   10
            Left            =   2550
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
            Index           =   9
            Left            =   0
            Top             =   1950
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "6η Γραμμή"
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
            TabIndex        =   36
            Top             =   3225
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "5η Γραμμή"
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
            TabIndex        =   35
            Top             =   2700
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "4η Γραμμή"
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
            TabIndex        =   34
            Top             =   2175
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "3η Γραμμή"
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
            TabIndex        =   33
            Top             =   1650
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "2η Γραμμή"
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
            TabIndex        =   32
            Top             =   1125
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "1η Γραμμή"
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
            TabIndex        =   31
            Top             =   600
            Width           =   750
         End
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   20
         Left            =   4275
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
         Index           =   0
         Left            =   8700
         Top             =   1275
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   8
         Left            =   0
         Top             =   1275
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   4815
      Index           =   1
      Left            =   5550
      TabIndex        =   23
      Top             =   10650
      Width           =   9165
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Επικεφαλίδες αναφορών "
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
         Height          =   2940
         Index           =   1
         Left            =   450
         TabIndex        =   24
         Top             =   525
         Width           =   8265
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   7
            Left            =   1650
            TabIndex        =   6
            Top             =   525
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   8
            Left            =   1650
            TabIndex        =   7
            Top             =   1050
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   9
            Left            =   1650
            TabIndex        =   8
            Top             =   1575
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin UserControls.newText txtCompanyData 
            Height          =   465
            Index           =   10
            Left            =   1650
            TabIndex        =   9
            Top             =   2100
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   820
            ForeColor       =   0
            MaxLength       =   50
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
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   31
            Left            =   1200
            Top             =   1050
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   18
            Left            =   7800
            Top             =   1200
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   840
            Index           =   17
            Left            =   0
            Top             =   1125
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   390
            Index           =   16
            Left            =   3225
            Top             =   2550
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape shpWedge 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00008000&
            Height          =   540
            Index           =   15
            Left            =   2550
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "4η Γραμμή"
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
            TabIndex        =   30
            Top             =   2175
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "3η Γραμμή"
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
            TabIndex        =   29
            Top             =   1650
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "2η Γραμμή"
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
            TabIndex        =   28
            Top             =   1125
            Width           =   750
         End
         Begin VB.Label lblLabel 
            BackColor       =   &H000080FF&
            Caption         =   "1η Γραμμή"
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
            Top             =   600
            Width           =   750
         End
      End
      Begin UserControls.newText txtPreviewReportsDescription 
         Height          =   465
         Left            =   2850
         TabIndex        =   20
         Top             =   3600
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
         Index           =   2
         Left            =   3525
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3600
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
         PicNormal       =   "TablesSettings.frx":0000
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   540
         Index           =   19
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
         Height          =   840
         Index           =   34
         Left            =   2400
         Top             =   3600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Προεπισκόπηση αναφορών"
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
         Index           =   27
         Left            =   450
         TabIndex        =   50
         Top             =   3675
         Width           =   1965
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   14
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame frmButtonFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   61
      Top             =   7650
      Width           =   6090
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
         Caption         =   "Επεξεργασία"
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
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   4890
      Index           =   2
      Left            =   11475
      TabIndex        =   25
      Top             =   4200
      Width           =   9240
      Begin UserControls.newText txtSalesAccountsCode 
         Height          =   465
         Left            =   2775
         TabIndex        =   11
         Top             =   1050
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
      Begin UserControls.newText txtVATAccountsCode 
         Height          =   465
         Left            =   2775
         TabIndex        =   12
         Top             =   1575
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
      Begin UserControls.newInteger mskCustomerCodeLength 
         Height          =   465
         Left            =   2775
         TabIndex        =   13
         Top             =   2100
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
      Begin UserControls.newText txtFileName 
         Height          =   465
         Left            =   2775
         TabIndex        =   14
         Top             =   2625
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
      Begin UserControls.newText txtCashAccountsCode 
         Height          =   465
         Left            =   2775
         TabIndex        =   10
         Top             =   525
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
         Caption         =   "Κωδ. ταμείου"
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
         Left            =   450
         TabIndex        =   60
         Top             =   600
         Width           =   1890
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   4
         Left            =   2325
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
         Index           =   3
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
         Height          =   540
         Index           =   2
         Left            =   0
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Μήκος κωδικών"
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
         Left            =   450
         TabIndex        =   43
         Top             =   2175
         Width           =   1890
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Κωδ. Φ.Π.Α. πωλήσεων"
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
         Index           =   15
         Left            =   450
         TabIndex        =   42
         Top             =   1650
         Width           =   1890
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Ονομα αρχείου πωλήσεων"
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
         TabIndex        =   41
         Top             =   2700
         Width           =   1890
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
         TabIndex        =   40
         Top             =   1125
         Width           =   1890
      End
   End
   Begin VB.Frame frmFrame 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   6015
      Index           =   3
      Left            =   1875
      TabIndex        =   26
      Top             =   1125
      Width           =   9240
      Begin UserControls.newInteger mskVAT 
         Height          =   465
         Left            =   3375
         TabIndex        =   15
         Top             =   525
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
      Begin UserControls.newInteger mskInvoiceCopies 
         Height          =   465
         Left            =   3375
         TabIndex        =   16
         Top             =   1050
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
      Begin UserControls.newText txtPreviewInvoicesDescription 
         Height          =   465
         Left            =   3375
         TabIndex        =   17
         Top             =   1575
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
      Begin UserControls.newText txtUsualPaymentTermDescription 
         Height          =   465
         Left            =   3375
         TabIndex        =   18
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
      Begin UserControls.newText txtUsualRemarks 
         Height          =   465
         Left            =   3375
         TabIndex        =   19
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
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   0
         Left            =   4050
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1575
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
         PicNormal       =   "TablesSettings.frx":059A
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton cmdIndex 
         Height          =   465
         Index           =   1
         Left            =   8400
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2100
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
         PicNormal       =   "TablesSettings.frx":0B34
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
         Left            =   8775
         Top             =   1950
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Συνήθης αιτιολογία"
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
         Left            =   450
         TabIndex        =   56
         Top             =   2700
         Width           =   2490
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Συνήθης όρος πληρωμής"
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
         TabIndex        =   53
         Top             =   2175
         Width           =   2490
      End
      Begin VB.Shape shpWedge 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   840
         Index           =   1
         Left            =   2925
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
         Height          =   540
         Index           =   5
         Left            =   0
         Top             =   450
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Προεπισκόπηση παραστατικών"
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
         TabIndex        =   39
         Top             =   1650
         Width           =   2490
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Πλήθος αντιγράφων παραστατικών"
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
         TabIndex        =   38
         Top             =   1125
         Width           =   2490
      End
      Begin VB.Label lblLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Ποσοστό Φ.Π.Α."
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
         TabIndex        =   37
         Top             =   600
         Width           =   2490
      End
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1890
      Left            =   225
      TabIndex        =   45
      Top             =   9675
      Width           =   4515
      Begin VB.TextBox txtUsualPaymentTermID 
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
         Top             =   1350
         Width           =   780
      End
      Begin VB.TextBox Text22 
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
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "Settings.UsualPaymentTermID"
         Top             =   1350
         Width           =   3540
      End
      Begin VB.TextBox txtPreviewReportsID 
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   975
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
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Settings.PreviewReportsID"
         Top             =   975
         Width           =   3540
      End
      Begin VB.TextBox txtID 
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   225
         Width           =   780
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
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Settings.ID"
         Top             =   225
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
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "Settings.PreviewInvoicesID"
         Top             =   600
         Width           =   3540
      End
      Begin VB.TextBox txtPreviewInvoicesID 
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   600
         Width           =   780
      End
   End
   Begin Dacara_dcButton.dcButton btnPanel 
      Height          =   990
      Index           =   0
      Left            =   450
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Φορολογικά στοιχεία"
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
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Αναφορές"
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
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Αρχείο Γεν. Λογιστικής"
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
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1746
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Πωλήσεις"
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
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   3225
      Top             =   7125
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   840
      Index           =   30
      Left            =   0
      Top             =   1425
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   11100
      Top             =   2025
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   2625
      Top             =   8325
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Παραμετροποίηση"
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
      TabIndex        =   44
      Top             =   75
      Width           =   4305
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   3
      Left            =   450
      Top             =   4275
      Width           =   3090
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   2
      Left            =   450
      Top             =   3225
      Width           =   3090
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   1
      Left            =   450
      Top             =   2175
      Width           =   3090
   End
   Begin VB.Shape shpBridge 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Index           =   0
      Left            =   450
      Top             =   1125
      Width           =   3090
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
Attribute VB_Name = "TablesSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnStatus As Boolean

Private Function PositionPanels()

    Dim intLoop As Integer
    
    For intLoop = 0 To 3
        frmFrame(intLoop).Visible = False
    Next intLoop
        
    For intLoop = 0 To 3
        btnPanel(intLoop).Enabled = True
        shpBridge(intLoop).Visible = False
        With frmFrame(intLoop)
            .Height = 6015
            .Width = 9165
            .Left = 1875
            .Top = 1125
            .BackColor = &HE0E0E0
        End With
    Next intLoop
    
    btnPanel(0).Enabled = False
    frmFrame(0).Visible = True
    shpBridge(0).Visible = True

End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Ονομα αρχείου πωλήσεων
    If Len(txtFileName.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 2
        txtFileName.SetFocus
        Exit Function
    End If
    
    '% Φ.Π.Α.
    If mskVAT.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 3
        mskVAT.SetFocus
        Exit Function
    End If
    
    If mskVAT.text = "0" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        btnPanel_Click 3
        mskVAT.SetFocus
        Exit Function
    End If
   
    'Πλήθος αντιγράφων παραστατικών
    If mskInvoiceCopies.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        btnPanel_Click 3
        mskInvoiceCopies.SetFocus
        Exit Function
    End If

    ValidateFields = True

End Function

Private Function AbortProcedure(blnStatus)

    If Not blnStatus Then
        If MyMsgBox(3, strApplicationName, strStandardMessages(3), 2) Then
            btnPanel_Click 0
            blnStatus = True
            DisableFields txtCompanyData(1), txtCompanyData(2), txtCompanyData(3), txtCompanyData(4), txtCompanyData(5), txtCompanyData(6), txtCompanyData(7), txtCompanyData(8), txtCompanyData(9), txtCompanyData(10), txtPreviewReportsDescription, txtUsualPaymentTermDescription, txtUsualRemarks, txtCashAccountsCode, txtSalesAccountsCode, txtVATAccountsCode, mskCustomerCodeLength, txtFileName, mskVAT, mskInvoiceCopies, txtPreviewInvoicesDescription, txtUsualPaymentTermDescription, cmdIndex(0), cmdIndex(1), cmdIndex(2)
            UpdateButtons Me, 3, 1, 0, 0, 1
        End If
        Exit Function
    End If
    
    If blnStatus Then
        Unload Me
    End If

End Function

Private Function EditRecord()

    Dim intLoop As Integer
    
    blnStatus = False
    
    EnableFields txtCompanyData(1), txtCompanyData(2), txtCompanyData(3), txtCompanyData(4), txtCompanyData(5), txtCompanyData(6), txtCompanyData(7), txtCompanyData(8), txtCompanyData(9), txtCompanyData(10), txtPreviewReportsDescription, txtUsualPaymentTermDescription, txtUsualRemarks, txtCashAccountsCode, txtSalesAccountsCode, txtVATAccountsCode, mskCustomerCodeLength, txtFileName, mskVAT, mskInvoiceCopies, txtPreviewInvoicesDescription, cmdIndex(0), cmdIndex(1), cmdIndex(2)
    
    UpdateButtons Me, 3, 0, 1, 1, 0
    
    For intLoop = 0 To btnPanel.Count - 1
        If Not btnPanel(intLoop).Enabled Then btnPanel_Click intLoop
    Next intLoop
    
End Function

Private Function SaveRecord()

    If Not ValidateFields Then Exit Function
    
    If MainSaveRecord("CommonDB", "Settings", False, "Settings", "ID", txtID.text, txtCompanyData(1).text, txtCompanyData(2).text, txtCompanyData(3).text, txtCompanyData(4).text, txtCompanyData(5).text, txtCompanyData(6).text, txtCompanyData(7).text, txtCompanyData(8).text, txtCompanyData(9).text, txtCompanyData(10).text, txtPreviewReportsID.text, IIf(txtUsualPaymentTermID.text <> "", txtUsualPaymentTermID.text, "0"), txtUsualRemarks.text, txtCashAccountsCode.text, txtSalesAccountsCode.text, txtVATAccountsCode.text, mskCustomerCodeLength.text, txtFileName.text, mskVAT.text, mskInvoiceCopies.text, txtPreviewInvoicesID.text) <> 0 Then
        btnPanel_Click 0
        blnStatus = True
        DisableFields txtCompanyData(1), txtCompanyData(2), txtCompanyData(3), txtCompanyData(4), txtCompanyData(5), txtCompanyData(6), txtCompanyData(7), txtCompanyData(8), txtCompanyData(9), txtCompanyData(10), txtCashAccountsCode, txtSalesAccountsCode, txtVATAccountsCode, mskCustomerCodeLength, txtFileName, mskVAT, mskInvoiceCopies, txtPreviewInvoicesDescription, txtUsualPaymentTermDescription, txtUsualRemarks, cmdIndex(0), cmdIndex(1), cmdIndex(2)
        UpdateButtons Me, 3, 1, 0, 0, 1
        If MyMsgBox(1, strApplicationName, strStandardMessages(21), 1) Then
        End If
    Else
        DisplayErrorMessage True, strStandardMessages(5)
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
        'Φορολογικά στοιχεία
        Case 0
            If cmdButton(1).Enabled Then
                If txtCompanyData(1).Enabled Then txtCompanyData(1).SetFocus
            End If
        'Αναφορές
        Case 1
            If cmdButton(1).Enabled Then
                If txtCompanyData(7).Enabled Then txtCompanyData(7).SetFocus
            End If
        'Αρχείο γεν. λογιστικής
        Case 2
            If cmdButton(1).Enabled Then
                If txtCashAccountsCode.Enabled Then txtCashAccountsCode.SetFocus
            End If
        'Πωλήσεις
        Case 3
            If cmdButton(1).Enabled Then
                If mskVAT.Enabled Then mskVAT.SetFocus
            End If
    End Select

End Sub

Private Sub cmdButton_Click(index As Integer)

    Select Case index
        Case 0
            EditRecord
        Case 1
            SaveRecord
        Case 2
            AbortProcedure False
        Case 3
            AbortProcedure True
    End Select

End Sub

Private Sub cmdIndex_Click(index As Integer)
    
    'Local variables
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Προεπισκόπηση παραστατικών
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtPreviewInvoicesDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtPreviewInvoicesID.text = tmpTableData.strCode
                txtPreviewInvoicesDescription.text = tmpTableData.strFirstField
            End If
        Case 1
            'Προεπισκόπηση αναφορών
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtPreviewReportsDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtPreviewReportsID.text = tmpTableData.strCode
                txtPreviewReportsDescription.text = tmpTableData.strFirstField
            End If
        Case 2
            'Όρος πληρωμής
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoDescription", "String", txtUsualPaymentTermDescription.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 3, 0, 1, 2, "ID", "Περιγραφή", "Πίστωση", 0, 40, 0, 1, 0, 0)
                txtUsualPaymentTermID.text = tmpTableData.strCode
                txtUsualPaymentTermDescription.text = tmpTableData.strFirstField
            End If
    End Select

End Sub

Private Sub Form_Activate()

    Dim tmpRecordset As Recordset

    If Me.Tag = "True" Then
        Me.Tag = "False"
        Me.Refresh
        If MainSeekRecord("CommonDB", "Settings", "ID", 1, True, txtID, txtCompanyData(1), txtCompanyData(2), txtCompanyData(3), txtCompanyData(4), txtCompanyData(5), txtCompanyData(6), txtCompanyData(7), txtCompanyData(8), txtCompanyData(9), txtCompanyData(10), txtPreviewReportsID, txtUsualPaymentTermID, txtUsualRemarks, txtCashAccountsCode, txtSalesAccountsCode, txtVATAccountsCode, mskCustomerCodeLength, txtFileName, mskVAT, mskInvoiceCopies, txtPreviewInvoicesID) Then
            'Προεπισκόπηση παραστατικών
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtPreviewInvoicesID.text)
            txtPreviewInvoicesID.text = tmpRecordset.Fields(0)
            txtPreviewInvoicesDescription.text = tmpRecordset.Fields(1)
            'Προεπισκόπηση αναφορών
            Set tmpRecordset = CheckForMatch("CommonDB", "YesOrNo", "YesOrNoID", "Numeric", txtPreviewReportsID.text)
            txtPreviewReportsID.text = tmpRecordset.Fields(0)
            txtPreviewReportsDescription.text = tmpRecordset.Fields(1)
            'Συνήθης όρος πληρωμής
            Set tmpRecordset = CheckForMatch("CommonDB", "PaymentTerms", "PaymentTermID", "Numeric", txtUsualPaymentTermID.text)
            If tmpRecordset.RecordCount > 0 Then
                txtUsualPaymentTermID.text = tmpRecordset.Fields(0)
                txtUsualPaymentTermDescription.text = tmpRecordset.Fields(1)
            Else
                ClearFields txtUsualPaymentTermID, txtUsualPaymentTermDescription
            End If
            '
            UpdateButtons Me, 3, 1, 0, 0, 1
        End If
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
        Case vbKeyE And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyF10 And cmdButton(1).Enabled, vbKeyS And CtrlDown = 4 And cmdButton(1).Enabled
            cmdButton_Click 1
        Case vbKeyEscape
            If cmdButton(2).Enabled Then cmdButton_Click 2: Exit Function
            If cmdButton(3).Enabled Then cmdButton_Click 3
        Case vbKeyPageUp
            GotoPreviousPanel Me, btnPanel.Count
        Case vbKeyPageDown
            GotoNextPanel Me, btnPanel.Count
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Function

Private Function GotoNextPanel(formName, panelCount)

    Dim intLoop As Integer
    
    For intLoop = 0 To panelCount - 1
    
        If Not formName.btnPanel(intLoop).Enabled Then
            If intLoop + 1 <= formName.btnPanel.Count - 1 Then
                If formName.btnPanel(intLoop + 1).Enabled Then
                    btnPanel_Click intLoop + 1
                    Exit Function
                End If
            End If
        End If
    
    Next intLoop

End Function

Private Function GotoPreviousPanel(formName, intPanelCount)

    Dim intLoop As Integer
    
    For intLoop = 0 To formName.btnPanel.Count - 1
    
        If Not formName.btnPanel(intLoop).Enabled Then
            If intLoop - 1 >= 0 Then
                If formName.btnPanel(intLoop - 1).Enabled Then
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
    ClearFields txtCompanyData(1), txtCompanyData(2), txtCompanyData(3), txtCompanyData(4), txtCompanyData(5), txtCompanyData(6), txtCompanyData(7), txtCompanyData(8), txtCompanyData(9), txtCompanyData(10), txtPreviewReportsID, txtPreviewReportsDescription, txtUsualPaymentTermID, txtUsualPaymentTermDescription, txtUsualRemarks, txtCashAccountsCode, txtSalesAccountsCode, txtVATAccountsCode, mskCustomerCodeLength, txtFileName, mskVAT, mskInvoiceCopies, txtPreviewInvoicesDescription
    DisableFields txtCompanyData(1), txtCompanyData(2), txtCompanyData(3), txtCompanyData(4), txtCompanyData(5), txtCompanyData(6), txtCompanyData(7), txtCompanyData(8), txtCompanyData(9), txtCompanyData(10), txtPreviewReportsDescription, txtUsualPaymentTermDescription, txtUsualRemarks, txtCashAccountsCode, txtSalesAccountsCode, txtVATAccountsCode, mskCustomerCodeLength, txtFileName, mskVAT, mskInvoiceCopies, txtPreviewInvoicesDescription, txtUsualPaymentTermDescription, cmdIndex(0), cmdIndex(1), cmdIndex(2)
    UpdateButtons Me, 3, 1, 0, 0, 1
    
End Sub

Private Sub mskInvoiceCopies_LostFocus()

    If mskInvoiceCopies.text = "" Then mskInvoiceCopies.text = "0"
    
End Sub

Private Sub mskVAT_LostFocus()

    If mskVAT.text = "" Then mskVAT.text = "0"

End Sub

Private Sub txtPreviewInvoicesDescription_Change()

    If txtPreviewInvoicesDescription.text = "" Then
        ClearFields txtPreviewInvoicesID
    End If

End Sub

Private Sub txtPreviewInvoicesDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0

End Sub

Private Sub txtPreviewInvoicesDescription_Validate(Cancel As Boolean)

    If txtPreviewInvoicesID.text = "" And txtPreviewInvoicesDescription.text <> "" Then cmdIndex_Click 0

End Sub

Private Sub txtPreviewReportsDescription_Change()

    If txtPreviewReportsDescription.text = "" Then
        ClearFields txtPreviewReportsID
    End If

End Sub

Private Sub txtPreviewReportsDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1

End Sub

Private Sub txtPreviewReportsDescription_Validate(Cancel As Boolean)

    If txtPreviewReportsID.text = "" And txtPreviewReportsDescription.text <> "" Then cmdIndex_Click 1

End Sub

Private Sub txtUsualPaymentTermDescription_Change()

    If txtUsualPaymentTermDescription.text = "" Then
        ClearFields txtUsualPaymentTermID
    End If

End Sub

Private Sub txtUsualPaymentTermDescription_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2

End Sub

Private Sub txtUsualPaymentTermDescription_Validate(Cancel As Boolean)

    If txtUsualPaymentTermID.text = "" And txtUsualPaymentTermDescription.text <> "" Then cmdIndex_Click 2

End Sub

