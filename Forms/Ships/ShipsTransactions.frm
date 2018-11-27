VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form ShipsTransactions 
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
      TabIndex        =   53
      Top             =   7350
      Width           =   8940
      Begin Dacara_dcButton.dcButton cmdButton 
         Height          =   690
         Index           =   0
         Left            =   225
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   57
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
         TabIndex        =   58
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
         TabIndex        =   59
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
      Height          =   3465
      Left            =   10350
      TabIndex        =   23
      Top             =   75
      Width           =   4515
      Begin VB.TextBox txtShipSaveAndNewID 
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3075
         Width           =   780
      End
      Begin VB.TextBox Text10 
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
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "Ships.ShipSaveAndNewID"
         Top             =   3075
         Width           =   3540
      End
      Begin VB.TextBox txtShipRepeatedEntriesID 
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
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2700
         Width           =   780
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
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "Ships.ShipRepeatedEntriesID"
         Top             =   2700
         Width           =   3540
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Manifest.DestinationID"
         Top             =   825
         Width           =   3540
      End
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   825
         Width           =   780
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
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Manifest.TripID"
         Top             =   75
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
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Manifers.SexID"
         Top             =   1950
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
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Manifest.PropertyID"
         Top             =   1575
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
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Manifest.ShipID"
         Top             =   1200
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
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   75
         Width           =   780
      End
      Begin VB.TextBox txtShipID 
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
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox txtPropertyID 
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1575
         Width           =   780
      End
      Begin VB.TextBox txtSexID 
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1950
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
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Manifest.AgeID"
         Top             =   2325
         Width           =   3540
      End
      Begin VB.TextBox txtAgeID 
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2325
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
         Text            =   "Manifest.RouteID"
         Top             =   450
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
         Top             =   450
         Width           =   780
      End
   End
   Begin UserControls.newDate mskDate 
      Height          =   465
      Left            =   2100
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
   Begin UserControls.newText txtRoute 
      Height          =   465
      Left            =   2100
      TabIndex        =   1
      Top             =   1650
      Width           =   765
      _ExtentX        =   1349
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
   Begin UserControls.newText txtDestination 
      Height          =   465
      Left            =   2100
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
   Begin UserControls.newText txtShip 
      Height          =   465
      Left            =   2100
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
   Begin UserControls.newText txtProperty 
      Height          =   465
      Left            =   2100
      TabIndex        =   4
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
   Begin UserControls.newText txtLastName 
      Height          =   465
      Left            =   2100
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
   Begin UserControls.newText txtFirstName 
      Height          =   465
      Left            =   2100
      TabIndex        =   6
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
   Begin UserControls.newText txtSex 
      Height          =   465
      Left            =   2100
      TabIndex        =   7
      Top             =   4800
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
   Begin UserControls.newText txtAge 
      Height          =   465
      Left            =   2100
      TabIndex        =   8
      Top             =   5325
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
   Begin UserControls.newText txtCare 
      Height          =   465
      Left            =   2100
      TabIndex        =   9
      Top             =   5850
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
   Begin UserControls.newText txtRemarks 
      Height          =   465
      Left            =   2100
      TabIndex        =   10
      Top             =   6375
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
      PicNormal       =   "ShipsTransactions.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   1
      Left            =   7125
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
      PicNormal       =   "ShipsTransactions.frx":059A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   2
      Left            =   7125
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
      PicNormal       =   "ShipsTransactions.frx":0B34
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   3
      Left            =   7125
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
      PicNormal       =   "ShipsTransactions.frx":10CE
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   4
      Left            =   7125
      TabIndex        =   43
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
      PicNormal       =   "ShipsTransactions.frx":1668
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   5
      Left            =   7125
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5325
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
      PicNormal       =   "ShipsTransactions.frx":1C02
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   6
      Left            =   3375
      TabIndex        =   45
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
      PicNormal       =   "ShipsTransactions.frx":219C
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   7
      Left            =   7575
      TabIndex        =   46
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
      PicNormal       =   "ShipsTransactions.frx":2736
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   8
      Left            =   7575
      TabIndex        =   47
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
      PicNormal       =   "ShipsTransactions.frx":2CD0
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Dacara_dcButton.dcButton cmdIndex 
      Height          =   465
      Index           =   9
      Left            =   7575
      TabIndex        =   48
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
      PicNormal       =   "ShipsTransactions.frx":326A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Shape shpBottomEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   3825
      Top             =   8025
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape shpWedge 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   540
      Index           =   1
      Left            =   3600
      Top             =   6825
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpRightEdge 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   840
      Left            =   9375
      Top             =   6825
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
      Left            =   1650
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
      Left            =   0
      Top             =   1950
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblRouteDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Δρομολόγιο"
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
      Left            =   3825
      TabIndex        =   38
      Top             =   1725
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Επιβαίνοντες πλοίων"
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
      TabIndex        =   22
      Top             =   75
      Width           =   4845
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Ιδιότητα"
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
      TabIndex        =   21
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label lblLabel 
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
      Index           =   9
      Left            =   450
      TabIndex        =   20
      Top             =   6450
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Ειδική φροντίδα"
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
      TabIndex        =   19
      Top             =   5925
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Ηλικία"
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
      TabIndex        =   18
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Φύλο"
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
      TabIndex        =   17
      Top             =   4875
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Ονομα"
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
      TabIndex        =   16
      Top             =   4350
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Επίθετο"
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
      TabIndex        =   15
      Top             =   3825
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Πλοίο"
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
      TabIndex        =   14
      Top             =   2775
      Width           =   1215
   End
   Begin VB.Label lblLabel 
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
      Index           =   1
      Left            =   450
      TabIndex        =   13
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H000080FF&
      Caption         =   "Δρομολόγιο"
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
      TabIndex        =   12
      Top             =   1725
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
      Index           =   2
      Left            =   450
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
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
Attribute VB_Name = "ShipsTransactions"
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
            ClearFields txtID, txtRouteID, txtDestinationID, txtShipID, txtPropertyID, txtSexID, txtAgeID
            ClearFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
            ClearFields lblRouteDescription
            DisableFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
            DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
            UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("ShipsRouteReport"), 0, 1), 0, 1
        End If
    End If
    
    If blnStatus Then
        Unload Me
    End If
    
End Function

Private Function DeleteRecord()
    
    If MainDeleteRecord("CommonDB", "Manifest", strApplicationName, "TripID", txtID.text, "True") Then
        ClearFields txtID, txtRouteID, txtDestinationID, txtShipID, txtPropertyID, txtSexID, txtAgeID
        ClearFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
        ClearFields lblRouteDescription
        DisableFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
        UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("ShipsRouteReport"), 0, 1), 0, 1
    End If
    
End Function

Private Function NewRecord()
    
    'Αν έχω επαναλαμβανόμενη καταχώρηση, βρίσκω την τελευταία εγγραφή και την εμφανίζω
    If txtShipRepeatedEntriesID.text = "1" Then
        If txtID.text <> "" Then
            DisplayLastRecord txtID.text
            ClearFields txtLastName, txtFirstName, txtSex, txtSexID, txtAge, txtAgeID, txtCare, txtRemarks
            txtLastName.SetFocus
        End If
    Else
        ClearFields txtID, txtRouteID, txtDestinationID, txtShipID, txtPropertyID, txtSexID, txtAgeID
        ClearFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
        ClearFields lblRouteDescription
        ClearFields mskDate, txtRoute, txtRouteID, lblRouteDescription, txtDestination, txtDestinationID, txtShip, txtShipID, txtProperty, txtPropertyID, txtLastName, txtFirstName, txtSex, txtSexID, txtAge, txtAgeID, txtCare, txtRemarks
    End If
    
    blnStatus = True
    EnableFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
    EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
    UpdateButtons Me, 5, 0, 1, 0, 0, 1, 0
    mskDate.SetFocus
    
    InitializeFields mskDate
    
End Function

Private Function DisplayLastRecord(lngLastRecord)

    If Not SeekRecord("Manifest", lngLastRecord) Then Exit Function

End Function

Private Function SaveRecord()
    
    If Not ValidateFields Then Exit Function
    
    txtID.text = MainSaveRecord("CommonDB", "Manifest", blnStatus, strApplicationName, "TripID", txtID.text, mskDate.text, txtRouteID.text, txtDestinationID.text, txtShipID.text, txtPropertyID.text, txtLastName.text, txtFirstName.text, txtSexID.text, txtAgeID.text, txtCare.text, txtRemarks.text, 1, strCurrentUser)
        
    If txtID.text <> "" Then
        SaveRecord = True
        ClearFields txtRouteID, txtDestinationID, txtShipID, txtPropertyID, txtSexID, txtAgeID
        ClearFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
        ClearFields lblRouteDescription
        DisableFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
        DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
        UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("ShipsRouteReport"), 0, 1), 0, 1
    Else
        DisplayErrorMessage True, strStandardMessages(5)
    End If
    
End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Ημερομηνία
    If Not CheckDate(mskDate.text, strApplicationName) Then
        mskDate.SetFocus
        Exit Function
    End If
    
    If Len(mskDate.text) <> 10 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        mskDate.SetFocus
        Exit Function
    End If
    
    'Δρομολόγιο
    If Len(txtRouteID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtRoute.SetFocus
        Exit Function
    End If
    
    'Προορισμός
    If Len(txtDestinationID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtDestination.SetFocus
        Exit Function
    End If
    
    'Πλοίο
    If Len(txtShipID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtShip.SetFocus
        Exit Function
    End If
    
    'Ιδιότητα
    If Len(txtPropertyID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtProperty.SetFocus
        Exit Function
    End If
    
    'Ονομα
    If Len(txtLastName.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtLastName.SetFocus
        Exit Function
    End If
    If InStr(txtLastName.text, ",") Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtLastName.SetFocus
        Exit Function
    End If
    
    'Ηλικία
    If Len(txtAgeID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtAge.SetFocus
        Exit Function
    End If
    
    'Φύλο
    If Len(txtSexID.text) = 0 Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtSex.SetFocus
        Exit Function
    End If
    
    'Ειδική φροντίδα
    If InStr(txtCare.text, ",") Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtCare.SetFocus
        Exit Function
    End If

    'Παρατηρήσεις
    If InStr(txtRemarks.text, ",") Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(2), 1) Then
        End If
        txtRemarks.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Sub cmdButton_Click(index As Integer)
                                                                                                                                
    Select Case index
        Case 0
            NewRecord
        Case 1
            If SaveRecord And blnStatus Then CheckToCreateNewRecord 'Ελέγχω αν θα κάνω αμέσως νέα εγγραφή
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

    With ShipsRouteReport
        .Tag = "True"
        .Show 1, Me
    End With
    
End Function

Private Sub cmdIndex_Click(index As Integer)

    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    
    Select Case index
        Case 0
            'Δρομολόγιο
            Set tmpRecordset = CheckForMatch("CommonDB", "Routes", "RouteDescription", "String", txtRoute.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 4, 0, 1, 2, 4, "ID", "Συντ.", "Λιμένας εκκίνησης", "Λιμένας τελικού προορισμού", 0, 5, 40, 40, 1, 1, 0, 0)
                txtRouteID.text = tmpTableData.strCode
                txtRoute.text = tmpTableData.strFirstField
                lblRouteDescription.Caption = tmpTableData.strSecondField & " - " & tmpTableData.strThirdField
            End If
        Case 1
            'Προορισμός
            Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationDescription", "String", txtDestination.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 2, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtDestinationID.text = tmpTableData.strCode
                txtDestination.text = tmpTableData.strFirstField
            End If
        Case 2
            'Πλοίο
            Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipDescription", "String", txtShip.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 4, 0, 1, 7, 8, "ID", "Περιγραφή", "Επαναλαμβανόμενη καταχώρηση", "Αποθήκευση και δημιουργία με ενα κλικ", 0, 40, 10, 10, 1, 0, 1, 1)
                txtShipID.text = tmpTableData.strCode
                txtShip.text = tmpTableData.strFirstField
                txtShipRepeatedEntriesID.text = tmpTableData.strSecondField
                txtShipSaveAndNewID.text = tmpTableData.strThirdField
            End If
        Case 3
            'Ιδιότητα
            Set tmpRecordset = CheckForMatch("CommonDB", "OccupantsDescriptions", "OccupantDescriptionDescription", "String", txtProperty.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtPropertyID.text = tmpTableData.strCode
                txtProperty.text = tmpTableData.strFirstField
            End If
        Case 5
            'Ηλικία
            Set tmpRecordset = CheckForMatch("CommonDB", "Ages", "AgeDescription", "String", txtAge.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtAgeID.text = tmpTableData.strCode
                txtAge.text = tmpTableData.strFirstField
            End If
        Case 4
            'Φύλο
            Set tmpRecordset = CheckForMatch("CommonDB", "Genders", "GenderDescription", "String", txtSex.text)
            If tmpRecordset.RecordCount > 0 Then
                tmpTableData = DisplayIndex(tmpRecordset, 2, True, 2, 0, 1, "ID", "Περιγραφή", 0, 40, 1, 0)
                txtSexID.text = tmpTableData.strCode
                txtSex.text = tmpTableData.strFirstField
            End If
        Case 6
            'Δρομολόγιο
            With TablesShipRoutes
                .Tag = "True"
                .Show 1, Me
            End With
        Case 7
            'Προορισμός
            With TablesDestinations
                .Tag = "True"
                .txtShowInList.text = "2"
                .Show 1, Me
            End With
        Case 8
            'Πλοίο
            With TablesShips
                .Tag = "True"
                .Show 1, Me
            End With
        Case 9
            'Ιδιότητα
            With TablesOccupantsDescriptions
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

Public Function SeekRecord(strTable, tmpPersonID)

    On Error GoTo ErrTrap
    
    Dim tmpTableData As typTableData
    Dim tmpRecordset As Recordset
    ClearFields txtID, txtRouteID, txtDestinationID, txtShipID, txtPropertyID, txtSexID, txtAgeID
    ClearFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
    ClearFields lblRouteDescription
    DisableFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
    
    SeekRecord = False
    
    If MainSeekRecord("CommonDB", strTable, "TripID", tmpPersonID, True, txtID, mskDate, txtRouteID, txtDestinationID, txtShipID, txtPropertyID, txtLastName, txtFirstName, txtSexID, txtAgeID, txtCare, txtRemarks) Then
        'Δρομολόγιο
        Set tmpRecordset = CheckForMatch("CommonDB", "Routes", "RouteID", "Numeric", txtRouteID.text)
        txtRouteID.text = tmpRecordset.Fields(0)
        txtRoute.text = tmpRecordset.Fields(1)
        lblRouteDescription.Caption = tmpRecordset.Fields(2) + " - " + tmpRecordset.Fields(4)
        'Προορισμός
        Set tmpRecordset = CheckForMatch("CommonDB", "Destinations", "DestinationID", "Numeric", txtDestinationID.text)
        txtDestinationID.text = tmpRecordset.Fields(0)
        txtDestination.text = tmpRecordset.Fields(2)
        'Πλοίο
        Set tmpRecordset = CheckForMatch("CommonDB", "Ships", "ShipID", "Numeric", txtShipID.text)
        txtShipID.text = tmpRecordset.Fields(0)
        txtShip.text = tmpRecordset.Fields(1)
        txtShipRepeatedEntriesID.text = tmpRecordset.Fields(7)
        txtShipSaveAndNewID.text = tmpRecordset.Fields(8)
        'Ιδιότητα
        Set tmpRecordset = CheckForMatch("CommonDB", "OccupantsDescriptions", "OccupantDescriptionID", "Numeric", txtPropertyID.text)
        txtPropertyID.text = tmpRecordset.Fields(0)
        txtProperty.text = tmpRecordset.Fields(1)
        'Φύλο
        Set tmpRecordset = CheckForMatch("CommonDB", "Genders", "GenderID", "Numeric", txtSexID.text)
        txtSexID.text = tmpRecordset.Fields(0)
        txtSex.text = tmpRecordset.Fields(1)
        'Ηλικία
        Set tmpRecordset = CheckForMatch("CommonDB", "Ages", "AgeID", "Numeric", txtAgeID.text)
        txtAgeID.text = tmpRecordset.Fields(0)
        txtAge.text = tmpRecordset.Fields(1)
        'Τα υπόλοιπα
        EnableFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
        EnableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
        UpdateButtons Me, 5, 0, 1, 1, 0, 1, 0
        blnStatus = False
        SeekRecord = True
    End If
    
    Exit Function
    
ErrTrap:
    blnErrors = True
    DisplayErrorMessage True, Err.Description

End Function

Private Sub CheckFunctionKeys(KeyCode, Shift)
    
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
            If cmdButton(4).Enabled Then cmdButton_Click 4: Exit Sub
            If cmdButton(5).Enabled Then cmdButton_Click 5
        Case vbKeyF12 And CtrlDown = 4
            ToggleInfoPanel Me
    End Select

End Sub

Private Sub Form_Load()

    UpdateColors Me, False

    ClearFields txtID, txtRouteID, txtDestinationID, txtShipID, txtPropertyID, txtSexID, txtAgeID
    ClearFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
    ClearFields lblRouteDescription
    DisableFields mskDate, txtRoute, txtDestination, txtShip, txtProperty, txtLastName, txtFirstName, txtSex, txtAge, txtCare, txtRemarks
    DisableFields cmdIndex(0), cmdIndex(1), cmdIndex(2), cmdIndex(3), cmdIndex(4), cmdIndex(5), cmdIndex(6), cmdIndex(7), cmdIndex(8), cmdIndex(9)
    UpdateButtons Me, 5, 1, 0, 0, IIf(CheckForLoadedForm("ShipsRouteReport"), 0, 1), 0, 1

End Sub

Private Sub txtAge_Change()

    If txtAge.text = "" Then
        ClearFields txtAgeID
    End If

End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 5

End Sub

Private Sub txtAge_Validate(Cancel As Boolean)

    If txtAgeID.text = "" And txtAge.text <> "" Then cmdIndex_Click 5

End Sub

Private Sub txtSex_Change()

    If txtSex.text = "" Then
        ClearFields txtSexID
    End If

End Sub

Private Sub txtSex_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 4

End Sub

Private Sub txtSex_Validate(Cancel As Boolean)

    If txtSexID.text = "" And txtSex.text <> "" Then cmdIndex_Click 4

End Sub

Private Sub txtProperty_Change()

    If txtProperty.text = "" Then
        ClearFields txtPropertyID
    End If

End Sub

Private Sub txtProperty_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 3
    If KeyCode = vbKeyF5 Then cmdIndex_Click 9

End Sub

Private Sub txtProperty_Validate(Cancel As Boolean)

    If txtPropertyID.text = "" And txtProperty.text <> "" Then cmdIndex_Click 3

End Sub

Private Sub txtRoute_Change()

    If txtRoute.text = "" Then
        ClearFields txtRouteID, lblRouteDescription
    End If

End Sub

Private Sub txtRoute_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 0
    If KeyCode = vbKeyF5 Then cmdIndex_Click 6

End Sub

Private Sub txtRoute_Validate(Cancel As Boolean)

    If txtRouteID.text = "" And txtRoute.text <> "" Then cmdIndex_Click 0

End Sub

Private Sub txtShip_Change()

    If txtShip.text = "" Then
        ClearFields txtShipID
    End If
    
End Sub

Private Sub txtShip_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 2
    If KeyCode = vbKeyF5 Then cmdIndex_Click 8

End Sub

Private Sub txtShip_Validate(Cancel As Boolean)

    If txtShipID.text = "" And txtShip.text <> "" Then cmdIndex_Click 2

End Sub

Private Sub txtDestination_Change()
                                                                
    If txtDestination.text = "" Then
        ClearFields txtDestinationID
    End If

End Sub

Private Sub txtDestination_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then cmdIndex_Click 1
    If KeyCode = vbKeyF5 Then cmdIndex_Click 7

End Sub

Private Sub txtDestination_Validate(Cancel As Boolean)

    If txtDestinationID.text = "" And txtDestination.text <> "" Then cmdIndex_Click 1
    
End Sub

Private Function CheckToCreateNewRecord()

    If txtShipSaveAndNewID.text = "1" Then
        cmdButton_Click 0
    End If

End Function

