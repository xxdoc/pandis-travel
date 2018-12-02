VERSION 5.00
Object = "{E3F0D4E9-96BB-4A6B-BA7B-D9C806E333BB}#1.0#0"; "Buttons.ocx"
Begin VB.Form CommonMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   13020
   ClientLeft      =   2025
   ClientTop       =   120
   ClientWidth     =   28410
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "CommonMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   868
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1894
   WindowState     =   2  'Maximized
   Begin VB.Frame menuFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   1140
      Index           =   8
      Left            =   24000
      TabIndex        =   87
      Top             =   225
      Width           =   3915
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Τερματισμός εφαρμογής"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   52
         Left            =   300
         TabIndex        =   89
         Top             =   600
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Αλλαγή εταιρίας"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   51
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   88
         Top             =   225
         Width           =   3315
      End
   End
   Begin VB.Frame menuFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   9390
      Index           =   7
      Left            =   4125
      TabIndex        =   62
      Top             =   3075
      Width           =   4890
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ποσοστά Φ.Π.Α."
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   49
         Left            =   315
         TabIndex        =   86
         Top             =   3975
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Δημιουργία αρχείου γενικής λογιστικής"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   47
         Left            =   300
         TabIndex        =   85
         Top             =   8475
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Χρήστες"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   46
         Left            =   315
         TabIndex        =   84
         Top             =   7725
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ελεγχος αρχείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   48
         Left            =   300
         TabIndex        =   83
         Top             =   8850
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Τρόποι πληρωμής"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   44
         Left            =   315
         TabIndex        =   82
         Top             =   6975
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Χαρακτηρισμοί επιβαινόντων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   45
         Left            =   315
         TabIndex        =   81
         Top             =   7350
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Τιμοκατάλογοι εκδρομών λεωφορείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   42
         Left            =   315
         TabIndex        =   80
         Top             =   6225
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Τράπεζες"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   43
         Left            =   315
         TabIndex        =   79
         Top             =   6600
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Τιμοκατάλογοι εκδρομών πλοίων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   41
         Left            =   315
         TabIndex        =   78
         Top             =   5850
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Σημεία παραλαβής επιβατών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   39
         Left            =   315
         TabIndex        =   77
         Top             =   5100
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Σύνδεση προορισμών με δρομολόγια λεωφορείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   40
         Left            =   315
         TabIndex        =   76
         Top             =   5475
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Προορισμοί πλοίων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   37
         Left            =   315
         TabIndex        =   75
         Top             =   4350
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Προορισμοί λεωφορείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   38
         Left            =   315
         TabIndex        =   74
         Top             =   4725
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Δρομολόγια πλοίων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   31
         Left            =   315
         TabIndex        =   73
         Top             =   1725
         Width           =   4290
      End
      Begin VB.Label mnuHeader 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Πίνακες"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   21
         Left            =   300
         TabIndex        =   72
         Top             =   1350
         Width           =   4290
      End
      Begin VB.Label mnuHeader 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Παραμετροποίηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   20
         Left            =   300
         TabIndex        =   71
         Top             =   225
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Γενικές παράμετροι"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   29
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   70
         Top             =   600
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Εκτυπωτές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   30
         Left            =   300
         TabIndex        =   69
         Top             =   975
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Οροι πληρωμής"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   35
         Left            =   315
         MousePointer    =   2  'Cross
         TabIndex        =   68
         Top             =   3225
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Δρομολόγια λεωφορείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   32
         Left            =   315
         TabIndex        =   67
         Top             =   2100
         Width           =   4290
      End
      Begin VB.Label mnuHeader 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Εργασίες"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   18
         Left            =   300
         TabIndex        =   66
         Top             =   8100
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κατηγορίες εξόδων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   33
         Left            =   315
         TabIndex        =   65
         Top             =   2475
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Οικονομικές υπηρεσίες"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   34
         Left            =   315
         TabIndex        =   64
         Top             =   2850
         Width           =   4290
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Πλοία"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   36
         Left            =   315
         MousePointer    =   2  'Cross
         TabIndex        =   63
         Top             =   3600
         Width           =   4290
      End
   End
   Begin VB.Frame menuFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2640
      Index           =   6
      Left            =   20025
      TabIndex        =   55
      Top             =   225
      Width           =   3915
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   27
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   61
         Top             =   1725
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κατάσταση επιβατών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   28
         Left            =   300
         TabIndex        =   60
         Top             =   2100
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Γρήγορη καταχώρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   16
         Left            =   300
         TabIndex        =   59
         Top             =   225
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   25
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   58
         Top             =   600
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Κανονική καταχώρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   14
         Left            =   300
         TabIndex        =   57
         Top             =   1350
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κατάσταση επιβατών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   26
         Left            =   300
         TabIndex        =   56
         Top             =   975
         Width           =   3315
      End
   End
   Begin VB.Frame menuFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2640
      Index           =   5
      Left            =   16050
      TabIndex        =   48
      Top             =   225
      Width           =   3915
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κατάσταση με βάση τις πωλήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   23
         Left            =   300
         TabIndex        =   54
         Top             =   1725
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κατάσταση επιβατών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   22
         Left            =   300
         TabIndex        =   53
         Top             =   1350
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Αναφορές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   17
         Left            =   300
         TabIndex        =   52
         Top             =   975
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κατάσταση με βάση τους επιβαίνοντες"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   24
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   51
         Top             =   2100
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   21
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   50
         Top             =   600
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Διαχείρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   15
         Left            =   300
         TabIndex        =   49
         Top             =   225
         Width           =   3315
      End
   End
   Begin VB.Frame menuFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3765
      Index           =   4
      Left            =   12075
      TabIndex        =   38
      Top             =   225
      Width           =   3915
      Begin VB.Label mnuHeader 
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   13
         Left            =   300
         TabIndex        =   47
         Top             =   975
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Διαχείρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   12
         Left            =   300
         TabIndex        =   46
         Top             =   225
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Αρχείο"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   16
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   45
         Top             =   600
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Διαχείρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   17
         Left            =   300
         TabIndex        =   44
         Top             =   1350
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Τύποι παραστατικών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   11
         Left            =   300
         TabIndex        =   43
         Top             =   2850
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Διαχείρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   20
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   42
         Top             =   3225
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Αναφορές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   10
         Left            =   300
         TabIndex        =   41
         Top             =   1725
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Καρτέλα"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   18
         Left            =   300
         TabIndex        =   40
         Top             =   2100
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ισοζύγιο"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   19
         Left            =   300
         TabIndex        =   39
         Top             =   2475
         Width           =   3315
      End
   End
   Begin VB.Frame menuFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   4515
      Index           =   3
      Left            =   8100
      TabIndex        =   26
      Top             =   225
      Width           =   3915
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Εκδρομών λεωφορείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   15
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   37
         Top             =   3975
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ισοζύγιο"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   13
         Left            =   300
         TabIndex        =   36
         Top             =   2850
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Καρτέλα"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   12
         Left            =   300
         TabIndex        =   35
         Top             =   2475
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Αναφορές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   9
         Left            =   300
         TabIndex        =   34
         Top             =   2100
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Εκδρομές λεωφορείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   11
         Left            =   300
         TabIndex        =   33
         Top             =   1725
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Εκδρομών πλοίων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   14
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   32
         Top             =   3600
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Τύποι παραστατικών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   300
         TabIndex        =   31
         Top             =   3225
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Εκδρομές πλοίων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   10
         Left            =   300
         TabIndex        =   30
         Top             =   1350
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Αρχείο"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   9
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   29
         Top             =   600
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Διαχείρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   7
         Left            =   300
         TabIndex        =   28
         Top             =   225
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   300
         TabIndex        =   27
         Top             =   975
         Width           =   3315
      End
   End
   Begin VB.Frame menuFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2640
      Index           =   2
      Left            =   4125
      TabIndex        =   19
      Top             =   225
      Width           =   3915
      Begin VB.Label mnuHeader 
         Caption         =   "Αναφορές"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   300
         TabIndex        =   25
         Top             =   975
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Διαχείρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   300
         TabIndex        =   24
         Top             =   225
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   23
         Top             =   600
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ημερολόγιο"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   7
         Left            =   300
         TabIndex        =   22
         Top             =   1350
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         Caption         =   "Τύποι παραστατικών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   21
         Top             =   1725
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Διαχείρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   20
         Top             =   2100
         Width           =   3315
      End
   End
   Begin VB.Frame menuFrame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3390
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   225
      Width           =   3915
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Διαχείρηση"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   18
         Top             =   2850
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   17
         Top             =   1725
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ημερολόγιο"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   300
         TabIndex        =   16
         Top             =   2100
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         AutoSize        =   -1  'True
         Caption         =   "Τύποι παραστατικών"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   300
         TabIndex        =   15
         Top             =   2475
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ημερολόγιο"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   14
         Top             =   975
         Width           =   3315
      End
      Begin VB.Label menuOption 
         BackColor       =   &H0080C0FF&
         Caption         =   "Κινήσεις"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   300
         MousePointer    =   2  'Cross
         TabIndex        =   13
         Top             =   600
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         AutoSize        =   -1  'True
         Caption         =   "Εκδρομές λεωφορείων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   12
         Top             =   1350
         Width           =   3315
      End
      Begin VB.Label mnuHeader 
         AutoSize        =   -1  'True
         Caption         =   "Εκδρομές πλοίων"
         BeginProperty Font 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   22
         Left            =   300
         TabIndex        =   11
         Top             =   225
         Width           =   3315
      End
   End
   Begin VB.Frame frmNavigation 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   10500
      TabIndex        =   2
      Top             =   5250
      Width           =   12840
      Begin GurhanButtonOCX.GurhanButton cmdMenu 
         Height          =   840
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1482
         Caption         =   "Εσοδα"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   2
         BackColor       =   12632064
         ForeColor       =   0
         BEVEL           =   0
         BEVELDEPTH      =   0
      End
      Begin GurhanButtonOCX.GurhanButton cmdMenu 
         Height          =   840
         Index           =   2
         Left            =   1725
         TabIndex        =   4
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1482
         Caption         =   "Εξοδα"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   2
         BackColor       =   12632064
         ForeColor       =   0
         BEVEL           =   0
         BEVELDEPTH      =   0
      End
      Begin GurhanButtonOCX.GurhanButton cmdMenu 
         Height          =   840
         Index           =   3
         Left            =   3300
         TabIndex        =   5
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1482
         Caption         =   "Χρεώστες"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   2
         BackColor       =   12632064
         ForeColor       =   0
         BEVEL           =   0
         BEVELDEPTH      =   0
      End
      Begin GurhanButtonOCX.GurhanButton cmdMenu 
         Height          =   840
         Index           =   4
         Left            =   4875
         TabIndex        =   6
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1482
         Caption         =   "Πιστωτές"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   2
         BackColor       =   12632064
         ForeColor       =   0
         BEVEL           =   0
         BEVELDEPTH      =   0
      End
      Begin GurhanButtonOCX.GurhanButton cmdMenu 
         Height          =   840
         Index           =   5
         Left            =   6450
         TabIndex        =   7
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1482
         Caption         =   "Επιβαίνοντες πλοίων"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   2
         BackColor       =   12632064
         ForeColor       =   0
         BEVEL           =   0
         BEVELDEPTH      =   0
      End
      Begin GurhanButtonOCX.GurhanButton cmdMenu 
         Height          =   840
         Index           =   6
         Left            =   8025
         TabIndex        =   8
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1482
         Caption         =   "Επιβαίνοντες λεωφορείων"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   2
         BackColor       =   12632064
         ForeColor       =   0
         BEVEL           =   0
         BEVELDEPTH      =   0
      End
      Begin GurhanButtonOCX.GurhanButton cmdMenu 
         Height          =   840
         Index           =   7
         Left            =   9600
         TabIndex        =   9
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1482
         Caption         =   "Βοηθητικά"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   2
         XPColor_Hover   =   8421376
         BackColor       =   8421376
         ForeColor       =   0
         BEVEL           =   0
         BEVELDEPTH      =   0
      End
      Begin GurhanButtonOCX.GurhanButton cmdMenu 
         Height          =   840
         Index           =   8
         Left            =   11175
         TabIndex        =   10
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1482
         Caption         =   "Εξοδος"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ubuntu Condensed"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   2
         BackColor       =   8421631
         ForeColor       =   0
         BEVEL           =   0
         BEVELDEPTH      =   0
      End
   End
   Begin VB.Image imgImage 
      Appearance      =   0  'Flat
      Height          =   2400
      Left            =   17400
      Picture         =   "CommonMain.frx":0ECA
      Top             =   2550
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   3150
      Top             =   2850
      Width           =   1215
   End
   Begin VB.Label lblCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Corfu Cruises"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   390
      Left            =   12825
      TabIndex        =   0
      Top             =   9600
      Width           =   3690
   End
End
Attribute VB_Name = "CommonMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim isFirstTime As Boolean

Private Function BuildMenu()

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    
    CommonMain.Tag = "True"
    
    'Κεντράρισμα πλοήγησης
    With frmNavigation
        .Left = (CommonMain.Width / 2) - (.Width / 2)
        .Top = CommonMain.Height - GetSetting(strApplicationName, "Settings", "Navigation Top")
    End With
    
    'Απόσταση των μενού από την πλοήγηση
    For intLoop = 1 To menuFrame.UBound
        menuFrame(intLoop).Top = frmNavigation.Top - menuFrame(intLoop).Height + 100
        menuFrame(intLoop).Left = frmNavigation.Left + cmdMenu(intLoop - 1).Left
    Next intLoop
    
    'Χρώματα επικεφαλίδων μενού
    For intLoop = 1 To mnuHeader.UBound
        mnuHeader(intLoop).BackColor = &HFFC0C0
        mnuHeader(intLoop).ForeColor = &H0&
    Next intLoop
    
    'Χρώματα επιλογών
    For intLoop = 1 To menuOption.UBound
        menuOption(intLoop).BackColor = menuFrame(1).BackColor
        menuOption(intLoop).Caption = Space(1) & menuOption(intLoop).Caption & Space(1)
        menuOption(intLoop).MousePointer = 0
    Next intLoop
    
    Exit Function
    
ErrTrap:
    If Err.Number = 340 Then Resume Next

End Function

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
End Function



Private Function HideAllMenus()

    Dim intLoop As Integer
    
    For intLoop = 1 To menuFrame.UBound
        menuFrame(intLoop).Visible = False
    Next intLoop

End Function

Private Function PositionImage()

    imgImage.Left = frmNavigation.Left + frmNavigation.Width - imgImage.Width - 1000
    imgImage.Top = frmNavigation.Top - imgImage.Height - 200

End Function

Private Function PositionCompanyLabel()
            
    With lblCompany
        .BackColor = CommonMain.BackColor
        .Left = frmNavigation.Left + 1000
        .Top = frmNavigation.Top - lblCompany.Height - 500
    End With

End Function

Private Sub cmdMenu_Click(index As Integer)

    Dim intLoop As Integer
    
    HideAllMenus
    
    For intLoop = 1 To cmdMenu.UBound
        If index = intLoop Then
            menuFrame(intLoop).Left = cmdMenu(index).Left + frmNavigation.Left + 20
            menuFrame(intLoop).Visible = Not menuFrame(intLoop).Visible
        End If
    Next intLoop

End Sub

Private Sub cmdMenu_MouseIn(index As Integer, Shift As Integer)

    Select Case index
        Case Is < 7
            cmdMenu(index).BackColor = &H808000
            cmdMenu(index).ForeColor = &HFFFFFF
        Case Is = 7
            cmdMenu(index).BackColor = &H404000
            cmdMenu(index).ForeColor = &HFFFFFF
        Case Is = 8
            cmdMenu(index).BackColor = &HFF&
            cmdMenu(index).ForeColor = &HFFFFFF
        End Select
    
End Sub

Private Sub cmdMenu_MouseOut(index As Integer, Shift As Integer)

    Select Case index
        Case Is < 7
            cmdMenu(index).BackColor = &HC0C000
            cmdMenu(index).ForeColor = &H0
        Case Is = 7
            cmdMenu(index).BackColor = &H808000
            cmdMenu(index).ForeColor = &H0
        Case Is = 8
            cmdMenu(index).BackColor = &H8080FF
            cmdMenu(index).ForeColor = &H0
        End Select
        
End Sub

Private Sub Form_Activate()

    If isFirstTime Then
        isFirstTime = False
        HideAllMenus
        BuildMenu
        PositionImage
        PositionCompanyLabel
    End If

End Sub

Private Sub Form_Click()

    HideAllMenus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    CheckFunctionKeys KeyCode, Shift
    
    KeyCode = ResetKeyCode(KeyCode, Shift)

End Sub

Private Sub Form_Load()

    strImageDirectory = GetSetting(strApplicationName, "Path Names", "Image Directory")

    With CommonMain
        .Height = Screen.Height
        .Width = Screen.Width
        .ScaleHeight = .Height
        .ScaleWidth = .Width
        .BackColor = vbBlack
        .Refresh
    End With
    
    strReportsPathName = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Reports Path Name")
    strUnicodeFile = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Reports Path Name") & "UnicodeFile.txt"
    strAsciiFile = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Reports Path Name") & "AsciiFile.txt"
    
    isFirstTime = True

    blnAppIsRunning = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim obj As Object
    
    'Επιλογή κλεισίματος απο το μενού συστήματος, κλικ στο Χ ή ALT-F4
    If UnloadMode = 0 Then
        If CloseApp Then
            For Each obj In Forms
                Unload obj
            Next
            KillProcess strApplicationEXEName: End
        Else
            Cancel = 1
            Exit Sub
        End If
    End If
    
    'Επιλογή κλεισίματος από την επιλογή Εξοδος > Τερματισμός
    If UnloadMode = 1 Then
        KillProcess strApplicationEXEName
    End If

End Sub

Private Function CloseApp()

    CloseApp = False
    
    If MyMsgBox(2, strApplicationName, strStandardMessages(16), 2) Then
        CloseApp = True
    End If

End Function

Private Sub menuFrame_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ErrTrap
    
    Dim intLoop As Integer
    
    For intLoop = 1 To menuOption.UBound
        menuOption(intLoop).BackColor = menuFrame(index).BackColor
    Next
    
    Exit Sub
    
ErrTrap:
    If Err.Number = 340 Then Resume Next

End Sub


Private Sub menuOption_Click(index As Integer)

    Dim obj As Object

    HideAllMenus
    menuOption(index).BackColor = &HFFFFC0
    
    Select Case index
        'Εσοδα
        Case 1
            With InvoicesOut
                .lblTitle.Caption = "Εκδρομές πλοίων"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 2
            With InvoicesOutIndex
                .lblTitle.Caption = "Ημερολόγιο εκδρομών πλοίων"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 3
            With InvoicesOut
                .lblTitle.Caption = "Εκδρομές λεωφορείων"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .lblLabel(5).Visible = False
                .txtShipDescription.Visible = False
                .cmdIndex(3).Visible = False
                .cmdIndex(8).Visible = False
                .Tag = "True"
                .Show 1, Me
            End With
        Case 4
            With InvoicesOutIndex
                .lblTitle.Caption = "Ημερολόγιο εκδρομών λεωφορείων"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 5
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών εσόδων"
                .txtCodeMasterRefersTo.text = "2"
                .txtCodeSecondaryRefersTo.text = "0"
                .Tag = "True"
                .Show 1, Me
            End With
        'Εξοδα
        Case 6
            With InvoicesIn
                .Tag = "True"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtInvoiceSecondaryRefersTo.text = ""
                .Show 1, Me
            End With
        Case 7
            With InvoicesInIndex
                .txtInvoiceMasterRefersTo.text = "1"
                .txtInvoiceSecondaryRefersTo.text = ""
                .Tag = "True"
                .Show 1, Me
            End With
        Case 8
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών εξόδων"
                .txtCodeMasterRefersTo.text = "1"
                .txtCodeSecondaryRefersTo.text = "0"
                .Tag = "True"
                .Show 1, Me
            End With
        'Χρεώστες
        Case 9
            With Persons
                .Tag = "True"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtInvoiceMasterRefersTo.text = "2"
                .lblTitle.Caption = "Πελάτες"
                .Show 1, Me
            End With
        Case 10
            With PersonsTransactions
                .lblTitle.Caption = "Κινήσεις εκδρομών πλοίων"
                .txtInvoiceMasterRefersTo.text = "4"
                .txtInvoiceSecondaryRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 11
            With PersonsTransactions
                .lblTitle.Caption = "Κινήσεις εκδρομών λεωφορείων"
                .txtInvoiceMasterRefersTo.text = "4"
                .txtInvoiceSecondaryRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 12
            With PersonsLedger
                .lblTitle.Caption = "Καρτέλα χρεώστη"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 13
            With PersonsBalanceSheet
                .lblTitle.Caption = "Ισοζύγιο χρεωστών"
                .txtInvoiceMasterRefersTo.text = "2"
                .txtCustomersOrSuppliers.text = "Customers"
                .txtPaymentInOrPaymentOut.text = "PaymentIn"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 14
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών εκδρομών πλοίων"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "4"
                .txtCodeSecondaryRefersTo.text = "1"
                .Show 1, Me
            End With
        Case 15
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών εκδρομών λεωφορείων"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "4"
                .txtCodeSecondaryRefersTo.text = "2"
                .Show 1, Me
            End With
        'Πιστωτές
        Case 16
            With Persons
                .Tag = "True"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtInvoiceMasterRefersTo.text = "1"
                .lblTitle.Caption = "Πιστωτές"
                .Show 1, Me
            End With
        Case 17
            With PersonsTransactions
                .lblTitle.Caption = "Κινήσεις πιστωτών"
                .txtInvoiceMasterRefersTo.text = "3"
                .txtInvoiceSecondaryRefersTo.text = ""
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 18
            With PersonsLedger
                .lblTitle.Caption = "Καρτέλα πιστωτή"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 19
            With PersonsBalanceSheet
                .lblTitle.Caption = "Ισοζύγιο πιστωτών"
                .txtInvoiceMasterRefersTo.text = "1"
                .txtCustomersOrSuppliers.text = "Suppliers"
                .txtPaymentInOrPaymentOut.text = "PaymentOut"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 20
            With TablesCodes
                .lblTitle.Caption = "Τύποι παραστατικών πιστωτών"
                .Tag = "True"
                .txtCodeMasterRefersTo.text = "3"
                .Show 1, Me
            End With
        'Επιβάτες πλοίων
        Case 21
            With ShipsTransactions
                .Tag = "True"
                .Show 1, Me
            End With
        Case 22
            With ShipsRouteReport
                .Tag = "True"
                .Show 1, Me
            End With
        Case 23
            With ShipsStatistics
                .lblTitle.Caption = "Στατιστικά με βάση τις πωλήσεις"
                .txtTable.text = "Sales"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 24
            With ShipsStatistics
                .lblTitle.Caption = "Στατιστικά με βάση τους επιβαίνοντες"
                .txtTable.text = "Manifest"
                .Tag = "True"
                .Show 1, Me
            End With
        'Transfers
        Case 25
            With CoachesPickupsBrief
                .txtRefersTo.text = "2"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 26
            With CoachesReport
                .Tag = "True"
                .txtRefersTo.text = "2"
                .txtCallingForm = "MainMenu"
                .grdCoachesReport.Tag = "grdCoachesReportBrief"
                .Show 1, Me
            End With
        Case 27
            With CoachesPickupsStandard
                .txtRefersTo.text = "1"
                .Tag = "True"
                .Show 1, Me
            End With
        Case 28
            With CoachesReport
                .Tag = "True"
                .txtRefersTo.text = "1"
                .txtCallingForm = "MainMenu"
                .grdCoachesReport.Tag = "grdCoachesReportStandard"
                .Show 1, Me
            End With
        'Βοηθητικά
        Case 29
            With TablesSettings
                .Tag = "True"
                .Show 1, Me
            End With
        Case 30
            With TablesPrinters
                .Tag = "True"
                .Show 1, Me
            End With
        Case 31
            With TablesShipRoutes
                .Tag = "True"
                .Show 1, Me
            End With
        Case 32
            With TablesCoachRoutes
                .Tag = "True"
                .Show 1, Me
            End With
        Case 33
            With TablesExpenseCategories
                .Tag = "True"
                .Show 1, Me
            End With
        Case 34
            With TablesTaxOffices
                .Tag = "True"
                .Show 1, Me
            End With
        Case 35
            With TablesPaymentTerms
                .Tag = "True"
                .Show 1, Me
            End With
        Case 36
            With TablesShips
                .Tag = "True"
                .Show 1, Me
            End With
        Case 49
            With TablesVATPercents
                .Tag = "True"
                .Show 1, Me
            End With
        Case 37
            With TablesDestinations
                .Tag = "True"
                .lblTitle.Caption = "Προορισμοί πλοίων"
                .txtShowInList.text = "1"
                .Show 1, Me
            End With
        Case 38
            With TablesDestinations
                .Tag = "True"
                .lblTitle.Caption = "Προορισμοί λεωφορείων"
                .txtShowInList.text = "2"
                .Show 1, Me
            End With
        Case 39
            With TablesPickupPoints
                .Tag = "True"
                .Show 1, Me
            End With
        Case 40
            With UtilsJoinDestinationsWithRoutes
                .Tag = "True"
                .Show 1, Me
            End With
        Case 41
            With TablesPrices
                .Tag = "True"
                .lblTitle.Caption = "Τιμοκατάλογοι εκδρομών πλοίων"
                .txtShowInList.text = "1"
                .Show 1, Me
            End With
        Case 42
            With TablesPrices
                .Tag = "True"
                .lblTitle.Caption = "Τιμοκατάλογοι εκδρομών λεωφορείων"
                .txtShowInList.text = "2"
                .Show 1, Me
            End With
        Case 43
            With TablesBanks
                .Tag = "True"
                .Show 1, Me
            End With
        Case 44
            With TablesPaymentWays
                .Tag = "True"
                .Show 1, Me
            End With
        Case 45
            With TablesOccupantsDescriptions
                .Tag = "True"
                .Show 1, Me
            End With
        Case 46
            With TablesUsers
                .Tag = "True"
                .Show 1, Me
            End With
        Case 47
            With UtilsSalesExport
                .Tag = "True"
                .Show 1, Me
            End With
        Case 48
            With UtilsCheckFiles
                .Tag = "True"
                .Show 1, Me
            End With
        Case 51
            With CommonLogin
                .Tag = "True"
                .Visible = True
            End With
        Case 52
            If CloseApp Then
                For Each obj In Forms
                    Unload obj
                Next
                End
            End If
    End Select

End Sub

Private Sub menuOption_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X >= menuOption(index).Left - 1200 Then
        menuOption(index).BackColor = &H80C0FF: Exit Sub
    End If

End Sub


