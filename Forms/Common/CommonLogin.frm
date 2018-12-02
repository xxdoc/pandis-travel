VERSION 5.00
Object = "{158C2A77-1CCD-44C8-AF42-AA199C5DCC6C}#1.0#0"; "dcButton.ocx"
Object = "{FFE4AEB4-0D55-4004-ADF2-3C1C84D17A72}#1.0#0"; "userControls.ocx"
Begin VB.Form CommonLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Pandis Travel"
   ClientHeight    =   6690
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   9495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000011&
   Icon            =   "CommonLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin UserControls.newText txtPassword 
      Height          =   465
      Left            =   7350
      TabIndex        =   3
      Top             =   4725
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      Alignment       =   2
      ForeColor       =   4194304
      MaxLength       =   10
      PasswordChar    =   "*"
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
   Begin VB.ComboBox cboUsers 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   6375
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3825
      Width           =   2505
   End
   Begin VB.ComboBox cboCompanies 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Ubuntu Condensed"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   5250
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4275
      Width           =   3630
   End
   Begin Dacara_dcButton.dcButton cmdButton 
      Height          =   690
      Index           =   0
      Left            =   6075
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1217
      BackColor       =   12640511
      ButtonShape     =   3
      ButtonStyle     =   4
      Caption         =   "Εναρξη"
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
      MousePointer    =   2
      PicOpacity      =   0
   End
   Begin Dacara_dcButton.dcButton cmdButton 
      Height          =   690
      Index           =   1
      Left            =   7500
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5475
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
      ForeColor       =   0
      MousePointer    =   2
      PicOpacity      =   0
   End
   Begin VB.Label lblProgress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Πρόοδος εργασιών"
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
      Height          =   255
      Left            =   5475
      TabIndex        =   8
      Top             =   3375
      Width           =   3390
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός"
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
      Height          =   255
      Index           =   2
      Left            =   6675
      TabIndex        =   7
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Χρήστης"
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
      Height          =   255
      Index           =   0
      Left            =   5730
      TabIndex        =   5
      Top             =   3870
      Width           =   585
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Πλατφόρμα: Win32"
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
      Index           =   99
      Left            =   5475
      TabIndex        =   0
      Top             =   2775
      Width           =   3390
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Εταιρία"
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
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   6
      Top             =   4335
      Width           =   510
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "(c) John Sourvinos 1999-2017"
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
      Left            =   5475
      TabIndex        =   4
      Top             =   3075
      Width           =   3390
   End
   Begin VB.Image Image1 
      Height          =   6540
      Left            =   75
      OLEDropMode     =   1  'Manual
      Picture         =   "CommonLogin.frx":000C
      Stretch         =   -1  'True
      Top             =   75
      Width           =   9315
   End
End
Attribute VB_Name = "CommonLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CheckFunctionKeys(KeyCode, Shift)
    
    Dim CtrlDown
    
    CtrlDown = Shift + vbCtrlMask
    
    Select Case KeyCode
        Case vbKeyF10 And cmdButton(0).Enabled, vbKeyC And CtrlDown = 4 And cmdButton(0).Enabled
            cmdButton_Click 0
        Case vbKeyEscape And cmdButton(1).Enabled
            cmdButton_Click 1
    End Select

End Function

Private Function DisplayHelloMessage(blnDisplay)

    'Ελεγχω αν θα εμφανίσω
    If Not blnDisplay Then Exit Function
    
    'Locals
    Dim intDayOfWeek As Integer
    Dim strMessageHasBeenDisplayed As String
    Dim strWeekdayGender As String
    
    'Μηνυμα αστείου!
    intDayOfWeek = Val(GetSetting(appName:=strApplicationName, Section:="Misc", Key:="Day Of Week"))
    strMessageHasBeenDisplayed = GetSetting(appName:=strApplicationName, Section:="Misc", Key:="Message Has Been Displayed")
    
    'Επιλογή άρθρου ημέρας
    Select Case intDayOfWeek
        Case 6
            strWeekdayGender = " τo "
        Case Else
            strWeekdayGender = " την "
    End Select
    
    'Εμφανίζω
    If Weekday(Now) = intDayOfWeek And strMessageHasBeenDisplayed = "False" Then
        SaveSetting strApplicationName, "Misc", "MessageHasBeenDisplayed", "True"
       If MyMsgBox(2, strApplicationName, strAppMessages(11) & strWeekdayGender & WeekdayName(Weekday(Now)) & ";", 1) Then
       End If
    End If
    
    'Επαναφορά!
    If Weekday(Now) = intDayOfWeek + 1 And strMessageHasBeenDisplayed = "True" Then
        SaveSetting strApplicationName, "Misc", "MessageHasBeenDisplayed", "False"
    End If

End Function

Private Function LoadCompanies()
    
    On Error GoTo ErrTrap
    
    Dim strCompanies As String
    Dim strCompany As String
    Dim bytPosition As Byte
    Dim obj As Object
    
    cboCompanies.Clear
    
    'If strDatabaseType = "Access" Then
        strPathName = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Database Path Name")
        strCompanies = Dir(strPathName & "*.mdb")
        Do While strCompanies <> ""
            If strCompanies <> "Printers.mdb" And strCompanies <> "Users.mdb" Then
                strCompany = ""
                bytPosition = 1
                While Mid(strCompanies, bytPosition, 1) <> "."
                    strCompany = strCompany + Mid(strCompanies, bytPosition, 1)
                    bytPosition = bytPosition + 1
                Wend
                cboCompanies.AddItem strCompany
            End If
            strCompanies = Dir
        Loop
    'End If
    
    LoadCompanies = True
    
    Exit Function
    
ErrTrap:
    LoadCompanies = False
    DisplayErrorMessage True, Err.Description
    
End Function

Private Function Start()

    lblProgress.Caption = strStandardMessages(20)
    lblProgress.Refresh
    
    If App.PrevInstance Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(15), 1) Then
        End If
        CloseApp
        End
    End If
    
    strCompanyName = cboCompanies.text & ".mdb"
    strCurrentUser = cboUsers.text
    
    If Not IsCorrectPassword(cboUsers.text, txtPassword.text) Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(11), 1) Then
        End If
        ClearFields lblProgress
        Exit Function
    End If
    
    If OpenDataBase(strCompanyName) Then
        If LoadParameters Then
            With CommonMain
                .Caption = "Server: " & strPathName & " - Εταιρία: " & Left(strCompanyName, Len(strCompanyName) - 4) & " - Χρήστης: " & strCurrentUser
                .lblCompany.Caption = Left(strCompanyName, Len(strCompanyName) - 4)
                If Not .Visible Then .Show
                Unload Me
                CommonMain.Show
                DisplayHelloMessage False
            End With
        End If
    Else
        If MyMsgBox(4, strApplicationName, strStandardMessages(11), 1) Then
        End If
    End If

End Function

Private Function ValidateFields()

    ValidateFields = False
    
    'Χρήστες
    If cboUsers.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        cboUsers.SetFocus
        Exit Function
    End If
    
    'Εταιρία
    If cboCompanies.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        cboCompanies.SetFocus
        Exit Function
    End If
    
    'Κωδικός
    If txtPassword.text = "" Then
        If MyMsgBox(4, strApplicationName, strStandardMessages(1), 1) Then
        End If
        txtPassword.SetFocus
        Exit Function
    End If
    
    ValidateFields = True

End Function

Private Sub cboCompanies_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub cboUsers_KeyPress(KeyAscii As Integer)

    ValidateInput (KeyAscii)

End Sub

Private Sub cmdButton_Click(Index As Integer)

    Dim obj As Object
    
    Select Case Index
        Case 0
            If ValidateFields Then Start
        Case 1
            If blnAppIsRunning Then
                Unload Me
            Else
                If CloseApp Then
                    For Each obj In Forms
                        Unload obj
                    Next
                End If
            End If
    End Select
        
End Sub

Private Function CloseApp()

    CloseApp = False
    
    If MyMsgBox(2, strApplicationName, strStandardMessages(16), 2) Then
        CloseApp = True
    End If

End Function

Private Sub cmdButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    CheckForArrows (KeyCode)

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

Private Sub Form_Load()

    Me.Tag = "True"
    Me.Show
    lblCopyright.Caption = "Created by John Sourvinos 1999 - " & Year(Date)
    ClearFields lblProgress, cboUsers, cboCompanies, txtPassword
    strApplicationName = GetSetting(Mid(App.Path, 4, Len(App.Path)), "Settings", "Application Name")
    strApplicationEXEName = GetSetting(strApplicationName, "Settings", "Application EXE Name")
    LoadMessages
    If Not LoadCompanies Then Exit Sub
    If Not LoadUsers Then Exit Sub
    If GetSetting(appName:=strApplicationName, Section:="Settings", Key:="IsDevelopment") = "1" Then
        cboUsers.ListIndex = 0
        cboCompanies.ListIndex = 0
        txtPassword.text = "1701"
    End If

End Sub
    
Private Function LoadUsers()
    
    On Error GoTo ErrTrap
    
    Dim rsUsers As Recordset
    
    strPathName = GetSetting(appName:=strApplicationName, Section:="Path Names", Key:="Database Path Name")
    Set UsersDB = DBEngine.OpenDataBase(strPathName + "Users.mdb", False, False)
    
    Set rsUsers = UsersDB.OpenRecordset("Users")
    With rsUsers
        While Not .EOF
            cboUsers.AddItem !username
            .MoveNext
        Wend
    End With
    
    LoadUsers = True
    UsersDB.Close
    
    Exit Function
    
ErrTrap:
    LoadUsers = False
    DisplayErrorMessage True, Err.Description

End Function
