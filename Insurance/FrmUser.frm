VERSION 5.00
Begin VB.Form FrmUser 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User and Password Data"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "FrmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4350
   Begin VB.CommandButton CmdBatal 
      BackColor       =   &H00C0C000&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton CmdHapus 
      BackColor       =   &H00C0C000&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton CmdUbah 
      BackColor       =   &H00C0C000&
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton CmdTambah 
      BackColor       =   &H00C0C000&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox CboStatus 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmUser.frx":0442
      Left            =   1920
      List            =   "FrmUser.frx":044F
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox TxtKunci 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Press ENTER to continue"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsUser As New ADODB.Recordset

Private Sub CmdBatal_Click()
TxtUser.Text = ""
Bersih
NonAktif
CmdTambah.Enabled = False
CmdUbah.Enabled = False
CmdHapus.Enabled = False
TxtUser.SetFocus
End Sub

Private Sub CmdHapus_Click()
SQL = "select * from Login where username='" & Trim(TxtUser.Text) & "'"
If RsUser.State > 0 Then RsUser.Close
RsUser.Open SQL, Cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Are you sure want to delete this user ?", vbQuestion + vbYesNo, "Konfirmasi Hapus")
If Tanya = vbYes Then
    RsUser.Delete
    MsgBox "User data deleted.", vbInformation, "Informasi"
    RsUser.Close
End If
Bersih
CmdUbah.Enabled = False
CmdHapus.Enabled = False
TxtUser.SetFocus
End Sub

Private Sub CmdTambah_Click()
    If RsUser.State > 0 Then RsUser.Close
    'Save data
    RsUser.Open "Login", Cn, adOpenDynamic, adLockOptimistic
    With RsUser
        .AddNew
        Simpan
        .Update
        .Close
    End With
    MsgBox "User data saved.", vbInformation, "Informasi"
    TxtUser.Text = ""
    Bersih
    CmdTambah.Enabled = False
    TxtUser.SetFocus
End Sub

Private Sub CmdUbah_Click()
    'Edit data
    SQL = "select * from Login where Username='" & Trim(TxtUser.Text) & "'"
    If RsUser.State > 0 Then RsUser.Close
    RsUser.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsUser
        Simpan
        .Update
        .Close
    End With
    MsgBox "User data changed.", vbInformation, "Informasi"
    TxtUser.Text = ""
    Bersih
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    TxtUser.SetFocus
End If
End Sub

Private Sub Form_Activate()
TxtUser.SetFocus
End Sub

Private Sub Form_Load()
Left = 600: Top = 600
TxtUser.Text = "": TxtKunci.Text = ""
NonAktif
CmdTambah.Enabled = False
CmdUbah.Enabled = False
CmdHapus.Enabled = False
Koneksi
RsUser.CursorLocation = adUseClient
End Sub

Private Sub Aktif()
TxtKunci.Enabled = True
CboStatus.Enabled = True
End Sub

Private Sub NonAktif()
TxtKunci.Enabled = False
CboStatus.Enabled = False
End Sub

Private Sub Bersih()
TxtKunci.Text = ""
CboStatus.Text = ""
End Sub

Private Sub Simpan()
'For saving data
With RsUser
    .Fields("Username") = Trim(TxtUser.Text)
    .Fields("Password") = Trim(TxtKunci.Text)
    If CboStatus.ListIndex = 0 Then
        .Fields("Status") = "USR"
    ElseIf CboStatus.ListIndex = 1 Then
        .Fields("Status") = "EXC"
    ElseIf CboStatus.ListIndex = 2 Then
        .Fields("Status") = "ADM"
    End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsUser = Nothing
End Sub

Private Sub TxtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Trim(TxtUser.Text) = "" Then Exit Sub
    SQL = "select * from Login where Username='" & Trim(TxtUser.Text) & "'"
    If RsUser.State > 0 Then RsUser.Close
    RsUser.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsUser
        If .RecordCount = 0 Then
            Aktif
            CmdTambah.Enabled = True
            TxtKunci.SetFocus
        Else
            'Retrieve data
            TxtKunci.Text = Trim(.Fields("Password"))
            If .Fields("Status") = "USR" Then
                CboStatus.ListIndex = 0
            ElseIf .Fields("Status") = "EXC" Then
                CboStatus.ListIndex = 1
            ElseIf .Fields("Status") = "ADM" Then
                CboStatus.ListIndex = 2
            End If
            Tanya = MsgBox("User data exist." + Chr(13) + "Do you want to edit this data ?", vbQuestion + vbYesNo, "Confirmation")
            If Tanya = vbYes Then
                Aktif
                CmdUbah.Enabled = True
                CmdHapus.Enabled = True
                TxtKunci.SetFocus
            ElseIf Tanya = vbNo Then
                Bersih
                TxtUser.Text = ""
                TxtUser.SetFocus
            End If
        End If
    End With
    RsUser.Close
End If
End Sub
