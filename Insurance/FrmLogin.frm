VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Login"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdKeluar 
      BackColor       =   &H00C0C000&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton CmdLogin 
      BackColor       =   &H00C0C000&
      Caption         =   "&LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox TxtKunci 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Press ENTER to continue"
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox TxtUser 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "Press ENTER to continue"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALLIANZ INSURANCE SYSTEM LOGIN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   3840
      Picture         =   "FrmLogin.frx":0442
      Top             =   120
      Width           =   840
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsLogin As New ADODB.Recordset

Private Sub CmdLogin_Click()
SQL = "select * from Login where Username='" & Trim(TxtUser.Text) & "'"
RsLogin.Open SQL, Cn, adOpenDynamic, adLockOptimistic
With RsLogin
    If .Fields("Status") = "ADM" Then
        FrmUtama.mn6.Visible = True
    ElseIf .Fields("Status") = "USR" Then
        FrmUtama.mn6.Visible = False
    ElseIf .Fields("Status") = "EXC" Then
        FrmUtama.mn6.Visible = False
    End If
End With
    Unload Me
    FrmSelamat.Show
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    TxtUser.SetFocus
End Sub

Private Sub Form_Load()
    CmdLogin.Enabled = False
    TxtKunci.Enabled = False
    Koneksi
    RsLogin.CursorLocation = adUseClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsLogin = Nothing
End Sub

Private Sub TxtKunci_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtKunci.Text) = "" Then
          MsgBox "Password not allow empty.", vbCritical, "Attention"
          TxtKunci.SetFocus
        Else
        SQL = "select * from Login where Username='" & Trim(TxtUser.Text) & "' and Password='" & Trim(TxtKunci.Text) & "'"
        If RsLogin.State > 0 Then RsLogin.Close
        RsLogin.Open SQL, Cn, adOpenDynamic, adLockOptimistic
        If RsLogin.RecordCount > 0 Then
            CmdLogin.Enabled = True
            TxtUser.Enabled = False
            TxtKunci.Enabled = False
            RsLogin.Close
            CmdLogin.SetFocus
        Else
            MsgBox "Sorry, Wrong password !" + Chr(13) + "Please retype again.", vbCritical, "Attention"
            TxtKunci.Text = ""
            TxtKunci.SetFocus
        End If
        End If
    End If
End Sub

Private Sub txtuser_Change()
    TxtUser.Text = LCase(TxtUser.Text)
    TxtUser.SelStart = Len(TxtUser.Text)
End Sub

Private Sub TxtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtUser.Text) = "" Then
          MsgBox "Username not allow empty.", vbCritical, "Attention"
          TxtUser.SetFocus
        Else
        SQL = "select * from Login where Username='" & Trim(TxtUser.Text) & "'"
        RsLogin.Open SQL, Cn, adOpenDynamic, adLockOptimistic
        If RsLogin.RecordCount > 0 Then
            TxtKunci.Enabled = True
            TxtKunci.SetFocus
        Else
            MsgBox "Username not registered !", vbCritical, "Attention"
            TxtUser.SetFocus
        End If
            RsLogin.Close
        End If
    End If
End Sub
