VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm FrmUtama 
   AutoShowChildren=   0   'False
   BackColor       =   &H00C0C0C0&
   Caption         =   "Insurance Information System Application Program"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8010
   Icon            =   "FrmUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4935
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Ready."
            TextSave        =   "Ready."
            Object.Tag             =   ""
            Object.ToolTipText     =   "Status: Ready."
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "25/08/2005"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Tanggal Sistem"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:46"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Waktu Sistem"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Indikator Capslock"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "NUM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Indikator Numlock"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "INS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Indikator Insert"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Text            =   "Copyright (c) 2005 Trisnadi Wijaya"
            TextSave        =   "Copyright (c) 2005 Trisnadi Wijaya"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Copyright (c) 2005 Trisnadi Wijaya"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mn1 
      Caption         =   "Master"
      Begin VB.Menu mn11 
         Caption         =   "Insured Data"
      End
      Begin VB.Menu mn13 
         Caption         =   "Search Insured Data"
      End
      Begin VB.Menu Garis11 
         Caption         =   "-"
      End
      Begin VB.Menu mn12 
         Caption         =   "Policy Holder Data"
      End
      Begin VB.Menu mn14 
         Caption         =   "Search Policy Holder Data"
      End
      Begin VB.Menu Garis12 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit Application"
      End
   End
   Begin VB.Menu mn2 
      Caption         =   "Policy"
      Begin VB.Menu mn21 
         Caption         =   "Insurance Policy Account Data"
      End
      Begin VB.Menu mn22 
         Caption         =   "Search Policy Account Data"
      End
   End
   Begin VB.Menu mn3 
      Caption         =   "Claim"
      Begin VB.Menu mn31 
         Caption         =   "Insurance Claim Data"
      End
      Begin VB.Menu mn32 
         Caption         =   "Search Insurance Claim Data"
      End
   End
   Begin VB.Menu mn4 
      Caption         =   "Premium"
      Begin VB.Menu mn41 
         Caption         =   "Premium Payment Data"
      End
      Begin VB.Menu mn42 
         Caption         =   "Search Premium Payment Data"
      End
   End
   Begin VB.Menu mn5 
      Caption         =   "Reports"
      Begin VB.Menu mn52 
         Caption         =   "Production Report"
      End
      Begin VB.Menu mn53 
         Caption         =   "Claim Report"
      End
      Begin VB.Menu mn54 
         Caption         =   "Premium Payment Report"
      End
   End
   Begin VB.Menu mn6 
      Caption         =   "Adiministrator"
      Begin VB.Menu mn61 
         Caption         =   "Authorization"
         Begin VB.Menu mn611 
            Caption         =   "User and Password Data"
         End
         Begin VB.Menu mn612 
            Caption         =   "Search User and Password Data"
         End
      End
      Begin VB.Menu mn62 
         Caption         =   "Insurance Product"
      End
   End
   Begin VB.Menu mn7 
      Caption         =   "Help"
      Begin VB.Menu mn71 
         Caption         =   "&Help"
      End
      Begin VB.Menu Garis61 
         Caption         =   "-"
      End
      Begin VB.Menu mn72 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Unload(Cancel As Integer)
Tanya = MsgBox("Are you sure want to exit ?", vbQuestion + vbYesNo, "Exit Confirmation")
If Tanya = vbYes Then
    Cancel = 0
ElseIf Tanya = vbNo Then
    Cn.Close
    Set Cn = Nothing
    Cancel = 1
End If
End Sub

Private Sub mn11_Click()
    FrmTertanggung.Show
End Sub

Private Sub mn12_Click()
    FrmPolHolder.Show
End Sub

Private Sub mn13_Click()
    FrmCari_Tertngg.Show 1
End Sub

Private Sub mn14_Click()
    FrmCari_PolHolder.Show 1
End Sub

Private Sub mn21_Click()
    FrmPolis.Show
End Sub

Private Sub mn22_Click()
    FrmCari_Polis.Show 1
End Sub

Private Sub mn23_Click()
    FrmPolKes.Show
End Sub

Private Sub mn24_Click()
    FrmCari_PolKes.Show 1
End Sub

Private Sub mn31_Click()
    FrmKlaim.Show
End Sub

Private Sub mn32_Click()
    FrmCari_Klaim.Show 1
End Sub

Private Sub mn41_Click()
    FrmPremi.Show
End Sub

Private Sub mn42_Click()
    FrmCari_Premi.Show 1
End Sub

Private Sub mn51_Click()
    FrmCNasabah.Show
End Sub

Private Sub mn52_Click()
    FrmCPolis.Show
End Sub

Private Sub mn53_Click()
    FrmCKlaim.Show
End Sub

Private Sub mn54_Click()
    FrmCPremi.Show
End Sub

Private Sub mn611_Click()
    FrmUser.Show
End Sub

Private Sub mn612_Click()
    FrmCari_User.Show 1
End Sub

Private Sub mn62_Click()
    FrmProduk.Show
End Sub

Private Sub mn72_Click()
    frmAbout.Show 1
End Sub

Private Sub mnExit_Click()
    Unload Me
End Sub
