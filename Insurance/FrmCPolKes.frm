VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCPolKes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Data Polis Kesehatan"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "FrmCPolKes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4485
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Halaman"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton OptHlmn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tertentu"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtNoPolis 
         Height          =   285
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptHlmn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Semua"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polis Kesehatan"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdBatal 
      BackColor       =   &H00C0C000&
      Caption         =   "&BATAL"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CmdCetak 
      BackColor       =   &H00C0C000&
      Caption         =   "&CETAK"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin Crystal.CrystalReport CryRpt 
      Left            =   2040
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmCPolKes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBatal_Click()
Unload Me
End Sub

Private Sub CmdCetak_Click()
CryRpt.ReportFileName = App.Path & "\PolKes.rpt"
If OptHlmn(0).Value = True Then
    CryRpt.Destination = crptToWindow
    CryRpt.Action = 1
ElseIf OptHlmn(1).Value = True Then
    CryRpt.SelectionFormula = "{PolKes.No_PolKes} = '" & TxtNoPolis.Text & "'"
    CryRpt.Destination = crptToWindow
    CryRpt.Action = 1
End If
End Sub

Private Sub Form_Load()
Top = 100: Left = 100
OptHlmn(0).Value = True
End Sub

Private Sub OptHlmn_Click(Index As Integer)
If Index = 0 Then
    TxtNoPolis.Enabled = False
ElseIf Index = 1 Then
    TxtNoPolis.Enabled = True
    TxtNoPolis.Text = ""
    TxtNoPolis.SetFocus
End If
End Sub

