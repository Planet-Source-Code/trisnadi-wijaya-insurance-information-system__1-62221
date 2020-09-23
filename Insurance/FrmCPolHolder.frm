VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCPolHolder 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Data Pemegang Polis"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3975
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Data yang Ingin Dicetak"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton OptHlmn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Semua"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtNoTertngg 
         Height          =   285
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtAwal 
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TxtAkhir 
         Height          =   285
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton OptHlmn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tertentu"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptHlmn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleksi"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. PolHolder"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "dari No."
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "sampai"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
   End
   Begin Crystal.CrystalReport CryRpt 
      Left            =   1800
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmCPolHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
Unload Me
End Sub

Private Sub CmdCetak_Click()
If OptHlmn(0).Value = True Then
    CryRpt.Destination = crptToWindow
    CryRpt.Action = 1
ElseIf OptHlmn(1).Value = True Then
    CryRpt.SelectionFormula = "{Nasabah1.No_PolHolder} = '" & TxtNoTertngg.Text & "'"
    CryRpt.Destination = crptToWindow
    CryRpt.Action = 1
ElseIf OptHlmn(2).Value = True Then
    CryRpt.SelectionFormula = "{Nasabah1.No_PolHolder} IN '" & TxtAwal.Text & "'TO '" _
    & TxtAkhir.Text & "'"
    CryRpt.Destination = crptToWindow
    CryRpt.Action = 1
End If
End Sub

Private Sub Form_Load()
Top = 100: Left = 100
OptHlmn(0).Value = True
CryRpt.ReportFileName = App.Path & "\Nasabah2.rpt"
End Sub

Private Sub OptHlmn_Click(Index As Integer)
If Index = 0 Then
    TxtNoTertngg.Enabled = False
    TxtAwal.Enabled = False
    TxtAkhir.Enabled = False
ElseIf Index = 1 Then
    TxtNoTertngg.Enabled = True
    TxtNoTertngg.Text = ""
    TxtNoTertngg.SetFocus
    TxtAwal.Enabled = False
    TxtAkhir.Enabled = False
ElseIf Index = 2 Then
    TxtNoTertngg.Enabled = False
    TxtAwal.Enabled = True
    TxtAwal.Text = "": TxtAkhir.Text = ""
    TxtAwal.SetFocus
    TxtAkhir.Enabled = True
End If
End Sub


