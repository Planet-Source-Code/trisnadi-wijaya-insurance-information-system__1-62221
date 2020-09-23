VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCPolis 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Report Print"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "FrmCPolis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4605
   Begin VB.CommandButton CmdCetak 
      BackColor       =   &H00C0C000&
      Caption         =   "PRINT"
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
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pages"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24903681
         CurrentDate     =   38525
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24903681
         CurrentDate     =   38525
      End
      Begin VB.OptionButton OptHlmn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Period"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton OptHlmn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "All"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtNoPolis 
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptHlmn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ones"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Policy No."
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport CryRpt 
      Left            =   2040
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmCPolis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBatal_Click()
Unload Me
End Sub

Private Sub CmdCetak_Click()
CryRpt.ReportFileName = App.Path & "\Produksi.rpt"
CryRpt.WindowTitle = "Production Report"
If OptHlmn(0).Value = True Then
    CryRpt.Destination = crptToWindow
    CryRpt.Action = 1
ElseIf OptHlmn(1).Value = True Then
    CryRpt.SelectionFormula = "{Polis.No_Polis} = '" & TxtNoPolis.Text & "'"
    CryRpt.Destination = crptToWindow
    CryRpt.Action = 1
ElseIf OptHlmn(2).Value = True Then
    CryRpt.ReportTitle = "Period : " & DTPicker1.Value & " - " & DTPicker2.Value
    CryRpt.SelectionFormula = "{Polis.Tgl_Kontrak} >= #" & Format(DTPicker1.Value, "dd/mm/yyyy") & "# and {Polis.Tgl_Kontrak} <= # " & Format(DTPicker2.Value, "dd/mm/yyyy") & "#"
    CryRpt.Destination = crptToWindow
    CryRpt.Action = 1
End If
End Sub

Private Sub Form_Load()
Top = 100: Left = 100
OptHlmn(0).Value = True
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub

Private Sub OptHlmn_Click(Index As Integer)
If Index = 0 Then
    TxtNoPolis.Enabled = False
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
ElseIf Index = 1 Then
    TxtNoPolis.Enabled = True
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    TxtNoPolis.Text = ""
    TxtNoPolis.SetFocus
ElseIf Index = 2 Then
    TxtNoPolis.Enabled = False
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
End If
End Sub
