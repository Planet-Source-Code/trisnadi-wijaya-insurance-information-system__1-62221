VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCari_Premi 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Premium Payment Data"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "FrmCari_Premi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search by"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox CboCari 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton CmdCari 
         BackColor       =   &H0080C0FF&
         Caption         =   "FIND"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TxtCari 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   720
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FIND"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin MSDataGridLib.DataGrid DtGridPolis 
      Bindings        =   "FrmCari_Premi.frx":0442
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   12648447
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Premium Payment Data"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCari_Premi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPremi As New ADODB.Recordset

Private Sub CmdCari_Click()
If CboCari.Text = "" Then
    SQL = "select * from Byr_Premi order by No_Kw"
    Set RsPremi = Cn.Execute(SQL)
    Set DtGridPolis.DataSource = RsPremi
End If
If (CboCari.ListIndex = 0) And (Trim(TxtCari.Text) <> "") Then
    SQL = "select * from Byr_Premi where No_Kw='" & TxtCari.Text & "'"
    Set RsPremi = Cn.Execute(SQL)
    Set DtGridPolis.DataSource = RsPremi
ElseIf (CboCari.ListIndex = 1) And (Trim(TxtCari.Text) <> "") Then
    SQL = "select * from Byr_Premi where No_Polis='" & TxtCari.Text & "'"
    Set RsPremi = Cn.Execute(SQL)
    Set DtGridPolis.DataSource = RsPremi
End If
End Sub

Private Sub Form_Activate()
TxtCari.SetFocus
End Sub

Private Sub Form_Load()
'Fill Search by combobox
CboCari.AddItem "Receipt No."
CboCari.AddItem "Policy No."
TxtCari.Text = ""
Koneksi
Cn.CursorLocation = adUseClient
SQL = "select * from Byr_Premi order by No_Kw"
Set RsPremi = Cn.Execute(SQL)
Set DtGridPolis.DataSource = RsPremi
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsPremi = Nothing
End Sub

Private Sub TxtCari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdCari_Click
End If
End Sub

