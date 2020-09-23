VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmKlaim 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insurance Claim Data"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   Icon            =   "FrmKlaim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   9495
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
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
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
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
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
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
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
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   855
   End
   Begin VB.CheckBox ChkOtomat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Automatic"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CmdCari 
      BackColor       =   &H008080FF&
      Caption         =   "Find"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.Frame FrmMaslahat 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   4680
      TabIndex        =   34
      Top             =   1560
      Width           =   4695
      Begin VB.Frame FrmTombol 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   120
         TabIndex        =   39
         Top             =   2400
         Width           =   4455
         Begin VB.CommandButton CmdHapus1 
            BackColor       =   &H0080C0FF&
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
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdBersih 
            BackColor       =   &H0080C0FF&
            Caption         =   "CLEAR"
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
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdTambah1 
            BackColor       =   &H0080C0FF&
            Caption         =   "ADD"
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
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DtG1 
         Height          =   1575
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2778
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Heirs List"
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
      Begin VB.TextBox TxtKTP 
         Height          =   285
         Left            =   840
         MaxLength       =   16
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtNama 
         Height          =   285
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "ID No."
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   2520
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame FrmKes 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   4680
      TabIndex        =   30
      Top             =   5040
      Width           =   4695
      Begin VB.TextBox TxtKlaim 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox CboBayar 
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Rp"
         Height          =   255
         Left            =   1680
         TabIndex        =   42
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Claim Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment for"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox TxtNoPolis 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   2040
      MaxLength       =   12
      TabIndex        =   18
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox CboJenis 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox TxtNoKlaim 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   1
      ToolTipText     =   "Press ENTER to continue"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame FrmMng 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   5040
      Width           =   4335
      Begin VB.TextBox TxtSebab 
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
      Begin MSMask.MaskEdBox MskTglMng 
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cause of Death"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Death"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp"
      Height          =   255
      Left            =   1800
      TabIndex        =   48
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label LblUP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2040
      TabIndex        =   47
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label LblNmProd 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2520
      TabIndex        =   46
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label LblKdProd 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   45
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Sum Insured"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Product"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label LblKdNasabah1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   41
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label LblKdNasabah 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   40
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HEIRS LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   35
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   375
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INSURANCE CLAIM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   8175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   8520
      Picture         =   "FrmKlaim.frx":0442
      Top             =   120
      Width           =   840
   End
   Begin VB.Label LblNmTertngg 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   29
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label LblNmPolHolder 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label LblTanggal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2040
      TabIndex        =   27
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Insured Name"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Policy Holder Name"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Policy No."
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Type"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Date"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Claim No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "FrmKlaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsKlaim As New ADODB.Recordset
Dim RsMaslahat As New ADODB.Recordset
Dim RsPolis As New ADODB.Recordset
Dim RsProduk As New ADODB.Recordset
Dim RsNasabah1 As New ADODB.Recordset
Dim RsNasabah2 As New ADODB.Recordset

Private Sub CboJenis_Click()
If CboJenis.ListIndex = 0 Then
    FrmMng.Enabled = True: FrmMaslahat.Enabled = True
    FrmKes.Enabled = False
ElseIf CboJenis.ListIndex = 1 Then
    FrmMng.Enabled = False: FrmMaslahat.Enabled = False
    FrmKes.Enabled = False
ElseIf CboJenis.ListIndex = 2 Then
    FrmMng.Enabled = False: FrmMaslahat.Enabled = False
    FrmKes.Enabled = True
End If
End Sub

Private Sub ChkOtomat_Click()
If ChkOtomat.Value = 1 Then
    Penomoran
    TxtNoKlaim.SetFocus
ElseIf ChkOtomat.Value = 0 Then
    TxtNoKlaim.Text = ""
End If
End Sub

'Auto numbering for Claim No.
Private Sub Penomoran()
    Dim Nom As String
    Dim NM As Integer
    SQL = "select * from Klaim order by No_Klaim"
    If RsKlaim.State > 0 Then RsKlaim.Close
    RsKlaim.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    If RsKlaim.RecordCount = 0 Then
        Nom = "00001"
    Else
        RsKlaim.MoveLast
        NM = Val(Trim(RsKlaim.Fields(0))) + 1
        Select Case NM
            Case Is < 10
                Nom = "0000" & NM
            Case Is < 100
                Nom = "000" & NM
            Case Is < 1000
                Nom = "00" & NM
            Case Is < 10000
                Nom = "0" & NM
            Case Else
                Nom = NM
            End Select
    End If
    TxtNoKlaim.Text = Nom
    RsKlaim.Close
End Sub

Private Sub CmdBatal_Click()
TxtNoKlaim.Text = ""
Bersih
Bersih1
ChkOtomat.Value = 0
NonAktif
CmdTambah.Enabled = False: CmdUbah.Enabled = False
CmdHapus.Enabled = False: CmdHapus1.Enabled = False
TxtNoKlaim.Enabled = True: TxtNoKlaim.SetFocus
ChkOtomat.Enabled = True
'Heirs List Refresh
RefGrid
LblTanggal.Caption = Date
End Sub

Private Sub CmdBersih_Click()
Bersih1
CmdHapus1.Enabled = False
TxtKTP.SetFocus
End Sub

Private Sub CmdCari_Click()
If Trim(CboJenis.Text) = "" Then
    MsgBox "Choose Claim Type first.", vbCritical, "Attention"
    CboJenis.SetFocus
Else
    If Left(CboJenis.Text, 1) = "3" Then
        FrmCari_PolKes.LblTujuan.Caption = Me.Name
        FrmCari_PolKes.Show 1
    Else
        FrmCari_Polis.LblTujuan.Caption = Me.Name
        FrmCari_Polis.Show 1
    End If
End If
End Sub

Private Sub CmdHapus_Click()
Tanya = MsgBox("Anda yakin akan menghapus data klaim ini ?", vbQuestion + vbYesNo, "Konfirmasi Hapus")
If Tanya = vbYes Then
SQL1 = "select * from Klaim where No_Klaim='" & TxtNoKlaim.Text & "'"
If RsKlaim.State > 0 Then RsKlaim.Close
RsKlaim.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
RsKlaim.Delete
RsKlaim.Close
'If Claim Type is Death Claim
If Left(CboJenis.Text, 1) = "1" Then
    'Delete Heirs List
    SQL3 = "select * from Maslahat where No_Klaim='" & TxtNoKlaim.Text & "'"
    If RsMaslahat.State > 0 Then RsMaslahat.Close
    RsMaslahat.Open SQL3, Cn, adOpenDynamic, adLockOptimistic
    For t = 1 To RsMaslahat.RecordCount
        RsMaslahat.Delete
        RsMaslahat.MoveNext
    Next t
    RsMaslahat.Close
End If
MsgBox "Claim data deleted.", vbInformation, "Information"
End If
TxtNoKlaim.Text = ""
Bersih
Bersih1
'Heirs List Refresh
RefGrid
CmdUbah.Enabled = False: CmdHapus.Enabled = False
TxtNoKlaim.Enabled = True: ChkOtomat.Enabled = True
ChkOtomat.Value = 0
TxtNoKlaim.SetFocus
End Sub

Private Sub CmdHapus1_Click()
SQL1 = "delete from Maslahat where No_Klaim='" & DtG1.Columns(0) & _
"'and No_KTP='" & DtG1.Columns(1) & "'"
If RsMaslahat.State > 0 Then RsMaslahat.Close
Set RsMaslahat = Cn.Execute(SQL1)
'Heirs List Refresh
RefGrid
Bersih1
TxtKTP.SetFocus
End Sub

Private Sub CmdTambah_Click()
SQL1 = "select * from Klaim order by No_Klaim"
If RsKlaim.State > 0 Then RsKlaim.Close
RsKlaim.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
With RsKlaim
    .AddNew
    .Fields("No_Klaim") = Trim(TxtNoKlaim.Text)
    .Fields("Tgl_Klaim") = CDate(LblTanggal.Caption)
    .Fields("Jns_Klaim") = Left(CboJenis.Text, 1)
    'If Claim Type is Death claim
    If (Left(CboJenis.Text, 1) = "1") Then
        .Fields("No_Polis") = Trim(TxtNoPolis.Text)
        .Fields("Tgl_Meninggal") = CDate(MskTglMng.Text)
        .Fields("Sebab_Meninggal") = TxtSebab.Text
        'Change Policy Status
        SQL4 = "select * from Polis where No_Polis='" & TxtNoPolis.Text & "'"
        If RsPolis.State > 0 Then RsPolis.Close
        RsPolis.Open SQL4, Cn, adOpenDynamic, adLockOptimistic
        RsPolis.Fields("Status") = "E"
        RsPolis.Update
        RsPolis.Close
    'If claim Type is Cash Value Claim
    ElseIf (Left(CboJenis.Text, 1) = "2") Then
        .Fields("No_Polis") = Trim(TxtNoPolis.Text)
        'Change Policy Status
        SQL4 = "select * from Polis where No_Polis='" & TxtNoPolis.Text & "'"
        If RsPolis.State > 0 Then RsPolis.Close
        RsPolis.Open SQL4, Cn, adOpenDynamic, adLockOptimistic
        RsPolis.Fields("Status") = "E"
        RsPolis.Update
        RsPolis.Close
    'If Claim Type is Health Claim
    ElseIf (Left(CboJenis.Text, 1) = "3") Then
        .Fields("No_PolKes") = Trim(TxtNoPolis.Text)
        .Fields("Klaim_Kes") = Left(CboBayar.Text, 1)
        .Fields("Besar_KlaimKes") = CCur(Trim(TxtKlaim.Text))
        'Change Policy Status
        SQL4 = "select * from PolKes where No_PolKes='" & TxtNoPolis.Text & "'"
        If RsPolis.State > 0 Then RsPolis.Close
        RsPolis.Open SQL4, Cn, adOpenDynamic, adLockOptimistic
        RsPolis.Fields("Status") = "E"
        RsPolis.Update
        RsPolis.Close
    End If
    .Update
    .Close
End With
MsgBox "Claim data saved.", vbInformation, "Information"
TxtNoKlaim.Text = ""
Bersih
Bersih1
'Heirs List Refresh
RefGrid
NonAktif
CmdTambah.Enabled = False
TxtNoKlaim.Enabled = True: ChkOtomat.Enabled = True
ChkOtomat.Value = 0
TxtNoKlaim.SetFocus
End Sub

Private Sub CmdTambah1_Click()
SQL = "select * from Maslahat where No_Klaim='" & TxtNoKlaim.Text & "'"
If RsMaslahat.State > 0 Then RsMaslahat.Close
RsMaslahat.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsMaslahat
    .AddNew
    .Fields("No_Klaim") = Trim(TxtNoKlaim.Text)
    .Fields("No_KTP") = Trim(TxtKTP.Text)
    .Fields("Nama") = TxtNama.Text
    .Update
    Set DtG1.DataSource = RsMaslahat
    End With
    Bersih1
    TxtKTP.SetFocus
End Sub

Private Sub CmdUbah_Click()
SQL1 = "select * from Klaim where No_Klaim='" & TxtNoKlaim.Text & "'"
If RsKlaim.State > 0 Then RsKlaim.Close
RsKlaim.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
With RsKlaim
    'If Claim Type is Death Claim
    If (Left(CboJenis.Text, 1) = "1") Then
        .Fields("No_Polis") = Trim(TxtNoPolis.Text)
        .Fields("Tgl_Meninggal") = CDate(MskTglMng.Text)
        .Fields("Sebab_Meninggal") = TxtSebab.Text
    'If Claim Type is Health Claim
    ElseIf (Left(CboJenis.Text, 1) = "3") Then
        .Fields("No_PolKes") = Trim(TxtNoPolis.Text)
        .Fields("Klaim_Kes") = Left(CboBayar.Text, 1)
        .Fields("Besar_KlaimKes") = CCur(Trim(TxtKlaim.Text))
    End If
    .Update
    .Close
End With
MsgBox "Claim data changed.", vbInformation, "Information"
TxtNoKlaim.Text = ""
Bersih
Bersih1
'Heirs List Refresh
RefGrid
CmdUbah.Enabled = False: CmdHapus.Enabled = False
TxtNoKlaim.Enabled = True: ChkOtomat.Enabled = True
TxtNoKlaim.SetFocus
End Sub

Private Sub DtG1_DblClick()
If RsMaslahat.RecordCount = 0 Then Exit Sub
TxtKTP.Text = DtG1.Columns(1)
TxtNama.Text = DtG1.Columns(2)
CmdHapus1.Enabled = True
End Sub

Private Sub Form_Load()
Left = 100: Top = 100
Koneksi
Cn.CursorLocation = adUseClient
LblTanggal.Caption = Date
TxtNoKlaim.Text = ""
Bersih
NonAktif
CmdTambah.Enabled = False: CmdUbah.Enabled = False
CmdHapus.Enabled = False: CmdHapus1.Enabled = False
'Fill Claim Type combobox
CboJenis.AddItem "1 - Death"
CboJenis.AddItem "2 - Cash Value"
CboJenis.AddItem "3 - Health"
'Fill Payment for combobox
CboBayar.AddItem "S - Surgery"
CboBayar.AddItem "M - Medication"
CboBayar.AddItem "T - Teeth"
CboBayar.AddItem "B - Born"
'Heirs List Refresh
RefGrid
End Sub

Private Sub RefGrid()
'Retrieve Heirs List
SQL = "select * from Maslahat where No_Klaim='" & TxtNoKlaim.Text & "'"
If RsMaslahat.State > 0 Then RsMaslahat.Close
Set RsMaslahat = Cn.Execute(SQL)
Set DtG1.DataSource = RsMaslahat
End Sub

Private Sub Bersih()
CboJenis.Text = "": TxtNoPolis.Text = ""
LblKdNasabah.Caption = "": LblKdNasabah1.Caption = ""
LblNmPolHolder.Caption = "": LblNmTertngg.Caption = ""
LblKdProd.Caption = "": LblNmProd.Caption = ""
LblUP.Caption = ""
MskTglMng.Mask = "": MskTglMng.Text = ""
MskTglMng.Mask = "##/##/####": TxtSebab.Text = ""
CboBayar.Text = "": TxtKlaim.Text = ""
End Sub

Private Sub Bersih1()
TxtNama.Text = "": TxtKTP.Text = ""
End Sub

Private Sub NonAktif()
CboJenis.Enabled = False: FrmMng.Enabled = False
FrmKes.Enabled = False: FrmMaslahat.Enabled = False
CmdCari.Enabled = False
End Sub

Private Sub Aktif()
CboJenis.Enabled = True: CmdCari.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsKlaim = Nothing
Set RsMaslahat = Nothing
Set RsPolis = Nothing
Set RsProduk = Nothing
Set RsNasabah1 = Nothing
Set RsNasabah2 = Nothing
End Sub

Private Sub TxtKlaim_Change()
If Trim(TxtKlaim.Text) <> "" Then
   dig$ = Mid(TxtKlaim.Text, Len(TxtKlaim.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtKlaim.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtKlaim.Text = digits$
          TxtKlaim.SelStart = Len(TxtKlaim.Text)
      End If
End If
End Sub

Private Sub TxtKTP_Change()
If Trim(TxtKTP.Text) <> "" Then
   dig$ = Mid(TxtKTP.Text, Len(TxtKTP.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtKTP.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtKTP.Text = digits$
          TxtKTP.SelStart = Len(TxtKTP.Text)
      End If
End If
End Sub

Private Sub TxtNoKlaim_Change()
If Trim(TxtNoKlaim.Text) <> "" Then
   dig$ = Mid(TxtNoKlaim.Text, Len(TxtNoKlaim.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtNoKlaim.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtNoKlaim.Text = digits$
          TxtNoKlaim.SelStart = Len(TxtNoKlaim.Text)
      End If
End If
End Sub

Private Sub TxtNoKlaim_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(TxtNoKlaim.Text) = "" Then Exit Sub
   If Len(Trim(TxtNoKlaim.Text)) < 5 Then
      MsgBox "Claim No. not allow less than 5 characters.", vbCritical, "Attention"
      NonAktif
      TxtNoKlaim.SetFocus
   Else
   If RsKlaim.State > 0 Then RsKlaim.Close
   RsKlaim.CursorLocation = adUseClient
   SQL1 = "select * from Klaim where No_Klaim='" & TxtNoKlaim.Text & "'"
   RsKlaim.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
   If RsKlaim.RecordCount = 0 Then
      Bersih
      Aktif
      TxtNoKlaim.Enabled = False: ChkOtomat.Enabled = False
      CmdTambah.Enabled = True
      CmdUbah.Enabled = False: CmdHapus.Enabled = False
      CboJenis.SetFocus
   ElseIf RsKlaim.RecordCount > 0 Then
      Bersih
      NonAktif
      'Retrieve claim data
      With RsKlaim
         LblTanggal.Caption = .Fields("Tgl_Klaim")
         CboJenis.ListIndex = .Fields("Jns_Klaim") - 1
        'If Health Insurance
        If Left(CboJenis.Text, 1) = "3" Then
           FrmMng.Enabled = False: FrmMaslahat.Enabled = False
           FrmKes.Enabled = True
           TxtNoPolis.Text = .Fields("No_PolKes")
           'Fill Payment for
           If .Fields("Klaim_Kes") = "S" Then
              CboBayar.Text = "S - Surgery"
           ElseIf .Fields("Klaim_Kes") = "M" Then
              CboBayar.Text = "M - Medication"
           ElseIf .Fields("Klaim_Kes") = "T" Then
              CboBayar.Text = "T - Teeth"
           ElseIf .Fields("Klaim_Kes") = "B" Then
              CboBayar.Text = "B - Born"
           End If
            TxtKlaim.Text = Format(.Fields("Besar_KlaimKes"), "#,#,#,#,0")
        Else
           'Retrieve Policy No.
           TxtNoPolis.Text = .Fields("No_Polis")
           'If Claim Type is Death Claim
           If Left(CboJenis.Text, 1) = "1" Then
              FrmMng.Enabled = True: FrmMaslahat.Enabled = True
              FrmKes.Enabled = False
              MskTglMng.Text = .Fields("Tgl_Meninggal")
              TxtSebab.Text = .Fields("Sebab_Meninggal")
              'Retrieve Heirs List
              RefGrid
              'If Claim Type is Cash Value Claim
           ElseIf Left(CboJenis.Text, 1) = "2" Then
              FrmMng.Enabled = False: FrmMaslahat.Enabled = False
              FrmKes.Enabled = False
           End If
      End If
      End With
      Tanya = MsgBox("Claim data already exist." + Chr(13) + "Do you want to edit ?", vbQuestion + vbYesNo, "Confirmation")
      If Tanya = vbYes Then
         TxtNoKlaim.Enabled = False: ChkOtomat.Enabled = False
         CboJenis.Enabled = False
         CmdTambah.Enabled = False
         CmdUbah.Enabled = True: CmdHapus.Enabled = True
      ElseIf Tanya = vbNo Then
         TxtNoKlaim.Text = ""
         Bersih
         ChkOtomat.Value = 0
         Bersih1
         NonAktif
         'Heirs List Refresh
         RefGrid
         LblTanggal.Caption = Date
         CmdTambah.Enabled = False
         CmdUbah.Enabled = False: CmdHapus.Enabled = False
      End If
    End If
   End If
End If
End Sub

Private Sub TxtNoPolis_Change()
'If Health Insurance
If Left(CboJenis.Text, 1) = "3" Then
   'Retrieve Health Policy data
   If RsPolis.State > 0 Then RsPolis.Close
   RsPolis.CursorLocation = adUseClient
   SQL1 = "select * from PolKes where No_PolKes='" & TxtNoPolis.Text & "'"
   RsPolis.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
   If RsPolis.RecordCount = 0 Then
      MsgBox "Policy data not found.", vbCritical, "Attention"
      Exit Sub
   Else
      LblKdProd.Caption = "": LblNmProd.Caption = ""
      LblUP.Caption = ""
      With RsPolis
        LblKdNasabah.Caption = .Fields("No_PolHolder")
        If RsNasabah2.State > 0 Then RsNasabah2.Close
        RsNasabah2.CursorLocation = adUseClient
        SQL2 = "select Nama from Nasabah2 where No_PolHolder='" & LblKdNasabah.Caption & "'"
        RsNasabah2.Open SQL2, Cn, adOpenDynamic, adLockOptimistic
        If RsNasabah2.RecordCount > 0 Then
            LblNmPolHolder.Caption = RsNasabah2.Fields("Nama")
        End If
        RsNasabah2.Close
      End With
   RsPolis.Close
   End If
ElseIf (Left(CboJenis.Text, 1) = "1") Or (Left(CboJenis.Text, 1) = "2") Then
   If RsPolis.State > 0 Then RsPolis.Close
   RsPolis.CursorLocation = adUseClient
   SQL3 = "select * from Polis where No_Polis='" & TxtNoPolis.Text & "'"
   RsPolis.Open SQL3, Cn, adOpenDynamic, adLockOptimistic
   If RsPolis.RecordCount = 0 Then
      MsgBox "Policy data not found.", vbCritical, "Attention"
      Exit Sub
   Else
      With RsPolis
        LblKdNasabah.Caption = .Fields("No_PolHolder")
        LblKdProd.Caption = .Fields("Kd_Prod")
        LblUP.Caption = Format(.Fields("UP"), "#,#,#,#")
        'Retrieve PolicyHolder name
        If RsNasabah2.State > 0 Then RsNasabah2.Close
        RsNasabah2.CursorLocation = adUseClient
        SQL4 = "select Nama from Nasabah2 where No_PolHolder='" & LblKdNasabah.Caption & "'"
        RsNasabah2.Open SQL4, Cn, adOpenDynamic, adLockOptimistic
        If RsNasabah2.RecordCount > 0 Then
           LblNmPolHolder.Caption = RsNasabah2.Fields("Nama")
        End If
        RsNasabah2.Close
        If .Fields("No_Tertngg") = "00000" Then
           LblKdNasabah1.Caption = .Fields("No_Tertngg")
           LblNmTertngg.Caption = LblNmPolHolder.Caption
        Else
           LblKdNasabah1.Caption = .Fields("No_Tertngg")
           'Retrieve Insured name
           If RsNasabah1.State > 0 Then RsNasabah1.Close
           RsNasabah1.CursorLocation = adUseClient
           SQL5 = "select Nama from Nasabah1 where No_Tertngg='" & LblKdNasabah1.Caption & "'"
           RsNasabah1.Open SQL5, Cn, adOpenDynamic, adLockOptimistic
           If RsNasabah1.RecordCount > 0 Then
              LblNmTertngg.Caption = RsNasabah1.Fields("Nama")
           End If
           RsNasabah1.Close
        End If
        End With
        RsPolis.Close
        'Retrieve Insurance Product name
        If RsProduk.State > 0 Then RsProduk.Close
        RsProduk.CursorLocation = adUseClient
        SQL6 = "select * from Produk where Kd_Prod='" & LblKdProd.Caption & "'"
        RsProduk.Open SQL6, Cn, adOpenDynamic, adLockOptimistic
        If RsProduk.RecordCount > 0 Then
            LblNmProd.Caption = RsProduk.Fields("Deskripsi")
        End If
        RsProduk.Close
   End If
End If
End Sub
