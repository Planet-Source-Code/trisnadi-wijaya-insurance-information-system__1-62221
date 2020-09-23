VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPolis 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insurance Policy Account Data"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   Icon            =   "FrmPolis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   10830
   Begin VB.Frame FrmBayar 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   6840
      TabIndex        =   57
      Top             =   4560
      Width           =   3855
      Begin VB.OptionButton OptTrans 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Transfer"
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptCC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Credit Card"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptCash 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cash/Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdCari 
      BackColor       =   &H008080FF&
      Caption         =   "Find"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox TxtUP 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      MaxLength       =   10
      TabIndex        =   11
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox TxtPremi 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      MaxLength       =   9
      TabIndex        =   10
      ToolTipText     =   "Tekan ENTER untuk tampilkan jumlah UP"
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox CboBayar 
      Height          =   315
      Left            =   7200
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ComboBox CboProduk 
      Height          =   315
      Left            =   7200
      TabIndex        =   8
      Top             =   2280
      Width           =   3375
   End
   Begin VB.ComboBox CboAgen 
      Height          =   315
      Left            =   6720
      TabIndex        =   4
      Top             =   840
      Width           =   3015
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
      Height          =   615
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5280
      Width           =   975
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5280
      Width           =   975
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Width           =   975
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox TxtMasa 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.Frame FrmPolHolder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Policy Holder"
      Height          =   2055
      Left            =   240
      TabIndex        =   28
      Top             =   1920
      Width           =   4455
      Begin VB.CommandButton CmdPolHolder 
         Caption         =   "...."
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
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox TxtPolHolder 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   51
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox CboRelasi 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label LblJKel2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   54
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   255
         Left            =   3480
         TabIndex        =   53
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   255
         Left            =   3360
         TabIndex        =   47
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label LblUmur2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   34
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label LblTglLhr2 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblNama2 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   32
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Policy Holder No."
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FrmTertngg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Insured"
      Height          =   1575
      Left            =   240
      TabIndex        =   21
      Top             =   4320
      Width           =   4455
      Begin VB.CommandButton CmdTertngg 
         Caption         =   "...."
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
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox TxtTertngg 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   52
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LblJKel1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   56
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   255
         Left            =   3480
         TabIndex        =   55
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   255
         Left            =   3360
         TabIndex        =   48
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label LblUmur1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label LblTglLhr1 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblNama1 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Insured No."
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   24707073
      CurrentDate     =   38465
   End
   Begin VB.TextBox TxtNoPolis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   1
      ToolTipText     =   "Press ENTER to continue"
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Premium Payment Method"
      Height          =   375
      Left            =   5040
      TabIndex        =   58
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp"
      Height          =   255
      Left            =   6840
      TabIndex        =   50
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp"
      Height          =   255
      Left            =   6840
      TabIndex        =   49
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label LblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
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
      Left            =   6720
      TabIndex        =   45
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Policy Status"
      Height          =   255
      Left            =   5040
      TabIndex        =   44
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Premium Payment Mode"
      Height          =   255
      Left            =   5040
      TabIndex        =   43
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Years"
      Height          =   255
      Left            =   7920
      TabIndex        =   42
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Period"
      Height          =   255
      Left            =   5040
      TabIndex        =   41
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Sum Assured"
      Height          =   255
      Left            =   5040
      TabIndex        =   40
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Premium Amount"
      Height          =   255
      Left            =   5040
      TabIndex        =   39
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Left            =   5040
      TabIndex        =   38
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INSURANCE PRODUCT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   37
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   375
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   9840
      Picture         =   "FrmPolis.frx":0442
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INSURANCE POLICY ACCOUNT"
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
      Left            =   240
      TabIndex        =   36
      Top             =   120
      Width           =   9255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      FillColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Code"
      Height          =   255
      Left            =   5040
      TabIndex        =   35
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contract Date"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Policy No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "FrmPolis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPolis As New ADODB.Recordset
Dim RsAgen As New ADODB.Recordset
Dim RsNasabah1 As New ADODB.Recordset
Dim RsNasabah2 As New ADODB.Recordset
Dim RsRelasi As New ADODB.Recordset
Dim RsProduk As New ADODB.Recordset

Private Sub CboRelasi_Click()
If Left(CboRelasi.Text, 2) = "IS" Then
    TxtTertngg.Text = "00000"
    LblNama1.Caption = LblNama2.Caption
    LblTglLhr1.Caption = LblTglLhr2.Caption
    FrmTertngg.Enabled = False
End If
End Sub

Private Sub CmdCari_Click()
FrmCari_Polis.LblTujuan.Caption = Me.Name
FrmCari_Polis.Show 1
End Sub

Private Sub CmdHapus_Click()
SQL = "select * from Polis where No_Polis='" & TxtNoPolis.Text & "'"
If RsPolis.State > 0 Then RsPolis.Close
RsPolis.Open SQL, Cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Are you sure want to delete this policy account ?", vbQuestion + vbYesNo, "Delete Confirmation")
If Tanya = vbYes Then
    RsPolis.Delete
    RsPolis.Close
    MsgBox "Policy data deleted.", vbInformation, "Information"
End If
TxtNoPolis.Text = ""
Bersih
NonAktif
CmdUbah.Enabled = False: CmdHapus.Enabled = False
TxtNoPolis.Enabled = True: CmdCari.Enabled = True
TxtNoPolis.SetFocus
End Sub

Private Sub CmdPolHolder_Click()
FrmCari_PolHolder.LblTujuan.Caption = Me.Name
FrmCari_PolHolder.Show 1
End Sub

Private Sub CmdTambah_Click()
    SQL = "select * from Polis order by No_Polis"
    If RsPolis.State > 0 Then RsPolis.Close
    RsPolis.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    'Save data
    With RsPolis
    .AddNew
        .Fields("No_Polis") = Trim(TxtNoPolis.Text)
        Simpan
        .Fields("Status") = Left(LblStatus.Caption, 1)
    .Update
    .Close
    End With
    MsgBox "Policy data saved.", vbInformation, "Information"
    TxtNoPolis.Text = ""
    Bersih
    NonAktif
    CmdTambah.Enabled = False
    TxtNoPolis.Enabled = True: CmdCari.Enabled = True
    TxtNoPolis.SetFocus
End Sub

Private Sub Simpan()
With RsPolis
    .Fields("Tgl_Kontrak") = DTP1.Value
    .Fields("No_PolHolder") = Trim(TxtPolHolder.Text)
    If Left(CboRelasi.Text, 2) = "IS" Then
        .Fields("No_Tertngg") = "00000"
    Else
        .Fields("No_Tertngg") = Trim(TxtTertngg.Text)
    End If
    .Fields("Relasi") = Left(CboRelasi.Text, 2)
    .Fields("Kd_Prod") = Left(CboProduk.Text, 3)
    .Fields("Lama") = Trim(TxtMasa.Text)
    .Fields("UP") = CCur(Trim(TxtUP.Text))
    .Fields("Premi") = CCur(Trim(TxtPremi.Text))
    .Fields("Periode_Byr") = Left(CboBayar.Text, 2)
    If OptCash.Value = True Then
        .Fields("Cara_Byr") = "T"
    ElseIf OptCC.Value = True Then
        .Fields("Cara_Byr") = "K"
    ElseIf OptTrans.Value = True Then
        .Fields("Cara_Byr") = "R"
    End If
    .Fields("Kd_Agen") = Left(CboAgen.Text, 8)
End With
End Sub

Private Sub CmdBatal_Click()
TxtNoPolis.Text = ""
Bersih
NonAktif
CmdTambah.Enabled = False
CmdUbah.Enabled = False: CmdHapus.Enabled = False
TxtNoPolis.Enabled = True: CmdCari.Enabled = True
TxtNoPolis.SetFocus
End Sub

Private Sub CmdTertngg_Click()
FrmCari_Tertngg.LblTujuan.Caption = Me.Name
FrmCari_Tertngg.Show 1
End Sub

Private Sub CmdUbah_Click()
    'Edit data
    SQL = "select * from Polis where No_Polis='" & TxtNoPolis.Text & "'"
    If RsPolis.State > 0 Then RsPolis.Close
    RsPolis.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsPolis
        Simpan
        .Update
        .Close
    End With
    MsgBox "Data Polis telah berhasil diubah dengan sukses.", vbInformation, "Informasi"
    TxtNoPolis.Text = ""
    Bersih
    NonAktif
    TxtNoPolis.Enabled = True: CmdCari.Enabled = True
    TxtNoPolis.SetFocus
End Sub

Private Sub Form_Load()
Left = 200: Top = 200
DTP1.Value = Date
Koneksi
'Fill Agent code combobox
SQL1 = "select Kd_Agen,Nama_Agen from Agen where Status='A' order by Kd_Agen"
If RsAgen.State > 0 Then RsAgen.Close
RsAgen.CursorLocation = adUseClient
RsAgen.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
RsAgen.MoveFirst
For t = 1 To RsAgen.RecordCount
    Isi = Trim(RsAgen.Fields("Kd_Agen")) & " - " & Trim(RsAgen.Fields("Nama_Agen"))
    CboAgen.AddItem Isi
    RsAgen.MoveNext
Next t
RsAgen.Close
'Fill Relation combobox
SQL3 = "select Kd_Relasi,Relasi from Hubungan order by Kd_Relasi"
If RsRelasi.State > 0 Then RsRelasi.Close
RsRelasi.CursorLocation = adUseClient
RsRelasi.Open SQL3, Cn, adOpenDynamic, adLockOptimistic
RsRelasi.MoveFirst
For W = 1 To RsRelasi.RecordCount
    Isi3 = Trim(RsRelasi.Fields("Kd_Relasi")) & " - " & Trim(RsRelasi.Fields("Relasi"))
    CboRelasi.AddItem Isi3
    RsRelasi.MoveNext
Next W
RsRelasi.Close
'Fill Insurance Product Code
SQL5 = "select Kd_Prod,Deskripsi from Produk order by Kd_Prod"
If RsProduk.State > 0 Then RsProduk.Close
RsProduk.CursorLocation = adUseClient
RsProduk.Open SQL5, Cn, adOpenDynamic, adLockOptimistic
RsProduk.MoveFirst
For H = 1 To RsProduk.RecordCount
    Isi5 = Trim(RsProduk.Fields("Kd_Prod")) & " - " & Trim(RsProduk.Fields("Deskripsi"))
    CboProduk.AddItem Isi5
    RsProduk.MoveNext
Next H
RsProduk.Close
'Fill Premium Payment Period combobox
CboBayar.AddItem "SA - Semi Anually"
CboBayar.AddItem "AN - Annually"
TxtNoPolis.Text = ""
Bersih
NonAktif
CmdTambah.Enabled = False
CmdUbah.Enabled = False
CmdHapus.Enabled = False
End Sub

Private Sub Aktif()
    DTP1.Enabled = True
    FrmTertngg.Enabled = True: FrmPolHolder.Enabled = True
    CboAgen.Enabled = True: CboProduk.Enabled = True
    TxtPremi.Enabled = True: TxtUP.Enabled = True
    TxtMasa.Enabled = True: CboBayar.Enabled = True
    FrmBayar.Enabled = True
End Sub

Private Sub NonAktif()
    DTP1.Enabled = False
    FrmTertngg.Enabled = False: FrmPolHolder.Enabled = False
    CboAgen.Enabled = False: CboProduk.Enabled = False
    TxtPremi.Enabled = False: TxtUP.Enabled = False
    TxtMasa.Enabled = False: CboBayar.Enabled = False
    FrmBayar.Enabled = False
End Sub

Private Sub Bersih()
    DTP1.Value = Date: LblStatus.Caption = "A - Active"
    CboAgen.Text = "": TxtTertngg.Text = ""
    LblNama1.Caption = "": LblUmur1.Caption = ""
    LblJKel1.Caption = "": LblJKel2.Caption = ""
    LblTglLhr1.Caption = "": CboRelasi.Text = ""
    TxtPolHolder.Text = "": LblNama2.Caption = ""
    LblUmur2.Caption = "": LblTglLhr2.Caption = ""
    CboProduk.Text = "": TxtPremi.Text = ""
    TxtUP.Text = "": CboBayar.Text = ""
    TxtMasa.Text = "": OptCash.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsPolis = Nothing
Set RsAgen = Nothing
Set RsNasabah1 = Nothing
Set RsNasabah2 = Nothing
Set RsRelasi = Nothing
Set RsProduk = Nothing
End Sub

Private Sub LblTglLhr1_Change()
If Trim(LblTglLhr1.Caption) = "" Then Exit Sub
'Count Insured Age
If (Month(Now) = Month(LblTglLhr1.Caption)) And (Day(Now) >= Day(LblTglLhr1.Caption)) Then
    LblUmur1.Caption = Year(Now) - Year(LblTglLhr1.Caption)
ElseIf Month(Now) > Month(LblTglLhr1.Caption) Then
    LblUmur1.Caption = Year(Now) - Year(LblTglLhr1.Caption)
Else
    LblUmur1.Caption = (Year(Now) - Year(LblTglLhr1.Caption)) - 1
End If
End Sub

Private Sub LblTglLhr2_Change()
If Trim(LblTglLhr2.Caption) = "" Then Exit Sub
'Count Policy Holder Age
If (Month(Now) = Month(LblTglLhr2.Caption)) And (Day(Now) >= Day(LblTglLhr2.Caption)) Then
    LblUmur2.Caption = Year(Now) - Year(LblTglLhr2.Caption)
ElseIf Month(Now) > Month(LblTglLhr2.Caption) Then
    LblUmur2.Caption = Year(Now) - Year(LblTglLhr2.Caption)
Else
    LblUmur2.Caption = (Year(Now) - Year(LblTglLhr2.Caption)) - 1
End If
End Sub

Private Sub TxtMasa_Change()
If Trim(TxtMasa.Text) <> "" Then
   dig$ = Mid(TxtMasa.Text, Len(TxtMasa.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtMasa.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtMasa.Text = digits$
          TxtMasa.SelStart = Len(TxtMasa.Text)
      End If
End If
End Sub

Private Sub TxtNoPolis_Change()
If Trim(TxtNoPolis.Text) <> "" Then
   dig$ = Mid(TxtNoPolis.Text, Len(TxtNoPolis.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtNoPolis.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtNoPolis.Text = digits$
          TxtNoPolis.SelStart = Len(TxtNoPolis.Text)
      End If
End If
End Sub

Private Sub TxtNoPolis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtNoPolis.Text) = "" Then Exit Sub
    SQL1 = "select * from Polis where No_Polis='" & TxtNoPolis.Text & "'"
    If RsPolis.State > 0 Then RsPolis.Close
    RsPolis.CursorLocation = adUseClient
    RsPolis.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
    If RsPolis.RecordCount = 0 Then
        Aktif
        Bersih
        TxtNoPolis.Enabled = False: CmdCari.Enabled = False
        CmdTambah.Enabled = True
        CmdUbah.Enabled = False: CmdHapus.Enabled = False
        CboAgen.SetFocus
    ElseIf RsPolis.RecordCount > 0 Then
        NonAktif
        'Retrieve data
        With RsPolis
            DTP1.Value = .Fields("Tgl_Kontrak")
            'Fill Policy Status description
            If .Fields("Status") = "A" Then
                LblStatus.Caption = "A - Active"
            ElseIf .Fields("Status") = "L" Then
                LblStatus.Caption = "L - Lapse"
            ElseIf .Fields("Status") = "E" Then
                LblStatus.Caption = "E - Expired"
            End If
            TxtPolHolder.Text = .Fields("No_PolHolder")
            If .Fields("Relasi") = "IS" Then
                TxtTertngg.Text = "00000"
            Else
                TxtTertngg.Text = .Fields("No_Tertngg")
            End If
            'Fill Relation description
            SQL4 = "select Relasi from Hubungan where Kd_Relasi='" & .Fields("Relasi") & "'"
            If RsRelasi.State > 0 Then RsRelasi.Close
            RsRelasi.Open SQL4, Cn, adOpenDynamic, adLockOptimistic
            Hubungan = .Fields("Relasi") & " - " & RsRelasi.Fields("Relasi")
            CboRelasi.Text = Hubungan
            RsRelasi.Close
            'Fill Agen Name
            SQL5 = "select Nama_Agen from Agen where Kd_Agen='" & .Fields("Kd_Agen") & "'"
            If RsAgen.State > 0 Then RsAgen.Close
            RsAgen.Open SQL5, Cn, adOpenDynamic, adLockOptimistic
            Nm_Agen = .Fields("Kd_Agen") & " - " & RsAgen.Fields("Nama_Agen")
            CboAgen.Text = Nm_Agen
            RsAgen.Close
            'Fill Product Name
            SQL6 = "select Deskripsi from Produk where Kd_Prod='" & .Fields("Kd_Prod") & "'"
            If RsProduk.State > 0 Then RsProduk.Close
            RsProduk.Open SQL6, Cn, adOpenDynamic, adLockOptimistic
            CboProduk.Text = .Fields("Kd_Prod") & " - " & RsProduk.Fields("Deskripsi")
            RsProduk.Close
            'Fill Premium Payment Period
            If .Fields("Periode_Byr") = "SA" Then
                CboBayar.ListIndex = 0
            ElseIf .Fields("Periode_Byr") = "AN" Then
                CboBayar.ListIndex = 1
            End If
            TxtPremi.Text = Format(.Fields("Premi"), "#,#,#,#,0")
            TxtUP.Text = Format(.Fields("UP"), "#,#,#,#,0")
            TxtMasa.Text = .Fields("Lama")
            If .Fields("Cara_Byr") = "T" Then
                OptCash.Value = True
            ElseIf .Fields("Cara_Byr") = "K" Then
                OptCC.Value = True
            ElseIf .Fields("Cara_Byr") = "R" Then
                OptTrans.Value = True
            End If
        End With
        Tanya = MsgBox("Policy data  already exist." + Chr(13) + "Do you want to edit ?", vbQuestion + vbYesNo, "Confirmation")
        If Tanya = vbYes Then
            Aktif
            TxtNoPolis.Enabled = False: CmdCari.Enabled = False
            CmdTambah.Enabled = False
            CmdUbah.Enabled = True: CmdHapus.Enabled = True
        ElseIf Tanya = vbNo Then
            Bersih
            TxtNoPolis.Text = ""
            NonAktif
            CmdUbah.Enabled = False: CmdHapus.Enabled = False
            TxtNoPolis.SetFocus
        End If
    End If
    RsPolis.Close
End If
End Sub

Private Sub TxtPolHolder_Change()
If TxtPolHolder.Text = "" Then Exit Sub
SQL = "select * from Nasabah2 where No_PolHolder='" & Left(TxtPolHolder.Text, 5) & "'"
If RsNasabah2.State > 0 Then RsNasabah2.Close
RsNasabah2.CursorLocation = adUseClient
RsNasabah2.Open SQL, Cn, adOpenDynamic, adLockOptimistic
If RsNasabah2.RecordCount > 0 Then
   With RsNasabah2
      LblNama2.Caption = .Fields("Nama")
      LblTglLhr2.Caption = .Fields("Tgl_Lahir")
      LblJKel2.Caption = .Fields("Jns_Kel")
      .Close
   End With
Else
      MsgBox "Policy Holder data not found.", vbCritical, "Attention"
      TxtPolHolder.SetFocus
End If
End Sub

Private Sub TxtPremi_Change()
If Trim(TxtPremi.Text) <> "" Then
   dig$ = Mid(TxtPremi.Text, Len(TxtPremi.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtPremi.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtPremi.Text = digits$
          TxtPremi.SelStart = Len(TxtPremi.Text)
      End If
End If
End Sub

Private Sub TxtTertngg_Change()
If TxtTertngg.Text = "" Then Exit Sub
SQL = "select * from Nasabah1 where No_Tertngg='" & Left(TxtTertngg.Text, 5) & "'"
If RsNasabah1.State > 0 Then RsNasabah1.Close
RsNasabah1.CursorLocation = adUseClient
RsNasabah1.Open SQL, Cn, adOpenDynamic, adLockOptimistic
If RsNasabah1.RecordCount > 0 Then
   With RsNasabah1
      LblNama1.Caption = .Fields("Nama")
      LblTglLhr1.Caption = .Fields("Tgl_Lahir")
      LblJKel1.Caption = .Fields("Jns_Kel")
      .Close
   End With
End If
If TxtTertngg.Text = "00000" Then
    LblNama1.Caption = LblNama2.Caption
    LblTglLhr1.Caption = LblTglLhr2.Caption
    LblJKel1.Caption = LblJKel2.Caption
End If
End Sub

Private Sub TxtUP_Change()
If Trim(TxtUP.Text) <> "" Then
   dig$ = Mid(TxtUP.Text, Len(TxtUP.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtUP.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtUP.Text = digits$
          TxtUP.SelStart = Len(TxtUP.Text)
      End If
End If
End Sub
