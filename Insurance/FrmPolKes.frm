VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmPolKes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Polis Asuransi Kesehatan Allianz"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "FrmPolKes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11070
   Begin VB.CommandButton CmdCari 
      BackColor       =   &H008080FF&
      Caption         =   "&Cari"
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
   Begin VB.Frame FrmTertngg 
      BackColor       =   &H00E0E0E0&
      Height          =   5175
      Left            =   4920
      TabIndex        =   54
      Top             =   2160
      Width           =   5895
      Begin MSMask.MaskEdBox MskTglLhr 
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtNama 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame FrmJKel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jenis Kelamin"
         Height          =   855
         Left            =   3600
         TabIndex        =   57
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton OptPria 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Laki-Laki"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptWanita 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Wanita"
            Height          =   255
            Left            =   1200
            TabIndex        =   20
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame FrmTmbl1 
         BackColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   120
         TabIndex        =   55
         Top             =   4080
         Width           =   5535
         Begin VB.CommandButton CmdBersih 
            BackColor       =   &H0080C0FF&
            Caption         =   "&BERSIH"
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
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton CmdHapus1 
            BackColor       =   &H0080C0FF&
            Caption         =   "&HAPUS"
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
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdTambah1 
            BackColor       =   &H0080C0FF&
            Caption         =   "&TAMBAH"
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
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DtG1 
         Height          =   2775
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4895
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
         Caption         =   "Data Tertanggung"
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
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtPremi 
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
      Left            =   1920
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame FrmProduk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Jenis Pertanggungan"
      Height          =   2055
      Left            =   240
      TabIndex        =   42
      Top             =   3960
      Width           =   4455
      Begin VB.ComboBox CboPlan 
         Height          =   315
         Index           =   4
         Left            =   3360
         TabIndex        =   16
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox CboPlan 
         Height          =   315
         Index           =   3
         Left            =   3360
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox CboPlan 
         Height          =   315
         Index           =   2
         Left            =   3360
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox CboPlan 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox CboRawat 
         Height          =   315
         Index           =   4
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox CboRawat 
         Height          =   315
         Index           =   3
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox CboRawat 
         Height          =   315
         Index           =   2
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox CboRawat 
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox CboRawat 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox CboPlan 
         Height          =   315
         Index           =   0
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
         Height          =   255
         Left            =   2760
         TabIndex        =   52
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
         Height          =   255
         Left            =   2760
         TabIndex        =   51
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Santunan Harian"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Gigi"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
         Height          =   255
         Left            =   2760
         TabIndex        =   46
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Melahirkan"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Rawat Jalan"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rawat Inap"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   240
      TabIndex        =   41
      Top             =   6240
      Width           =   4455
      Begin VB.CommandButton CmdHapus2 
         BackColor       =   &H00C0C000&
         Caption         =   "HA&PUS"
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdUbah2 
         BackColor       =   &H00C0C000&
         Caption         =   "U&BAH"
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdTambah2 
         BackColor       =   &H00C0C000&
         Caption         =   "TA&MBAH"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdBatal 
         BackColor       =   &H00C0C000&
         Caption         =   "BATA&L"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox CboAgen 
      Height          =   315
      Left            =   6600
      TabIndex        =   5
      Top             =   840
      Width           =   3015
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
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   1
      ToolTipText     =   "Tekan ENTER untuk lanjutkan"
      Top             =   840
      Width           =   1815
   End
   Begin VB.Frame FrmPolHolder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pemegang Polis"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   4455
      Begin VB.CommandButton CmdPolHolder 
         Caption         =   "...."
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
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox TxtPolHolder 
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
         Height          =   360
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   62
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   255
         Left            =   3480
         TabIndex        =   64
         Top             =   480
         Width           =   375
      End
      Begin VB.Label LblJKel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   63
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Umur"
         Height          =   255
         Left            =   3360
         TabIndex        =   60
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pemegang Polis"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pemegang Polis"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label LblNama 
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
         TabIndex        =   30
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label LblTglLhr 
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
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblUmur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   28
         Top             =   1200
         Width           =   495
      End
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   24969217
      CurrentDate     =   38465
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp"
      Height          =   255
      Left            =   1560
      TabIndex        =   61
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Besar Premi"
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DAFTAR TERTANGGUNG"
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
      Left            =   4920
      TabIndex        =   40
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POLIS ASURANSI KESEHATAN ALLIANZ"
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
      TabIndex        =   39
      Top             =   120
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Polis"
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
      TabIndex        =   38
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Kontrak"
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Agen"
      Height          =   255
      Left            =   4920
      TabIndex        =   36
      Top             =   840
      Width           =   855
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
      Width           =   9375
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   9960
      Picture         =   "FrmPolKes.frx":0442
      Top             =   120
      Width           =   840
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   375
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Polis"
      Height          =   255
      Left            =   4920
      TabIndex        =   35
      Top             =   1200
      Width           =   855
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
      Left            =   6600
      TabIndex        =   34
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "FrmPolKes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPolis As New ADODB.Recordset
Dim RsAgen As New ADODB.Recordset
Dim RsNasabah2 As New ADODB.Recordset
Dim RsRelasi As New ADODB.Recordset
Dim RsDetPolis As New ADODB.Recordset

Private Sub CmdPolHolder_Click()
FrmCari_PolHolder.LblTujuan.Caption = Me.Name
FrmCari_PolHolder.Show 1
End Sub

Private Sub TxtPolHolder_Change()
If TxtPolHolder.Text = "" Then Exit Sub
'Isi data pemegang polis
SQL = "select * from Nasabah2 where No_PolHolder='" & Left(TxtPolHolder.Text, 5) & "'"
RsNasabah2.Open SQL, Cn, adOpenDynamic, adLockOptimistic
If RsNasabah2.RecordCount > 0 Then
With RsNasabah2
    LblNama.Caption = .Fields("Nama")
    LblTglLhr.Caption = .Fields("Tgl_Lahir")
    LblJKel.Caption = .Fields("Jns_Kel")
    .Close
End With
If TxtPolHolder.Text <> "" Then
    FrmProduk.Enabled = True
End If
End If
End Sub

Private Sub CboRawat_Click(Index As Integer)
If CboRawat(Index).ListIndex = 0 Then
    CboPlan(Index).Enabled = True
    CboPlan(Index).Text = ""
ElseIf CboRawat(Index).ListIndex = 1 Then
    CboPlan(Index).Enabled = False
    CboPlan(Index).Text = ""
End If
End Sub

Private Sub CmdBatal_Click()
TxtNoPolis.Text = ""
Bersih
NonAktif
CmdTambah2.Enabled = False
CmdUbah2.Enabled = False: CmdHapus2.Enabled = False
'Refresh Daftar Tertanggung
RefGrid
TxtNoPolis.Enabled = True: CmdCari.Enabled = True
TxtNoPolis.SetFocus
End Sub

Private Sub CmdBersih_Click()
Bersih1
CmdHapus1.Enabled = False
TxtNama.SetFocus
End Sub

Private Sub CmdCari_Click()
FrmCari_PolKes.LblTujuan.Caption = Me.Name
FrmCari_PolKes.Show 1
End Sub

Private Sub CmdHapus1_Click()
SQL1 = "delete from Detail_Polkes where No_Polkes='" & DtG1.Columns(0) & _
"'and Nama='" & DtG1.Columns(1) & "'"
If RsDetPolis.State > 0 Then RsDetPolis.Close
Set RsDetPolis = Cn.Execute(SQL1)
'Refresh Daftar Tertanggung
RefGrid
Bersih1
TxtNama.SetFocus
End Sub

Private Sub CmdHapus2_Click()
SQL1 = "select * from PolKes where No_PolKes='" & TxtNoPolis.Text & "'"
RsPolis.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
SQL2 = "select * from Detail_Polkes where No_PolKes='" & TxtNoPolis.Text & "'"
If RsDetPolis.State > 0 Then RsDetPolis.Close
Set DtG1.DataSource = Cn.Execute(SQL2)
RsDetPolis.Open SQL2, Cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Anda yakin ingin menghapus data Polis ini ?", vbQuestion + vbYesNo, "Konfirmasi Hapus")
If Tanya = vbYes Then
    RsPolis.Delete
    RsPolis.Close
    'Hapus Daftar Tertanggung
    For i = 1 To RsDetPolis.RecordCount
        RsDetPolis.Delete
        RsDetPolis.MoveNext
    Next i
    RsDetPolis.Close
    MsgBox "Data Polis telah berhasil dihapus.", vbInformation, "Informasi"
End If
TxtNoPolis.Text = ""
'Refresh Daftar Tertanggung
RefGrid
Bersih
Bersih1
NonAktif
CmdUbah2.Enabled = False: CmdHapus2.Enabled = False
TxtNoPolis.Enabled = True: CmdCari.Enabled = True
TxtNoPolis.SetFocus
End Sub

Private Sub CmdTambah1_Click()
If Trim(TxtNama.Text) = "" Then
    MsgBox "Nama Tertanggung tidak boleh kosong.", vbCritical, "Perhatian"
    TxtNama.SetFocus
ElseIf Trim(MskTglLhr.Text) = "" Then
    MsgBox "Tanggal Lahir Tertanggung tidak boleh kosong.", vbCritical, "Perhatian"
    MskTglLhr.SetFocus
Else
SQL = "select * from Detail_PolKes where No_PolKes='" & TxtNoPolis.Text & "'"
If RsDetPolis.State > 0 Then RsDetPolis.Close
RsDetPolis.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsDetPolis
    .AddNew
    .Fields("No_PolKes") = Trim(TxtNoPolis.Text)
    .Fields("Nama") = TxtNama.Text
    .Fields("Tgl_Lahir") = CDate(Trim(MskTglLhr.Text))
    If OptPria.Value = True Then
        .Fields("J_Kel") = "L"
    ElseIf OptWanita.Value = True Then
        .Fields("J_Kel") = "P"
    End If
    .Update
    Set DtG1.DataSource = RsDetPolis
    End With
    Bersih1
    TxtNama.SetFocus
End If
End Sub

Private Sub Bersih1()
TxtNama.Text = "": MskTglLhr.Mask = ""
MskTglLhr.Text = "": MskTglLhr.Mask = "##/##/####"
OptPria.Value = True: CmdHapus1.Enabled = False
End Sub

Private Sub CmdTambah2_Click()
If Trim(CboAgen.Text) = "" Then
    MsgBox "Pilih salah satu Agen yang ada.", vbCritical, "Perhatian"
    CboAgen.SetFocus
ElseIf Trim(TxtPremi.Text) = "" Then
    MsgBox "Besar Premi tidak boleh kosong.", vbCritical, "Perhatian"
    TxtPremi.SetFocus
ElseIf Trim(TxtPolHolder.Text) = "" Then
    MsgBox "Kode Pemegang Polis tidak boleh kosong.", vbCritical, "Perhatian"
    CmdPolHolder.SetFocus
ElseIf Trim(CboRawat(Index).Text) = "" Then
    MsgBox "Pilih Jenis Perawatan yang ada.", vbCritical, "Perhatian"
ElseIf (CboRawat(Index).ListIndex = 0) And (Trim(CboPlan(Index).Text) = "") Then
    MsgBox "Pilih salah satu Plan yang ada.", vbCritical, "Perhatian"
Else
SQL = "select * from PolKes order by No_PolKes"
If RsPolis.State > 0 Then RsPolis.Close
RsPolis.Open SQL, Cn, adOpenDynamic, adLockOptimistic
With RsPolis
    .AddNew
    .Fields("No_PolKes") = Trim(TxtNoPolis.Text)
    Simpan
    .Fields("Status") = Left(LblStatus.Caption, 1)
    .Update
    .Close
End With
MsgBox "Data Polis telah berhasil ditambahkan dengan sukses.", vbInformation, "Informasi"
TxtNoPolis.Text = ""
'Refresh Daftar Tertanggung
RefGrid
Bersih
Bersih1
NonAktif
CmdTambah2.Enabled = False
TxtNoPolis.Enabled = True: CmdCari.Enabled = True
TxtNoPolis.SetFocus
End If
End Sub

Private Sub CmdUbah2_Click()
If Trim(CboAgen.Text) = "" Then
    MsgBox "Pilih salah satu Agen yang ada.", vbCritical, "Perhatian"
    CboAgen.SetFocus
ElseIf Trim(TxtPremi.Text) = "" Then
    MsgBox "Besar Premi tidak boleh kosong.", vbCritical, "Perhatian"
    TxtPremi.SetFocus
ElseIf Trim(TxtPolHolder.Text) = "" Then
    MsgBox "Kode Pemegang Polis tidak boleh kosong.", vbCritical, "Perhatian"
    CmdPolHolder.SetFocus
ElseIf Trim(CboRawat(Index).Text) = "" Then
    MsgBox "Pilih Jenis Perawatan yang ada.", vbCritical, "Perhatian"
ElseIf (CboRawat(Index).ListIndex = 0) And (Trim(CboPlan(Index).Text) = "") Then
    MsgBox "Pilih salah satu Plan yang ada.", vbCritical, "Perhatian"
Else
SQL = "select * from PolKes where No_PolKes='" & TxtNoPolis.Text & "'"
If RsPolis.State > 0 Then RsPolis.Close
RsPolis.Open SQL, Cn, adOpenDynamic, adLockOptimistic
With RsPolis
    Simpan
    .Update
    .Close
End With
MsgBox "Data Polis telah berhasil diubah dengan sukses.", vbInformation, "Informasi"
TxtNoPolis.Text = ""
'Refresh Daftar Tertanggung
RefGrid
Bersih1
Bersih
NonAktif
CmdUbah2.Enabled = False: CmdHapus2.Enabled = False
TxtNoPolis.Enabled = True: CmdCari.Enabled = True
TxtNoPolis.SetFocus
End If
End Sub

Private Sub DtG1_DblClick()
If RsDetPolis.RecordCount = 0 Then Exit Sub
TxtNama.Text = DtG1.Columns(1)
If Trim(DtG1.Columns(2)) = "L" Then
    OptPria.Value = True
ElseIf Trim(DtG1.Columns(2)) = "P" Then
    OptWanita.Value = True
End If
MskTglLhr.Text = DtG1.Columns(3)
CmdHapus1.Enabled = True
End Sub

Private Sub Form_Activate()
TxtNoPolis.SetFocus
End Sub

Private Sub Form_Load()
Left = 100: Top = 100
DTP1.Value = Date
Koneksi
Cn.CursorLocation = adUseClient
RsPolis.CursorLocation = adUseClient
'Isi combobox Kode Agen
RsAgen.CursorLocation = adUseClient
SQL1 = "select Kd_Agen,Nama_Agen from Agen where Status='A' order by Kd_Agen"
RsAgen.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
RsAgen.MoveFirst
For t = 1 To RsAgen.RecordCount
    Isi = Trim(RsAgen.Fields("Kd_Agen")) & " - " & Trim(RsAgen.Fields("Nama_Agen"))
    CboAgen.AddItem Isi
    RsAgen.MoveNext
Next t
RsAgen.Close
'Isi combobox Rawat dan Plan
For W = 0 To 4
    CboRawat(W).AddItem "Y"
    CboRawat(W).AddItem "N"
    CboPlan(W).AddItem "A"
    CboPlan(W).AddItem "B"
    CboPlan(W).AddItem "C"
    CboPlan(W).AddItem "D"
    CboPlan(W).AddItem "E"
    CboPlan(W).AddItem "F"
Next W
'Refresh Daftar Tertanggung
RefGrid
TxtNoPolis.Text = ""
Bersih
Bersih1
NonAktif
CmdTambah2.Enabled = False
CmdUbah2.Enabled = False
CmdHapus1.Enabled = False
CmdHapus2.Enabled = False
End Sub

Private Sub Simpan()
With RsPolis
    .Fields("Tgl_Kontrak") = DTP1.Value
    .Fields("No_PolHolder") = Trim(TxtPolHolder.Text)
    .Fields("Premi") = CCur(Trim(TxtPremi.Text))
    .Fields("Kd_Agen") = Left(CboAgen.Text, 8)
    If CboRawat(0).ListIndex = 0 Then
        .Fields("Rwt_Inap") = CboRawat(0).Text & CboPlan(0).Text
    ElseIf CboRawat(0).ListIndex = 1 Then
        .Fields("Rwt_Inap") = CboRawat(0).Text & " "
    End If
    If CboRawat(1).ListIndex = 0 Then
        .Fields("Rwt_Jln") = CboRawat(1).Text & CboPlan(1).Text
    ElseIf CboRawat(1).ListIndex = 1 Then
        .Fields("Rwt_Jln") = CboRawat(1).Text & " "
    End If
    If CboRawat(2).ListIndex = 0 Then
        .Fields("Melahirkan") = CboRawat(2).Text & CboPlan(2).Text
    ElseIf CboRawat(2).ListIndex = 1 Then
        .Fields("Melahirkan") = CboRawat(2).Text & " "
    End If
    If CboRawat(3).ListIndex = 0 Then
        .Fields("Gigi") = CboRawat(3).Text & CboPlan(3).Text
    ElseIf CboRawat(3).ListIndex = 1 Then
        .Fields("Gigi") = CboRawat(3).Text & " "
    End If
    If CboRawat(4).ListIndex = 0 Then
        .Fields("Stn_Harian") = CboRawat(4).Text & CboPlan(4).Text
    ElseIf CboRawat(4).ListIndex = 1 Then
        .Fields("Stn_Harian") = CboRawat(4).Text & " "
    End If
End With
End Sub

Private Sub Bersih()
    DTP1.Value = Date: LblStatus.Caption = "A - Aktif"
    CboAgen.Text = "": TxtPolHolder.Text = ""
    LblNama.Caption = "": LblUmur.Caption = "": LblTglLhr.Caption = ""
    TxtPremi.Text = "": OptPria.Value = True
    For i = 0 To 4
        CboRawat(i).Text = ""
        CboPlan(i).Text = ""
    Next i
    LblJKel.Caption = ""
End Sub

Private Sub NonAktif()
    TxtPremi.Enabled = False: CboAgen.Enabled = False
    FrmPolHolder.Enabled = False: FrmProduk.Enabled = False
    FrmTertngg.Enabled = False: DTP1.Enabled = False
End Sub

Private Sub Aktif()
    TxtPremi.Enabled = True: CboAgen.Enabled = True
    FrmPolHolder.Enabled = True: FrmProduk.Enabled = True
    FrmTertngg.Enabled = True: DTP1.Enabled = True
End Sub

Private Sub RefGrid()
'Tampilkan Daftar Tertanggung
SQL = "select * from Detail_PolKes where No_PolKes='" & TxtNoPolis.Text & "'"
If RsDetPolis.State > 0 Then RsDetPolis.Close
Set RsDetPolis = Cn.Execute(SQL)
Set DtG1.DataSource = RsDetPolis
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsPolis = Nothing
    Set RsAgen = Nothing
    Set RsNasabah2 = Nothing
    Set RsRelasi = Nothing
    Set RsDetPolis = Nothing
End Sub

Private Sub LblTglLhr_Change()
If Trim(LblTglLhr.Caption) = "" Then Exit Sub
'Hitung umur Pemegang Polis
If (Month(Now) = Month(LblTglLhr.Caption)) And (Day(Now) >= Day(LblTglLhr.Caption)) Then
    LblUmur.Caption = Year(Now) - Year(LblTglLhr.Caption)
ElseIf Month(Now) > Month(LblTglLhr.Caption) Then
    LblUmur.Caption = Year(Now) - Year(LblTglLhr.Caption)
Else
    LblUmur.Caption = (Year(Now) - Year(LblTglLhr.Caption)) - 1
End If
End Sub

Private Sub TxtNoPolis_Change()
If Trim(TxtNoPolis.Text) <> "" Then
   dig$ = Mid(TxtNoPolis.Text, Len(TxtNoPolis.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Hanya dapat diinput dengan angka !", vbCritical, "Error"
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
    If Len(Trim(TxtNoPolis.Text)) < 12 Then Exit Sub
    SQL1 = "select * from PolKes where No_PolKes='" & TxtNoPolis.Text & "'"
    If RsPolis.State > 0 Then RsPolis.Close
    RsPolis.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
    If RsPolis.RecordCount = 0 Then
        Aktif
        Bersih
        'Refresh Daftar Tertanggung
        RefGrid
        TxtNoPolis.Enabled = False: CmdCari.Enabled = False
        CmdTambah2.Enabled = True
        CmdUbah2.Enabled = False: CmdHapus2.Enabled = False
        CboAgen.SetFocus
    ElseIf RsPolis.RecordCount > 0 Then
        NonAktif
        'Tampilkan data
        With RsPolis
            DTP1.Value = .Fields("Tgl_Kontrak")
            'Isi keterangan Status
            If .Fields("Status") = "A" Then
                LblStatus.Caption = "A - Aktif"
            ElseIf .Fields("Status") = "L" Then
                LblStatus.Caption = "L - Lapse"
            ElseIf .Fields("Status") = "B" Then
                LblStatus.Caption = "B - Berakhir"
            End If
            TxtPolHolder.Text = .Fields("No_PolHolder")
            'Tampilkan data Pemegang Polis
            SQL3 = "select Nama,Tgl_Lahir from Nasabah2 where No_PolHolder='" & TxtPolHolder.Text & "'"
            RsNasabah2.Open SQL3, Cn, adOpenDynamic, adLockOptimistic
            LblNama.Caption = RsNasabah2.Fields("Nama")
            LblTglLhr.Caption = RsNasabah2.Fields("Tgl_Lahir")
            RsNasabah2.Close
            'Isi Nama Agen
            SQL4 = "select Nama_Agen from Agen where Kd_Agen='" & .Fields("Kd_Agen") & "'"
            RsAgen.Open SQL4, Cn, adOpenDynamic, adLockOptimistic
            Nm_Agen = .Fields("Kd_Agen") & " - " & RsAgen.Fields("Nama_Agen")
            CboAgen.Text = Nm_Agen
            RsAgen.Close
            TxtPremi.Text = Format(.Fields("Premi"), "#,#,#,#,0")
            'Tampilkan data Pertanggungan
            CboRawat(0).Text = Left(.Fields("Rwt_Inap"), 1)
            CboRawat(1).Text = Left(.Fields("Rwt_Jln"), 1)
            CboRawat(2).Text = Left(.Fields("Melahirkan"), 1)
            CboRawat(3).Text = Left(.Fields("Gigi"), 1)
            CboRawat(4).Text = Left(.Fields("Stn_Harian"), 1)
            CboPlan(0).Text = Right(.Fields("Rwt_Inap"), 1)
            CboPlan(1).Text = Right(.Fields("Rwt_Jln"), 1)
            CboPlan(2).Text = Right(.Fields("Melahirkan"), 1)
            CboPlan(3).Text = Right(.Fields("Gigi"), 1)
            CboPlan(4).Text = Right(.Fields("Stn_Harian"), 1)
            'Tampilkan Daftar Tertanggung
            RefGrid
        End With
        Tanya = MsgBox("Data Polis telah ada." + Chr(13) + "Anda ingin mengubahnya ?", vbQuestion + vbYesNo, "Konfirmasi")
        If Tanya = vbYes Then
            Aktif
            TxtNoPolis.Enabled = False: CmdCari.Enabled = False
            CmdUbah2.Enabled = True: CmdHapus2.Enabled = True
        ElseIf Tanya = vbNo Then
            Bersih
            TxtNoPolis.Text = ""
            'Refresh Daftar Tertanggung
            RefGrid
            NonAktif
            CmdUbah2.Enabled = False: CmdHapus2.Enabled = False
            TxtNoPolis.SetFocus
        End If
    End If
    RsPolis.Close
End If
End Sub

Private Sub TxtPremi_Change()
If Trim(TxtPremi.Text) <> "" Then
   dig$ = Mid(TxtPremi.Text, Len(TxtPremi.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Hanya dapat diinput dengan angka !", vbCritical, "Error"
      For i = 1 To Len(TxtPremi.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtPremi.Text = digits$
          TxtPremi.SelStart = Len(TxtPremi.Text)
      End If
End If
End Sub
