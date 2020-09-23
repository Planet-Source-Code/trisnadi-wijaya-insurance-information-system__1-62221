VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmTertanggung 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insured Data"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "FrmNasabah.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   7215
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
      Left            =   5640
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6360
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
      Left            =   3960
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6360
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
      Left            =   2280
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6360
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
      Left            =   600
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6360
      Width           =   975
   End
   Begin TabDlg.SSTab TabTertanggung 
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   128
      TabCaption(0)   =   "Insured Personal Data"
      TabPicture(0)   =   "FrmNasabah.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label17"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Image1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtNo1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtNama1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CboStatus1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CboAgama1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtAlmRmh1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtKotaRmh1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtPosRmh1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtAreaRmh1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtTlpRmh1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtHP1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ChkOtomat1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "FrmJKel"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CmdTampil"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "MskTglLhr"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Insured Occupation Data"
      TabPicture(1)   =   "FrmNasabah.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtJabatan1"
      Tab(1).Control(1)=   "TxtKotaKntr1"
      Tab(1).Control(2)=   "FrmKoresponden"
      Tab(1).Control(3)=   "TxtTlpKntr1"
      Tab(1).Control(4)=   "TxtAreaKntr1"
      Tab(1).Control(5)=   "TxtPosKntr1"
      Tab(1).Control(6)=   "TxtAlmKntr1"
      Tab(1).Control(7)=   "TxtKntr1"
      Tab(1).Control(8)=   "Label27"
      Tab(1).Control(9)=   "Label23"
      Tab(1).Control(10)=   "Label22"
      Tab(1).Control(11)=   "Label21"
      Tab(1).Control(12)=   "Label20"
      Tab(1).Control(13)=   "Label19"
      Tab(1).Control(14)=   "Label18"
      Tab(1).ControlCount=   15
      Begin MSMask.MaskEdBox MskTglLhr 
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtJabatan1 
         Height          =   285
         Left            =   -72240
         MaxLength       =   20
         TabIndex        =   16
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton CmdTampil 
         BackColor       =   &H008080FF&
         Caption         =   "SEARCH"
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
         TabIndex        =   29
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox TxtKotaKntr1 
         Height          =   285
         Left            =   -72240
         TabIndex        =   19
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Frame FrmKoresponden 
         Height          =   615
         Left            =   -72240
         TabIndex        =   46
         Top             =   3840
         Width           =   2295
         Begin VB.OptionButton OptRmh1 
            Caption         =   "Home"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptKntr1 
            Caption         =   "Office"
            Height          =   195
            Left            =   1200
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame FrmJKel 
         Caption         =   "Gender"
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   4440
         TabIndex        =   44
         Top             =   1920
         Width           =   2295
         Begin VB.OptionButton OptLaki1 
            Caption         =   "Male"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptWanita1 
            Caption         =   "Female"
            Height          =   195
            Left            =   1320
            TabIndex        =   45
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox TxtTlpKntr1 
         Height          =   285
         Left            =   -71520
         MaxLength       =   7
         TabIndex        =   21
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox TxtAreaKntr1 
         Height          =   285
         Left            =   -72240
         MaxLength       =   4
         TabIndex        =   20
         Top             =   3300
         Width           =   615
      End
      Begin VB.TextBox TxtPosKntr1 
         Height          =   285
         Left            =   -72240
         MaxLength       =   6
         TabIndex        =   18
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtAlmKntr1 
         Height          =   285
         Left            =   -72240
         MaxLength       =   40
         TabIndex        =   17
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox TxtKntr1 
         Height          =   285
         Left            =   -72240
         MaxLength       =   20
         TabIndex        =   15
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox ChkOtomat1 
         Caption         =   "Automatic"
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtHP1 
         Height          =   285
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   12
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox TxtTlpRmh1 
         Height          =   285
         Left            =   3120
         MaxLength       =   7
         TabIndex        =   11
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox TxtAreaRmh1 
         Height          =   285
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   10
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox TxtPosRmh1 
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   9
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox TxtKotaRmh1 
         Height          =   285
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   8
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox TxtAlmRmh1 
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   7
         Top             =   3240
         Width           =   4455
      End
      Begin VB.ComboBox CboAgama1 
         Height          =   315
         ItemData        =   "FrmNasabah.frx":047A
         Left            =   2400
         List            =   "FrmNasabah.frx":048D
         TabIndex        =   6
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox CboStatus1 
         Height          =   315
         ItemData        =   "FrmNasabah.frx":04C2
         Left            =   2400
         List            =   "FrmNasabah.frx":04D2
         TabIndex        =   4
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox TxtNama1 
         Height          =   285
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox TxtNo1 
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
         Height          =   420
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   0
         ToolTipText     =   "Press ENTER to continue"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   795
         Left            =   5880
         Picture         =   "FrmNasabah.frx":04F8
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label27 
         Caption         =   "City"
         Height          =   255
         Left            =   -74760
         TabIndex        =   47
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "Corespondency Address"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label22 
         Caption         =   "Job Title"
         Height          =   255
         Left            =   -74760
         TabIndex        =   42
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Office Phone No."
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   3300
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Office Zip Code"
         Height          =   255
         Left            =   -74760
         TabIndex        =   40
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "Office Address"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Company Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label17 
         Caption         =   "Cell Phone No."
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Home Phone No."
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Home Zip Code"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "City"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Home Address"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Religion"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Marital Status"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Insured Name"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Insured No."
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmTertanggung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTertngg As New ADODB.Recordset

'Generates auto numbering for Insured No.
Private Sub Penomoran()
    Dim Nom As String
    Dim NM As Integer
    SQL = "select * from Nasabah1 order by No_Tertngg"
    If RsTertngg.State > 0 Then RsTertngg.Close
    RsTertngg.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    If RsTertngg.RecordCount = 0 Then
        Nom = "00001"
    Else
        RsTertngg.MoveLast
        NM = Val(Trim(RsTertngg.Fields(0))) + 1
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
    TxtNo1.Text = Nom
    RsTertngg.Close
End Sub

Private Sub ChkOtomat1_Click()
If ChkOtomat1.Value = 1 Then
    Penomoran
    TxtNo1.SetFocus
ElseIf ChkOtomat1.Value = 0 Then
    TxtNo1.Text = ""
    TxtNo1.Enabled = True
    TxtNo1.SetFocus
End If
End Sub

Private Sub CmdBatal_Click()
Bersih
NonAktif
CmdTambah.Enabled = False
CmdUbah.Enabled = False: CmdHapus.Enabled = False
TxtNo1.Enabled = True: ChkOtomat1.Enabled = True
TxtNo1.SetFocus
End Sub

Private Sub Simpan()
With RsTertngg
    .Fields("Nama") = TxtNama1.Text
    .Fields("Tgl_Lahir") = MskTglLhr.Text
    If OptLaki1.Value = True Then
        .Fields("Jns_Kel") = "M"
    ElseIf OptWanita1.Value = True Then
        .Fields("Jns_Kel") = "F"
    End If
    If CboStatus1.ListIndex = 0 Then
        .Fields("Marital_Status") = "S"
    ElseIf CboStatus1.ListIndex = 1 Then
        .Fields("Marital_Status") = "M"
    ElseIf CboStatus1.ListIndex = 2 Then
        .Fields("Marital_Status") = "D"
    ElseIf CboStatus1.ListIndex = 3 Then
        .Fields("Marital_Status") = "W"
    End If
    If CboAgama1.ListIndex = 0 Then
        .Fields("Agama") = "I"
    ElseIf CboAgama1.ListIndex = 1 Then
        .Fields("Agama") = "P"
    ElseIf CboAgama1.ListIndex = 2 Then
        .Fields("Agama") = "K"
    ElseIf CboAgama1.ListIndex = 3 Then
        .Fields("Agama") = "B"
    ElseIf CboAgama1.ListIndex = 4 Then
        .Fields("Agama") = "H"
    End If
    .Fields("Alm_Rmh") = TxtAlmRmh1.Text
    .Fields("Kota_Rmh") = TxtKotaRmh1.Text
    .Fields("Kd_Pos_Rmh") = Trim(TxtPosRmh1.Text)
    .Fields("Area_Telp_Rmh") = Trim(TxtAreaRmh1.Text)
    .Fields("No_Telp_Rmh") = Trim(TxtTlpRmh1.Text)
    .Fields("Ponsel") = Trim(TxtHP1.Text)
    .Fields("Tmpt_Kerja") = TxtKntr1.Text
    .Fields("Jabatan") = TxtJabatan1.Text
    .Fields("Alm_Kntr") = TxtAlmKntr1.Text
    .Fields("Kd_Pos_Kntr") = Trim(TxtPosKntr1.Text)
    .Fields("Kota_Kntr") = TxtKotaKntr1.Text
    .Fields("Area_Telp_Kntr") = Trim(TxtAreaKntr1.Text)
    .Fields("No_Telp_Kntr") = Trim(TxtTlpKntr1.Text)
    If OptRmh1.Value = True Then
        .Fields("Korespondensi") = "H"
    ElseIf OptKntr1.Value = True Then
        .Fields("Korespondensi") = "O"
    End If
End With
End Sub

Private Sub CmdHapus_Click()
SQL = "select * from Nasabah1 where No_Tertngg ='" & TxtNo1.Text & "'"
If RsTertngg.State > 0 Then RsTertngg.Close
RsTertngg.Open SQL, Cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Are you sure want to delete " & TxtNama1.Text & "?", vbQuestion + vbYesNo, "Delete Confirmation")
If Tanya = vbYes Then
    RsTertngg.Delete
    MsgBox "Insured data deleted.", vbInformation, "Information"
End If
RsTertngg.Close
Bersih
NonAktif
CmdUbah.Enabled = False: CmdHapus.Enabled = False
TxtNo1.Enabled = True: ChkOtomat1.Enabled = True
TxtNo1.SetFocus
End Sub

Private Sub CmdTambah_Click()
    If RsTertngg.State > 0 Then RsTertngg.Close
    'Save data
    RsTertngg.Open "Nasabah1", Cn, adOpenDynamic, adLockOptimistic
    With RsTertngg
        .AddNew
        .Fields("No_Tertngg") = Trim(TxtNo1.Text)
        Simpan
        .Update
        .Close
    End With
MsgBox "Insured data saved.", vbInformation, "Information"
Bersih
NonAktif
TxtNo1.Enabled = True: ChkOtomat1.Enabled = True
TxtNo1.SetFocus
End Sub

Private Sub CmdTampil_Click()
FrmCari_Tertngg.Show 1
End Sub

Private Sub CmdUbah_Click()
    'Edit data
    SQL = "select * from Nasabah1 where No_Tertngg ='" & TxtNo1.Text & "'"
    If RsTertngg.State > 0 Then RsTertngg.Close
    RsTertngg.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsTertngg
        Simpan
        .Update
        .Close
    End With
MsgBox "Insured data changed.", vbInformation, "Information"
Bersih
NonAktif
CmdUbah.Enabled = False: CmdHapus.Enabled = False
TxtNo1.Enabled = True: ChkOtomat1.Enabled = True
TxtNo1.SetFocus
End Sub

Private Sub Form_Activate()
TxtNo1.SetFocus
End Sub

Private Sub Form_Load()
Top = 100: Left = 100
Koneksi
RsTertngg.CursorLocation = adUseClient
Bersih
NonAktif
CmdTambah.Enabled = False
CmdUbah.Enabled = False
CmdHapus.Enabled = False
End Sub

Private Sub Bersih()
For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
    If TypeOf Control Is ComboBox Then
        Control.Text = ""
    End If
Next Control
MskTglLhr.Mask = ""
MskTglLhr.Text = "": MskTglLhr.Mask = "##/##/####"
OptLaki1.Value = True: ChkOtomat1.Value = 0
OptRmh1.Value = True
End Sub

Private Sub Aktif()
For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
        Control.Enabled = True
    End If
    If TypeOf Control Is ComboBox Then
        Control.Enabled = True
    End If
    If TypeOf Control Is OptionButton Then
        Control.Enabled = True
    End If
Next Control
    MskTglLhr.Enabled = True
End Sub

Private Sub NonAktif()
For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
        Control.Enabled = False
    End If
    If TypeOf Control Is ComboBox Then
        Control.Enabled = False
    End If
    If TypeOf Control Is OptionButton Then
        Control.Enabled = False
    End If
Next Control
MskTglLhr.Enabled = False
TxtNo1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsTertngg = Nothing
End Sub

Private Sub TxtAreaKntr1_Change()
If Trim(TxtAreaKntr1.Text) <> "" Then
   dig$ = Mid(TxtAreaKntr1.Text, Len(TxtAreaKntr1.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtAreaKntr1.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtAreaKntr1.Text = digits$
          TxtAreaKntr1.SelStart = Len(TxtAreaKntr1.Text)
      End If
End If
End Sub

Private Sub TxtAreaRmh1_Change()
If Trim(TxtAreaRmh1.Text) <> "" Then
   dig$ = Mid(TxtAreaRmh1.Text, Len(TxtAreaRmh1.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtAreaRmh1.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtAreaRmh1.Text = digits$
          TxtAreaRmh1.SelStart = Len(TxtAreaRmh1.Text)
      End If
End If
End Sub

Private Sub TxtHP1_Change()
If Trim(TxtHP1.Text) <> "" Then
   dig$ = Mid(TxtHP1.Text, Len(TxtHP1.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtHP1.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtHP1.Text = digits$
          TxtHP1.SelStart = Len(TxtHP1.Text)
      End If
End If
End Sub

Private Sub TxtNo1_Change()
If Trim(TxtNo1.Text) <> "" Then
   dig$ = Mid(TxtNo1.Text, Len(TxtNo1.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtNo1.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtNo1.Text = digits$
          TxtNo1.SelStart = Len(TxtNo1.Text)
      End If
End If
End Sub

Private Sub TxtNo1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And Trim(TxtNo1.Text) <> "" Then
    If Not IsNumeric(TxtNo1.Text) Then Exit Sub
    SQL = "select * from Nasabah1 where No_Tertngg ='" & TxtNo1.Text & "'"
    If RsTertngg.State > 0 Then RsTertngg.Close
    RsTertngg.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsTertngg
    If .RecordCount = 0 Then
        Aktif
        TxtNo1.Enabled = False: ChkOtomat1.Enabled = False
        CmdTambah.Enabled = True
        CmdUbah.Enabled = False: CmdHapus.Enabled = False
        TxtNama1.SetFocus
    Else
        'Retrieve data
        TxtNama1.Text = .Fields("Nama")
        MskTglLhr.Text = .Fields("Tgl_Lahir")
        If .Fields("Jns_Kel") = "M" Then
            OptLaki1.Value = True
        ElseIf .Fields("Jns_Kel") = "F" Then
            OptWanita1.Value = True
        End If
        If .Fields("Marital_Status") = "S" Then
            CboStatus1.ListIndex = 0
        ElseIf .Fields("Marital_Status") = "M" Then
            CboStatus1.ListIndex = 1
        ElseIf .Fields("Marital_Status") = "D" Then
            CboStatus1.ListIndex = 2
        ElseIf .Fields("Marital_Status") = "W" Then
            CboStatus1.ListIndex = 3
        End If
        If .Fields("Agama") = "I" Then
            CboAgama1.ListIndex = 0
        ElseIf .Fields("Agama") = "P" Then
            CboAgama1.ListIndex = 1
        ElseIf .Fields("Agama") = "K" Then
            CboAgama1.ListIndex = 2
        ElseIf .Fields("Agama") = "B" Then
            CboAgama1.ListIndex = 3
        ElseIf .Fields("Agama") = "H" Then
            CboAgama1.ListIndex = 4
        End If
        TxtAlmRmh1.Text = .Fields("Alm_Rmh")
        TxtKotaRmh1.Text = .Fields("Kota_Rmh")
        TxtPosRmh1.Text = Trim(.Fields("Kd_Pos_Rmh"))
        TxtAreaRmh1.Text = Trim(.Fields("Area_Telp_Rmh"))
        TxtTlpRmh1.Text = Trim(.Fields("No_Telp_Rmh"))
        TxtHP1.Text = Trim(.Fields("Ponsel"))
        TxtKntr1.Text = .Fields("Tmpt_Kerja")
        TxtJabatan1.Text = .Fields("Jabatan")
        TxtAlmKntr1.Text = .Fields("Alm_Kntr")
        TxtKotaKntr1.Text = .Fields("Kota_Kntr")
        TxtPosKntr1.Text = Trim(.Fields("Kd_Pos_Kntr"))
        TxtAreaKntr1.Text = Trim(.Fields("Area_Telp_Kntr"))
        TxtTlpKntr1.Text = Trim(.Fields("No_Telp_Kntr"))
        If .Fields("Korespondensi") = "H" Then
            OptRmh1.Value = True
        ElseIf .Fields("Korespondensi") = "O" Then
            OptKntr1.Value = True
        End If
        Tanya = MsgBox("Insured data already exist." + Chr(13) + "Do you want to edit ?", vbQuestion + vbYesNo, "Confirmation")
        If Tanya = vbYes Then
            Aktif
            TxtNo1.Enabled = False: ChkOtomat1.Enabled = False
            CmdTambah.Enabled = False
            CmdUbah.Enabled = True: CmdHapus.Enabled = True
            TxtNama1.SetFocus
        ElseIf Tanya = vbNo Then
            NonAktif
            Bersih
            CmdUbah.Enabled = False: CmdHapus.Enabled = False
            TxtNo1.Enabled = True: ChkOtomat1.Enabled = True
            TxtNo1.SetFocus
        End If
    End If
    End With
    RsTertngg.Close
End If
End Sub

Private Sub TxtPosKntr1_Change()
If Trim(TxtPosKntr1.Text) <> "" Then
   dig$ = Mid(TxtPosKntr1.Text, Len(TxtPosKntr1.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtPosKntr1.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtPosKntr1.Text = digits$
          TxtPosKntr1.SelStart = Len(TxtPosKntr1.Text)
      End If
End If
End Sub

Private Sub TxtPosRmh1_Change()
If Trim(TxtPosRmh1.Text) <> "" Then
   dig$ = Mid(TxtPosRmh1.Text, Len(TxtPosRmh1.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtPosRmh1.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtPosRmh1.Text = digits$
          TxtPosRmh1.SelStart = Len(TxtPosRmh1.Text)
      End If
End If
End Sub

Private Sub TxtTlpKntr1_Change()
If Trim(TxtTlpKntr1.Text) <> "" Then
   dig$ = Mid(TxtTlpKntr1.Text, Len(TxtTlpKntr1.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtTlpKntr1.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtTlpKntr1.Text = digits$
          TxtTlpKntr1.SelStart = Len(TxtTlpKntr1.Text)
      End If
End If
End Sub

Private Sub TxtTlpRmh1_Change()
If Trim(TxtTlpRmh1.Text) <> "" Then
   dig$ = Mid(TxtTlpRmh1.Text, Len(TxtTlpRmh1.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtTlpRmh1.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtTlpRmh1.Text = digits$
          TxtTlpRmh1.SelStart = Len(TxtTlpRmh1.Text)
      End If
End If
End Sub
