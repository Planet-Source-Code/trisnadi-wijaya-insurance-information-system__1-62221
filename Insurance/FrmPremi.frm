VERSION 5.00
Begin VB.Form FrmPremi 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Premium Payment Data"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "FrmPremi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   7455
   Begin VB.ComboBox CboAgen 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Frame FrmJthTempo 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1215
      Left            =   4440
      TabIndex        =   48
      Top             =   3480
      Width           =   2055
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Next Due Date :"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label LblNextTempo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame FrmBayar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment Method"
      Height          =   1575
      Left            =   3840
      TabIndex        =   41
      Top             =   4920
      Width           =   3255
      Begin VB.OptionButton OptBayar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cheque"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptBayar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cash"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox TxtNmBank 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TxtNoCek 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No."
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdBatal 
      BackColor       =   &H00C0C000&
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
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
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   975
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox TxtNoPolis 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Height          =   405
      Left            =   2160
      MaxLength       =   12
      TabIndex        =   4
      ToolTipText     =   "Tekan ENTER untuk lanjutkan"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CheckBox ChkOtomat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Automatic"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TxtNoKw 
      Appearance      =   0  'Flat
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
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   1
      ToolTipText     =   "Press ENTER to continue"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Code"
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp"
      Height          =   255
      Left            =   1800
      TabIndex        =   56
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
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
      Left            =   2160
      TabIndex        =   55
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payment"
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp"
      Height          =   255
      Left            =   5760
      TabIndex        =   53
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label LblMaterai 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   52
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Stamp Duty"
      Height          =   255
      Left            =   4800
      TabIndex        =   51
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LblTglKontrak 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5760
      TabIndex        =   47
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contract Date"
      Height          =   255
      Left            =   5640
      TabIndex        =   46
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label LblJthTempo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5760
      TabIndex        =   45
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   255
      Left            =   5520
      TabIndex        =   44
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REGULAR PREMIUM PAYMENT"
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
      TabIndex        =   40
      Top             =   120
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   6360
      Picture         =   "FrmPremi.frx":0442
      Top             =   120
      Width           =   840
   End
   Begin VB.Label LblDari 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3120
      TabIndex        =   39
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      Height          =   255
      Left            =   2760
      TabIndex        =   38
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label LblKdByr 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   37
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp"
      Height          =   255
      Left            =   1800
      TabIndex        =   36
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp"
      Height          =   255
      Left            =   1800
      TabIndex        =   35
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label LblDenda 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   34
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label LblKdNasabah 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   33
      Top             =   2520
      Width           =   615
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
      Left            =   2160
      TabIndex        =   32
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Fine"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Day(s)"
      Height          =   255
      Left            =   2760
      TabIndex        =   30
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label LblLambat 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Lateness"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Years"
      Height          =   255
      Left            =   2760
      TabIndex        =   27
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label LblKe 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Premium Payment At -"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label LblLama 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Period"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label LblPremi 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Premium Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label LblPeriode 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label LblNamaProd 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label LblKdProd 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Product"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label LblNama 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Policy Holder Name"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Policy No."
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Date"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1245
   End
End
Attribute VB_Name = "FrmPremi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPremi As New ADODB.Recordset
Dim RsPolis As New ADODB.Recordset
Dim RsProduk As New ADODB.Recordset
Dim RsNasabah As New ADODB.Recordset
Dim RsAgen As New ADODB.Recordset
Dim HR As Integer

Private Sub ChkOtomat_Click()
If ChkOtomat.Value = 1 Then
    Penomoran
    TxtNoKw.SetFocus
ElseIf ChkOtomat.Value = 0 Then
    TxtNoKw.Text = ""
    TxtNoKw.SetFocus
End If
End Sub

Private Sub CmdBatal_Click()
TxtNoKw.Text = ""
Bersih
FrmBayar.Enabled = False
CboAgen.Enabled = False
TxtNoPolis.Enabled = False: CmdCari.Enabled = False
ChkOtomat.Value = 0: CmdSimpan.Enabled = False
TxtNoKw.SetFocus
End Sub

Private Sub CmdCari_Click()
FrmCari_Polis.LblTujuan.Caption = Me.Name
FrmCari_Polis.Show 1
End Sub

Private Sub CmdSimpan_Click()
   SQL = "select * from Byr_Premi order by No_kw"
   If RsPremi.State > 0 Then RsPremi.Close
   RsPremi.Open SQL, Cn, adOpenDynamic, adLockOptimistic
   'Save data
   With RsPremi
    .AddNew
    .Fields("No_Kw") = Trim(TxtNoKw.Text)
    .Fields("Tgl_Kw") = CDate(LblTanggal.Caption)
    .Fields("Kd_Agen") = Left(CboAgen.Text, 8)
    .Fields("No_Polis") = Trim(TxtNoPolis.Text)
    .Fields("Byr_Ke") = LblKe.Caption
    .Fields("JTempo") = CDate(LblJthTempo.Caption)
    .Fields("Terlambat") = LblLambat.Caption
    .Fields("Denda") = CCur(LblDenda.Caption)
    .Fields("Next_Tempo") = CDate(LblNextTempo.Caption)
    'If payment use cheque
    If OptBayar(1).Value = True Then
      If (Trim(TxtNoCek.Text) = "") Or (Trim(TxtNmBank.Text) = "") Then
         MsgBox "No. Cek atau Nama Bank tidak boleh kosong.", vbCritical, "Perhatian"
      Else
         .Fields("No_Cek") = Trim(TxtNoCek.Text)
         .Fields("Nama_Bank") = TxtNmBank.Text
      End If
    End If
    .Update
   End With
   RsPremi.Close
   'Change Policy Status
   If (LblLambat.Caption > 30) And (LblLambat <= 60) Then
       If RsPolis.State > 0 Then RsPolis.Close
       SQL2 = "select * from Polis where No_Polis='" & TxtNoPolis.Text & "'"
       RsPolis.Open SQL2, Cn, adOpenDynamic, adLockOptimistic
       If RsPolis.RecordCount > 0 Then
          RsPolis.Fields("Status") = "A"
          RsPolis.Update
          RsPolis.Close
       End If
   End If
   MsgBox "Premium payment saved.", vbInformation, "Information"
   TxtNoKw.Text = ""
   Bersih
   CmdCari.Enabled = False
   ChkOtomat.Value = 0
   CmdSimpan.Enabled = False
   TxtNoKw.SetFocus
End Sub

Private Sub Form_Load()
Left = 50: Top = 50
Koneksi
If RsPremi.State > 0 Then RsPremi.Close
RsPremi.CursorLocation = adUseClient
'Fill Agent Code combobox
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
LblTanggal.Caption = Date
LblMaterai.Caption = Format(CCur(12000), "#,#,#")
TxtNoKw.Text = ""
Bersih
CboAgen.Enabled = False
TxtNoCek.Enabled = False: TxtNmBank.Enabled = False
TxtNoPolis.Enabled = False: FrmBayar.Enabled = False
CmdSimpan.Enabled = False: CmdCari.Enabled = False
End Sub

'Auto numbering for Receipt No.
Private Sub Penomoran()
    Dim Nom As String
    Dim NM As Integer
    SQL = "select * from Byr_Premi order by No_Kw"
    If RsPremi.State > 0 Then RsPremi.Close
    RsPremi.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    If RsPremi.RecordCount = 0 Then
        Nom = "0000001"
    Else
        RsPremi.MoveLast
        NM = Val(Trim(RsPremi.Fields(0))) + 1
        Select Case NM
            Case Is < 10
                Nom = "000000" & NM
            Case Is < 100
                Nom = "00000" & NM
            Case Is < 1000
                Nom = "0000" & NM
            Case Is < 10000
                Nom = "000" & NM
            Case Is < 100000
                Nom = "00" & NM
            Case Is < 1000000
                Nom = "0" & NM
            Case Else
                Nom = NM
            End Select
    End If
    TxtNoKw.Text = Nom
    RsPremi.Close
End Sub

Private Sub Bersih()
LblTanggal.Caption = Date: OptBayar(0).Value = True
CboAgen.Text = ""
TxtNoPolis.Text = "": LblNextTempo.Caption = ""
LblKdNasabah.Caption = "": LblNama.Caption = ""
LblKdProd.Caption = "": LblNamaProd.Caption = ""
LblTglKontrak.Caption = "": LblJthTempo.Caption = ""
LblKdByr.Caption = "": LblPeriode.Caption = ""
LblPremi.Caption = "": LblLama.Caption = ""
LblKe.Caption = "": LblLambat.Caption = ""
LblDenda.Caption = "": LblDari.Caption = ""
TxtNoCek.Text = "": TxtNmBank.Text = ""
LblTotal.Caption = ""
End Sub

'Count payment lateness
Private Sub HitTelat()
sekarang = LblTanggal.Caption
jthtempo = LblJthTempo.Caption
'If lateness on same month and year
If (Month(sekarang) = Month(jthtempo)) And (Year(sekarang) = Year(jthtempo)) Then
    LblLambat.Caption = Day(sekarang) - Day(LblJthTempo.Caption)
'If lateness on different month and same year
ElseIf (Month(sekarang) > Month(jthtempo)) And (Year(sekarang) = Year(jthtempo)) Then
    'Count how many month lateness
    bln = Month(sekarang) - Month(jthtempo)
    'If more than one months lateness
    If bln > 1 Then
        bln = bln - 1
        'Count days every month
        hari = bln * 30
        'Count remaining day on due date
        JumHari (Month(jthtempo))
        sisa = (HR - Day(jthtempo))
        'Retrieve payment day lateness
        LblLambat.Caption = sisa + hari + Day(sekarang)
    'If more than one months lateness
    ElseIf bln = 1 Then
        JumHari (Month(jthtempo))
        'Count remaining day on due date
        sisa = (HR - Day(jthtempo))
        'Retrieve payment day lateness
        LblLambat.Caption = sisa + Day(sekarang)
    End If
'If lateness payment more than one year
Else
    LblLambat.Caption = "0"
End If
'Count Fine
'Fine a day Rp 10.000
LblDenda.Caption = Format(Val(LblLambat.Caption) * 10000, "#,#,#,0")
End Sub

'Count days on a month
Private Sub JumHari(BL As Integer)
Select Case BL
Case 1, 3, 5, 7, 8, 10, 12
    HR = 31
Case 4, 6, 9, 11
    HR = 30
Case 2
    If (Year(Now) Mod 4) = 0 Then
        HR = 29
    Else
        HR = 28
    End If
End Select
End Sub

'Count next due date
Private Sub HitTempo()
If LblKdByr.Caption = "TH" Then
    bulan = Month(LblJthTempo.Caption)
    tahun = Year(LblJthTempo.Caption) + 1
ElseIf LblKdByr.Caption = "SM" Then
    bulan = Month(LblJthTempo.Caption) + 6
    tahun = Year(LblJthTempo.Caption)
End If
'If more 12 months
If bulan > 12 Then
    bulan = bulan - 12
    tahun = Year(LblJthTempo.Caption) + 1
'If lest than 10 months
ElseIf bulan < 10 Then
    bulan = "0" & bulan
End If
    'Retrieve due date
    hari = Day(LblJthTempo.Caption)
    If hari < 10 Then
        hari = "0" & hari
    End If
    tanggal = hari & "/" & bulan & "/" & tahun
    LblNextTempo.Caption = tanggal
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsPremi = Nothing
Set RsPolis = Nothing
Set RsProduk = Nothing
Set RsNasabah = Nothing
Set RsAgen = Nothing
End Sub

Private Sub OptBayar_Click(Index As Integer)
If Index = 0 Then
    TxtNoCek.Text = "": TxtNmBank.Text = ""
    TxtNoCek.Enabled = False: TxtNmBank.Enabled = False
Else
    TxtNoCek.Text = "": TxtNmBank.Text = ""
    TxtNoCek.Enabled = True: TxtNmBank.Enabled = True
End If
End Sub

Private Sub TxtNoCek_Change()
If Trim(TxtNoCek.Text) <> "" Then
   dig$ = Mid(TxtNoPolis.Text, Len(TxtNoCek.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtNoCek.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtNoCek.Text = digits$
          TxtNoCek.SelStart = Len(TxtNoCek.Text)
      End If
End If
End Sub

Private Sub TxtNoKw_Change()
If Trim(TxtNoKw.Text) <> "" Then
   dig$ = Mid(TxtNoKw.Text, Len(TxtNoKw.Text), 1)
   If Asc(dig$) < 48 Or Asc(dig$) > 59 Then
      MsgBox "Only numbers !", vbCritical, "Error"
      For i = 1 To Len(TxtNoKw.Text) - 1
          digits$ = digits$ + digi$
      Next i
          TxtNoKw.Text = digits$
          TxtNoKw.SelStart = Len(TxtNoKw.Text)
      End If
End If
End Sub

Private Sub TxtNoKw_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(TxtNoKw.Text) = "" Then Exit Sub
   If Len(Trim(TxtNoKw.Text)) < 7 Then
      MsgBox "Receipt No. not allow less than 7 characters.", vbCritical, "Attention"
      TxtNoPolis.Enabled = False: CmdCari.Enabled = False
      TxtNoKw.SetFocus
   Else
   SQL1 = "select * from Byr_Premi where No_Kw='" & TxtNoKw.Text & "'"
   If RsPremi.State > 0 Then RsPremi.Close
   RsPremi.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
   If RsPremi.RecordCount = 0 Then
      Bersih
      CboAgen.Enabled = True
      CmdCari.Enabled = True
      CmdCari.SetFocus
   ElseIf RsPremi.RecordCount > 0 Then
      CboAgen.Enabled = False
      CmdCari.Enabled = False
      CmdSimpan.Enabled = False
      'Retrieve premium payment
      With RsPremi
         LblTanggal.Caption = CDate(.Fields("Tgl_Kw"))
         CboAgen.Text = .Fields("Kd_Agen")
         TxtNoPolis.Text = .Fields("No_Polis")
         LblKe.Caption = .Fields("Byr_Ke")
         LblLambat.Caption = .Fields("Terlambat")
         LblJthTempo.Caption = .Fields("JTempo")
         LblDenda.Caption = Format(.Fields("Denda"), "#,#,#,0")
         LblNextTempo.Caption = .Fields("Next_Tempo")
      'Retrieve cheque payment data
      If .Fields("No_Cek") <> "" Then
         OptBayar(1).Value = True
         TxtNoCek.Text = .Fields("No_Cek")
         TxtNmBank.Text = .Fields("Nama_Bank")
      Else
         OptBayar(0).Value = True
         TxtNoCek.Text = ""
         TxtNmBank.Text = ""
      End If
      'Retrieve Agent name
      SQL2 = "select Nama_Agen from Agen where Kd_Agen='" & .Fields("Kd_Agen") & "'"
      If RsAgen.State > 0 Then RsAgen.Close
      RsAgen.Open SQL2, Cn, adOpenDynamic, adLockOptimistic
      Nm_Agen = .Fields("Kd_Agen") & " - " & RsAgen.Fields("Nama_Agen")
      CboAgen.Text = Nm_Agen
      RsAgen.Close
      End With
   End If
   End If
End If
End Sub

Private Sub TxtNoPolis_Change()
'Retrive policy data
If RsPolis.State > 0 Then RsPolis.Close
RsPolis.CursorLocation = adUseClient
SQL2 = "select * from Polis where No_Polis='" & TxtNoPolis.Text & "'"
RsPolis.Open SQL2, Cn, adOpenDynamic, adLockOptimistic
If RsPolis.RecordCount > 0 Then
  With RsPolis
    LblKdNasabah.Caption = .Fields("No_PolHolder")
    CboAgen.Text = .Fields("Kd_Agen")
    LblKdProd.Caption = .Fields("Kd_Prod")
    LblLama.Caption = .Fields("Lama")
    LblKdByr.Caption = .Fields("Periode_Byr")
    LblPremi.Caption = Format(.Fields("Premi"), "#,#,#,#,0")
    LblTglKontrak.Caption = .Fields("Tgl_Kontrak")
    If LblKdByr.Caption = "AN" Then
        LblPeriode.Caption = "Annually"
        LblDari.Caption = LblLama.Caption * 1
    ElseIf LblKdByr.Caption = "SA" Then
        LblPeriode.Caption = "Semi Annually"
        LblDari.Caption = LblLama.Caption * 2
    End If
  End With
End If
RsPolis.Close
'Retrieve Policy Holder name
If RsNasabah.State > 0 Then RsNasabah.Close
RsNasabah.CursorLocation = adUseClient
SQL3 = "select Nama from Nasabah2 where No_PolHolder='" & LblKdNasabah.Caption & "'"
RsNasabah.Open SQL3, Cn, adOpenDynamic, adLockOptimistic
If RsNasabah.RecordCount > 0 Then
    LblNama.Caption = RsNasabah.Fields("Nama")
End If
RsNasabah.Close
'Retrieve Insurance Product Name
If RsProduk.State > 0 Then RsProduk.Close
RsProduk.CursorLocation = adUseClient
SQL4 = "select Deskripsi from Produk where Kd_Prod='" & LblKdProd.Caption & "'"
RsProduk.Open SQL4, Cn, adOpenDynamic, adLockOptimistic
If RsProduk.RecordCount > 0 Then
    LblNamaProd.Caption = RsProduk.Fields("Deskripsi")
End If
RsProduk.Close
'Retrieve due date
If RsPremi.State > 0 Then RsPremi.Close
SQL1 = "select * from Byr_Premi where No_Polis='" & TxtNoPolis.Text & "'"
RsPremi.Open SQL1, Cn, adOpenDynamic, adLockOptimistic
'If first premium payment
If RsPremi.RecordCount = 0 Then
    LblJthTempo.Caption = LblTglKontrak.Caption
    LblKe.Caption = "1"
Else
    RsPremi.MoveLast
    LblJthTempo.Caption = RsPremi.Fields("Next_Tempo")
    'Retrieve Premium Payment at-
    LblKe.Caption = Val(RsPremi.Fields("Byr_Ke")) + 1
End If
If RsPolis.State > 0 Then RsPolis.Close
SQL4 = "select Status,Cara_Byr from Polis where No_Polis='" & TxtNoPolis.Text & "'"
RsPolis.Open SQL4, Cn, adOpenDynamic, adLockOptimistic
If RsPolis.RecordCount > 0 Then
'Check Policy status
If RsPolis.Fields("Status") = "E" Then
    MsgBox "Expired policy.", vbCritical, "Attention"
    FrmBayar.Enabled = False
    CmdSimpan.Enabled = False
    Exit Sub
'Check Premium Payment method
ElseIf (RsPolis.Fields("Status") = "A") And (RsPolis.Fields("Cara_Byr") <> "T") Then
    MsgBox "Premium payment method not cash.", vbCritical, "Attention"
    FrmBayar.Enabled = False
    CmdSimpan.Enabled = False
    Exit Sub
Else
    'Count Lateness
    HitTelat
    'Retrieve Next Due Date
    HitTempo
    'Count Total Payment
    If LblDenda.Caption = "" Then
      Tot = CCur(LblPremi.Caption) + CCur(LblMaterai.Caption)
    Else
      Tot = CCur(LblPremi.Caption) + CCur(LblDenda.Caption) + CCur(LblMaterai.Caption)
    End If
      LblTotal.Caption = Format(Tot, "#,#,#,#,0")
    If CDate(LblTanggal.Caption) >= CDate(LblJthTempo.Caption) Then
        FrmBayar.Enabled = True
        CmdSimpan.Enabled = True
    Else
        FrmBayar.Enabled = False
        CmdSimpan.Enabled = False
    End If
End If
End If
End Sub
