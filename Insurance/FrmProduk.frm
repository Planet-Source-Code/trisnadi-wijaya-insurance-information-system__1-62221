VERSION 5.00
Begin VB.Form FrmProduk 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insurance Product"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "FrmProduk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4350
   Begin VB.TextBox TxtKdProd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   0
      ToolTipText     =   "Press ENTER to continue"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox TxtNama 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1815
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
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
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
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
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
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   855
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
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
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Produk Name"
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
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "FrmProduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsProduk As New ADODB.Recordset

Private Sub CmdBatal_Click()
TxtKdProd.Text = ""
Bersih
NonAktif
CmdTambah.Enabled = False
CmdUbah.Enabled = False
CmdHapus.Enabled = False
TxtKdProd.SetFocus
End Sub

Private Sub CmdHapus_Click()
SQL = "select * from Produk where Kd_Prod='" & Trim(TxtKdProd.Text) & "'"
If RsProduk.State > 0 Then RsProduk.Close
RsProduk.Open SQL, Cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Are you sure want to delete this product ?", vbQuestion + vbYesNo, "Konfirmasi Hapus")
If Tanya = vbYes Then
    RsProduk.Delete
    MsgBox "Product data deleted.", vbInformation, "Information"
    RsProduk.Close
End If
Bersih
CmdUbah.Enabled = False
CmdHapus.Enabled = False
TxtKdProd.SetFocus
End Sub

Private Sub CmdTambah_Click()
    If RsProduk.State > 0 Then RsProduk.Close
    'Save data
    RsProduk.Open "Produk", Cn, adOpenDynamic, adLockOptimistic
    With RsProduk
        .AddNew
        Simpan
        .Update
        .Close
    End With
    MsgBox "Product data saved.", vbInformation, "Information"
    TxtKdProd.Text = ""
    Bersih
    CmdTambah.Enabled = False
    TxtKdProd.SetFocus
End Sub

Private Sub CmdUbah_Click()
    'Edit data
    SQL = "select * from Produk where Kd_Prod='" & Trim(TxtKdProd.Text) & "'"
    If RsProduk.State > 0 Then RsProduk.Close
    RsProduk.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsProduk
        Simpan
        .Update
        .Close
    End With
    MsgBox "Product data changed.", vbInformation, "Information"
    TxtKdProd.Text = ""
    Bersih
    CmdUbah.Enabled = False
    CmdHapus.Enabled = False
    TxtKdProd.SetFocus
End If
End Sub

Private Sub Form_Activate()
TxtKdProd.SetFocus
End Sub

Private Sub Form_Load()
Left = 600: Top = 600
TxtKdProd.Text = "": TxtNama.Text = ""
NonAktif
CmdTambah.Enabled = False
CmdUbah.Enabled = False
CmdHapus.Enabled = False
Koneksi
RsProduk.CursorLocation = adUseClient
End Sub

Private Sub Aktif()
TxtNama.Enabled = True
End Sub

Private Sub NonAktif()
TxtNama.Enabled = False
End Sub

Private Sub Bersih()
TxtNama.Text = ""
End Sub

Private Sub Simpan()
With RsProduk
    .Fields("Kd_Prod") = Trim(TxtKdProd.Text)
    .Fields("Deskripsi") = Trim(TxtNama.Text)
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsProduk = Nothing
End Sub

Private Sub TxtKdProd_Change()
TxtKdProd.Text = UCase(TxtKdProd.Text)
TxtKdProd.SelStart = Len(TxtKdProd.Text)
End Sub

Private Sub TxtKdProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Trim(TxtKdProd.Text) = "" Then Exit Sub
    SQL = "select * from Produk where Kd_Prod='" & Trim(TxtKdProd.Text) & "'"
    If RsProduk.State > 0 Then RsProduk.Close
    RsProduk.Open SQL, Cn, adOpenDynamic, adLockOptimistic
    With RsProduk
        If .RecordCount = 0 Then
            Aktif
            CmdTambah.Enabled = True
            TxtNama.SetFocus
        Else
            'Retrieve data
            TxtNama.Text = Trim(.Fields("Deskripsi"))
            Tanya = MsgBox("Product data already exist." + Chr(13) + "Do you want to edit ?", vbQuestion + vbYesNo, "Confirmation")
            If Tanya = vbYes Then
                Aktif
                CmdUbah.Enabled = True
                CmdHapus.Enabled = True
                TxtNama.SetFocus
            ElseIf Tanya = vbNo Then
                Bersih
                TxtKdProd.Text = ""
                TxtKdProd.SetFocus
            End If
        End If
    End With
    RsProduk.Close
End If
End Sub

