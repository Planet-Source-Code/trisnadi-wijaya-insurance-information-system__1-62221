Attribute VB_Name = "Module1"
Option Explicit
Public Cn As New ADODB.Connection

Public Sub Koneksi()
If Cn.State > 0 Then Exit Sub
With Cn
    .ConnectionTimeout = 30
    .ConnectionString = "Provider=SQLOLEDB.1;" & _
                        "User ID=trisnadi; Password=trisnadi;" & _
                        "InitialCatalog=Insurance"
    .Open
End With
End Sub

