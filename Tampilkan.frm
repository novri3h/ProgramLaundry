VERSION 5.00
Begin VB.Form Tampilkan 
   BackColor       =   &H80000009&
   Caption         =   "ESC / Enter = Tutup"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Tampilkan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Unload Me
ElseIf Keyascii = 13 Then
    Call Cetak
End If
End Sub

Function Cetak()
Call BukaDB
RSPesanan.Open "select * from Pesanan Where NomorPsn In(Select Max(NomorPsn)From Pesanan)Order By NomorPsn Desc", Conn
Dim JmlHarga, JmlJual, JmlHasil As Double
Dim MGrs As String
Printer.Font = "Courier New"
Printer.Print
Printer.Print
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSPesanan!Kodeksr & "'", Conn
RSKonsumen.Open "select * From Konsumen where Nomorksm= '" & RSPesanan!NomorKsm & "'", Conn
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print Tab(5); "Nomor      :   "; RSPesanan!NomorPsn
Printer.Print Tab(5); "Tanggal    :   "; Format(RSPesanan!TanggalPsn, "DD-MMMM-YYYY")
Printer.Print Tab(5); "Kasir      :   "; RSKasir!Namaksr
MGrs = String$(33, "-")

Printer.Print Tab(5); "Pemesan    :   "; RSKonsumen!NamaKsm
Printer.Print Tab(5); "Alamat     :   "; RSKonsumen!AlamatKsm
Printer.Print Tab(5); "Telepon    :   "; RSKonsumen!TeleponKsm

Printer.Print Tab(5); MGrs
RSDetailPsn.Open "select * from detailpsn Where NomorPsn='" & RSPesanan!NomorPsn & "'", Conn
RSDetailPsn.MoveFirst
no = 0
Do While Not RSDetailPsn.EOF
    no = no + 1
    Set RSBarang = New ADODB.Recordset
    RSBarang.Open "select * From Barang where Kodebrg= '" & RSDetailPsn!Kodebrg & "'", Conn
    RSBarang.Requery
    Harga = RSBarang!Tarif
    Jumlah = RSDetailPsn!JumlahPsn
    Hasil = Harga * Jumlah
    Printer.Print Tab(5); no; Space(2); RSBarang!NamaBrg
    Printer.Print Tab(10); RKanan(Jumlah, "##"); Space(1); "X";
    Printer.Print Tab(15); Format(Harga, "###,###,###");
    Printer.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailPsn.MoveNext
Loop
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total      :";
Printer.Print Tab(25); RKanan(RSPesanan!TotalHrg, "###,###,###");
Printer.Print Tab(5); "Uang Muka  :";
Printer.Print Tab(25); RKanan(RSPesanan!DP, "###,###,###");

Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Sisa       :";
Printer.Print Tab(25); RKanan(RSPesanan!Sisa, "###,###,###");

Printer.Print Tab(5); MGrs
Printer.Print
Printer.Print
Printer.Print
Conn.Close
Printer.EndDoc
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

