VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Menu 
   Caption         =   "Menu Utama"
   ClientHeight    =   3390
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5820
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   10650
   ScaleWidth      =   20160
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   1560
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnkasir 
         Caption         =   "Kasir"
      End
      Begin VB.Menu mnbarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnkurir 
         Caption         =   "Kurir"
      End
      Begin VB.Menu mnkonsumen 
         Caption         =   "Konsumen"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mncucian 
         Caption         =   "Penyerahan Cucian"
      End
      Begin VB.Menu mnpengiriman 
         Caption         =   "Pengiriman Cucian"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlapbarang 
         Caption         =   "Data Barang"
      End
      Begin VB.Menu mnlappesanan 
         Caption         =   "Rincian Pemesanan"
      End
      Begin VB.Menu mnlappengiriman 
         Caption         =   "Rincian Pengiriman"
      End
      Begin VB.Menu mnakumpsn 
         Caption         =   "Akumulasi Pemesanan"
      End
      Begin VB.Menu mnkumkrm 
         Caption         =   "Akumulasi Pengiriman"
      End
   End
   Begin VB.Menu mnjejak 
      Caption         =   "Jejak Transaksi"
      Begin VB.Menu mnjejakpsn 
         Caption         =   "Pemesanan"
      End
      Begin VB.Menu mnjejakkrm 
         Caption         =   "Pengiriman"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub

Private Sub mnakumpsn_Click()
lapAkumPsn.Show
End Sub

Private Sub mnbarang_Click()
Barang.Show
End Sub

Private Sub mncucian_Click()
Cucian.Show
End Sub

Private Sub mnjejakkrm_Click()
DataKrm.Show
End Sub

Private Sub mnjejakpsn_Click()
DataPsn.Show
End Sub

Private Sub mnkasir_Click()
Kasir.Show
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnkonsumen_Click()
Konsumen.Show vbModal
End Sub

Private Sub mnkumkrm_Click()
lapAkumKrm.Show
End Sub

Private Sub mnkurir_Click()
Kurir.Show
End Sub

Private Sub mnlapbarang_Click()
CR.ReportFileName = App.Path & "\Lap Barang.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub mnlappengiriman_Click()
LapKirim.Show
End Sub

Private Sub mnlappesanan_Click()
LapPesan.Show
End Sub

Private Sub mnpengiriman_Click()
Pengiriman.Show
End Sub

Private Sub mnpesanan_Click()
Pesanan.Show
End Sub

Private Sub mnujisql_Click()
UjiSQL.Show
End Sub

Private Sub nlaporan_Click()

End Sub

Private Sub MNSQL_Click()
UjiSQL.Show
End Sub
