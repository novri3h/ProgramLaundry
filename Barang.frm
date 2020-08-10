VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Barang 
   Caption         =   "Data Barang"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
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
   ScaleHeight     =   3585
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1250
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   4260
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   1250
   End
   Begin VB.CommandButton Cmdinput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   1000
   End
   Begin VB.CommandButton Cmdtutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   1000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Barang.frx":0000
      Height          =   1850
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3254
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Kodebrg"
         Caption         =   "Kode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Namabrg"
         Caption         =   "Nama Barang"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Tarif"
         Caption         =   "Tarif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tarif"
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1005
   End
End
Attribute VB_Name = "Barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBLoundry.mdb"
Adodc1.RecordSource = "barang"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Form_Load()
Text1.MaxLength = 5
Text2.MaxLength = 30
Text3.MaxLength = 8
KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSBarang.Open "Select * From Barang where KodeBrg='" & Text1 & "'", Conn
End Function

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Text3 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdHapus.Caption = "&Hapus"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
    CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSBarang
        If Not RSBarang.EOF Then
            Text2 = RSBarang!NamaBrg
            Text3 = RSBarang!Tarif
        End If
    End With
End Sub

Private Sub CmdRefresh_Click()
    If CmdInput.Caption = "&Simpan" Then
        CmdInput.SetFocus
    ElseIf CmdEdit.Caption = "&Simpan" Then
        CmdEdit.SetFocus
    End If
    Call KondisiAwal
    Form_Activate
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Text1.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Barang (KodeBrg,NamaBrg,Tarif) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "')"
            Conn.Execute SQLTambah
            Call KondisiAwal
            Form_Activate
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If CmdEdit.Caption = "&Edit" Then
        CmdInput.Enabled = False
        CmdEdit.Caption = "&Simpan"
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        Text1.SetFocus
    Else
        If Text2 = "" Or Text3 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Barang Set NamaBrg= '" & Text2 & "', tarif='" & Text3 & "' where KodeBrg='" & Text1 & "'"
            Conn.Execute SQLEdit
            Call KondisiAwal
            Form_Activate
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If CmdHapus.Caption = "&Hapus" Then
        CmdInput.Enabled = False
        CmdEdit.Enabled = False
        CmdTutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Text1.SetFocus
    End If
End Sub

Private Sub CmdTutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Text1) < 5 Then
        MsgBox "Kode Harus 5 Digit"
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
            If Not RSBarang.EOF Then
                TampilkanData
                MsgBox "Kode Barang Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSBarang.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Barang Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSBarang.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Barang where kodebrg= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    Call KondisiAwal
                    Form_Activate
                Else
                    KondisiAwal
                    CmdHapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                Text1.SetFocus
            End If
    End If
End If
End Sub

Private Sub Text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


