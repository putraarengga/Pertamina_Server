Imports Spire.Barcode
Imports System.Data
Imports System.Data.Odbc
Imports System.ComponentModel
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Frm_Main
    Dim databaru As Boolean
    Dim selectDataBase As String
    Shared Property indexKendaraan As String
    Shared Property indexTujuan As String
    Dim counter As Integer
    Dim pesan, simpan, TextToPrint As String
    Dim IDTujuan, IDUser, IDKendaraan, IDDistribusi As Integer

    Private Sub Frm_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\icons\IDSF.ico")
        Me.Icon = img
        'If (FormMenu.dbOnline = True) Then
        '    If CheckNewData() Then
        '        SynchronizeDB()
        '        ClearLocalDB()
        '    End If
        'End If
        GetCounter()
        IsiGrid()
        databaru = False
        PrintDocument1.PrinterSettings.PrinterName = "Xprinter XP-350B II"
       

    End Sub

    Public Function CheckNewData() As Boolean
        Dim newData As Boolean = False

        selectDataBase = "SELECT * FROM tdistribusi"
        bukaDBoffline()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DT = New DataTable
        DT.Clear()

        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            newData = True
        End If
        Return newData
    End Function
    Sub SynchronizeDB()

    End Sub
    Sub ClearLocalDB()

    End Sub

    Sub IsiGrid()
        Dim tanggalSekarang As String
        tanggalSekarang = Format(DateTime.Now, "yyyy-MM-dd")
        selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO,tdatatujuan.namaTujuan,tkendaraan.noPolKendaraan,tdistribusi.Keterangan,tdatauser.NamaUser,tdistribusi.dataBarcode, " +
                        " tdistribusi.tglMuat,tdistribusi.wktMuat,tdistribusi.wktSampai,tdistribusi.tglSampai,tkendaraan.callCenter,tkendaraan.kapasitasTruk, tdistribusi.tempatLoading " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                        " JOIN tdatatujuan ON tdatatujuan.IDTujuan = tdistribusi.IDTujuan " +
                        " JOIN tdatauser ON tdistribusi.IDUser = tdatauser.IDUser WHERE tdistribusi.tglMuat='" + tanggalSekarang + "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = DS.Tables("tdistribusi")
        DataGridView1.Sort(DataGridView1.Columns(6), ListSortDirection.Descending)
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Transportir"
            .Columns(1).HeaderCell.Value = "Nomor DO"
            .Columns(2).HeaderCell.Value = "Tempat Tujuan"
            .Columns(3).HeaderCell.Value = "No Pol Kendaraan"
            .Columns(4).HeaderCell.Value = "Keterangan"
            .Columns(5).HeaderCell.Value = "User Server"
            .Columns(6).HeaderCell.Value = "Barcode"
            .Columns(7).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(8).HeaderCell.Value = "Waktu Pengiriman"
            .Columns(9).HeaderCell.Value = "Tanggal Sampai"
            .Columns(10).HeaderCell.Value = "Waktu Sampai"
            .Columns(11).HeaderCell.Value = "Call Center"
            .Columns(12).HeaderCell.Value = "Liter"
            .Columns(13).HeaderCell.Value = "Tempat Loading"

        End With
        DataGridView1.Enabled = True
    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Form2.Show()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
       
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles btn_6.Click
        databaru = True
        Bersih()
        GetCounter()
        If counter > 999 Then
            tb_9.Text = Format(DateTime.Now, "yyyyMMdd") & counter
        ElseIf counter > 99 Then
            tb_9.Text = Format(DateTime.Now, "yyyyMMdd") & "0" & counter
        ElseIf counter > 9 Then
            tb_9.Text = Format(DateTime.Now, "yyyyMMdd") & "00" & counter
        ElseIf counter >= 0 Then
            tb_9.Text = Format(DateTime.Now, "yyyyMMdd") & "000" & counter
        End If

        If FormMenu.dbOnline = False Then
            tb_9.Text = "001" & tb_9.Text
        End If



        btn_7.Enabled = False
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Bersih()
        btn_4.Enabled = True
        isitextbox(e.RowIndex)
        btn_7.Enabled = True
        databaru = False
        GetIDDistribusi()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles btn_4.Click
        Dim generator As New BarCodeGenerator(BarCodeControl1)
        Dim barcode As Image
        barcode = generator.GenerateImage()

        'save the barcode as an image
        Try
            barcode.Save("barcode.png")
            barcode.Dispose()
            barcode = Nothing
        Catch ex As Exception

        End Try
        
        PrintHeader()
        ItemsToBePrinted2()
        Dim printControl = New Printing.StandardPrintController

        PrintDocument1.PrintController = printControl
        Try
            PrintDocument1.Print()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub GroupBox8_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub isitextbox(ByVal x As Integer)
        Try

            tb_1.Text = DataGridView1.Rows(x).Cells(3).Value
            tb_5.Text = DataGridView1.Rows(x).Cells(0).Value
            tb_6.Text = DataGridView1.Rows(x).Cells(1).Value
            tb_9.Text = DataGridView1.Rows(x).Cells(6).Value
            tb_7.Text = DataGridView1.Rows(x).Cells(2).Value

            indexKendaraan = tb_1.Text
            indexTujuan = tb_7.Text
            GetDataKendaraan()
            GetDataTujuan()

            If DataGridView1.Rows(x).Cells(5).Value.ToString = "" Then
                DateTimePicker1.CustomFormat = " "  'An empty SPACE
                DateTimePicker1.Format = DateTimePickerFormat.Custom
            Else
                'DateTimePicker1.CustomFormat = "dd/MM/yyyy"
                DateTimePicker1.Value = DataGridView1.Rows(x).Cells(5).Value
                DateTimePicker1.Format = DateTimePickerFormat.Custom
                'DateTimePicker1.CustomFormat = "yyyy-MM-dd"
                'vdate = DateTimePicker1.Value.Year + "-" + DateTimePicker1.Value.Month + "-" + DateTimePicker1.Value.Day
            End If

        Catch ex As Exception
        End Try
    End Sub

    Sub Bersih()
        'tb_1.Enabled = True
        'tb_2.Enabled = True
        'tb_3.Enabled = True
        'tb_4.Enabled = True
        'tb_5.Enabled = True
        tb_6.Enabled = True
        'tb_7.Enabled = True
        'tb_8.Enabled = True
        tb_9.Enabled = True
        btn_1.Enabled = True
        btn_2.Enabled = True
        btn_3.Enabled = True
        btn_5.Enabled = True
        DateTimePicker1.Enabled = True
        DateTimePicker2.Enabled = True

        tb_1.Text = ""
        tb_2.Text = ""
        tb_3.Text = ""
        tb_4.Text = ""
        tb_5.Text = ""
        tb_6.Text = ""
        tb_7.Text = ""
        tb_8.Text = ""
        tb_9.Text = " "
        DateTimePicker1.Value = DateTime.Now
        DateTimePicker2.Value = DateTime.Now
        DateTimePicker3.Value = DateTime.Now
        DateTimePicker4.Value = DateTime.Now
        tb_5.Focus()


    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub


    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btn_1.Click
        formDataKendaraan.Button6.Visible = True
        formDataKendaraan.Show()
        formDataKendaraan.Focus()
    End Sub

    Sub GetDataKendaraan()

        selectDataBase = "SELECT * FROM tkendaraan WHERE noPolKendaraan='" & indexKendaraan & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            tb_1.Text = DT.Rows(0).Item("noPolKendaraan")
            tb_2.Text = DT.Rows(0).Item("namaSopir")
            tb_3.Text = DT.Rows(0).Item("namaKernet")
            tb_4.Text = DT.Rows(0).Item("kapasitasTruk")
            tb_5.Text = DT.Rows(0).Item("namaPerusahaan")
        End If
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btn_2.Click
        formDataTujuan.Button6.Visible = True
        formDataTujuan.Show()
        formDataTujuan.Focus()
        GetDataTujuan()


    End Sub

    Sub GetDataTujuan()
        selectDataBase = "SELECT * FROM tdatatujuan WHERE namaTujuan='" & indexTujuan & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            tb_7.Text = DT.Rows(0).Item("namaTujuan")
            tb_8.Text = DT.Rows(0).Item("alamatTujuan")
            IDTujuan = DT.Rows(0).Item("IDTujuan")
        End If
    End Sub


    Sub GetIDDistribusi()
        selectDataBase = "SELECT * FROM tdistribusi WHERE NoDO = '" & tb_6.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            IDDistribusi = DT.Rows(0).Item("IDDistribusi")
        End If
    End Sub

    Sub GetCounter()
        selectDataBase = "SELECT * FROM tdistribusi ORDER BY IDDistribusi DESC LIMIT 1"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            counter = DT.Rows(0).Item("IDDistribusi") + 1
        End If
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles tb_5.TextChanged

    End Sub

    Private Sub BarCodeControl1_Click(sender As Object, e As EventArgs) Handles BarCodeControl1.Click

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles tb_9.TextChanged
        If tb_9.Text = "" Then
            tb_9.Text = " "
        End If

        BarCodeControl1.Data = tb_9.Text
        BarCodeControl1.Data2D = tb_9.Text


    End Sub
    Private Sub GetIDKendaraan()
        selectDataBase = "SELECT * FROM tkendaraan WHERE noPolKendaraan='" & tb_1.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            IDKendaraan = DT.Rows(0).Item("IDKendaraan")
        End If
    End Sub
    Private Sub GetIDTujuan()
        selectDataBase = "SELECT * FROM tdatatujuan WHERE namaTujuan='" & tb_7.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            IDTujuan = DT.Rows(0).Item("IDTujuan")
        End If
    End Sub

    Private Sub btn_3_Click(sender As Object, e As EventArgs) Handles btn_3.Click
        Dim tmpWaktu, tmpTanggal As String
        tmpWaktu = DateTimePicker1.Value.ToString("hh:mm:00")
        tmpTanggal = DateTimePicker2.Value.ToString("yyyy-MM-dd")
        GetIDKendaraan()
        GetIDTujuan()
        If tb_5.Text = "" Or tb_6.Text = "" Or tb_7.Text = "" Or tb_9.Text = "" Then
            MsgBox("Tidak bisa menyimpan data ke server. (Data Kurang Lengkap)")
            Return

        End If

        If databaru Then
            pesan = MsgBox("Apakah anda yakin data ini akan ditambah ke database?", MsgBoxStyle.YesNo, "IDSF")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "INSERT INTO tdistribusi(IDKendaraan,IDUser, " +
                        "IDTujuan,NoDO,wktMuat,tglMuat, " +
                        " dataBarcode,Keterangan) " +
                     "VALUES ('" & IDKendaraan & "','" & FormMenu.idUser & "'" +
                        ",'" & IDTujuan & "','" & tb_6.Text & "','" & tmpWaktu & "','" & tmpTanggal & "'" +
                        ",'" & tb_9.Text & "','REGISTERED')"
        Else
            pesan = MsgBox("Apakah anda yakin data ini akan diupdate ke database?", MsgBoxStyle.YesNo, "IDSF")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If
            simpan = "UPDATE tdistribusi SET IDKendaraan = '" & IDKendaraan & "',IDUser = '" & FormMenu.idUser & "', " +
                       " IDTujuan = '" & IDTujuan & "',NoDO='" & tb_6.Text & "',wktMuat='" & tmpWaktu & "',tglMuat='" & tmpTanggal & "', " +
                       " dataBarcode='" & tb_9.Text & "' " +
                      " WHERE IDDistribusi = '" & IDDistribusi & "'"

            'simpan = "UPDATE tdatatujuan SET namaTujuan= '" & TextBox4.Text & "', alamatTujuan = '" & TextBox5.Text & "' WHERE IDTujuan= '" & TextBox2.Text & "' "
        End If
        jalankansql(simpan)
        DataGridView1.Refresh()
        IsiGrid()
        Bersih()
        btn_7.Enabled = False

    End Sub
    Private Sub jalankansql(ByVal sQL As String)
        Dim objcmd As New System.Data.Odbc.OdbcCommand
        bukaDB()
        Try
            objcmd.Connection = konek
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQL
            objcmd.ExecuteNonQuery()
            objcmd.Dispose()
            MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub

    Private Sub btn_5_Click(sender As Object, e As EventArgs) Handles btn_5.Click
        Bersih()
        GetCounter()
        IsiGrid()
    End Sub

    Private Sub btn_7_Click(sender As Object, e As EventArgs) Handles btn_7.Click
        Dim hapussql As String
        Dim pesan As String
        pesan = MsgBox("Apakah anda yakin untuk menghapus data pada server? ", vbExclamation + MsgBoxStyle.YesNo, "IDSF")
        If pesan = MsgBoxResult.No Then Exit Sub

        hapussql = "DELETE FROM tdistribusi WHERE NoDO ='" & tb_6.Text & "'"
        If tb_6.Text = "" Then Exit Sub
        jalankansql(hapussql)
        DataGridView1.Refresh()
        IsiGrid()
        Bersih()

    End Sub
    Public Sub PrintHeader()
        TextToPrint = " "
        TextToPrint &= Environment.NewLine
        TextToPrint &= Environment.NewLine


        TextToPrint &= Environment.NewLine
        Dim StringToPrint As String = "Untuk Pelanggan"
        Dim LineLen As Integer = StringToPrint.Length
        Dim spcLen1 As New String(" "c, Math.Round((17 - LineLen))) 'This line is used to center text in the middle of the receipt
        TextToPrint &= spcLen1 & StringToPrint & Environment.NewLine


        TextToPrint &= Environment.NewLine
        StringToPrint = "Standard-PLN"
        LineLen = StringToPrint.Length
        Dim spcLen2 As New String(" "c, Math.Round((14 - LineLen)))
        TextToPrint &= spcLen2 & StringToPrint & Environment.NewLine

        StringToPrint = "2601-Kota Jayapura"
        LineLen = StringToPrint.Length
        Dim spcLen3 As New String(" "c, Math.Round((20 - LineLen)))
        TextToPrint &= spcLen3 & StringToPrint & Environment.NewLine
        TextToPrint &= Environment.NewLine

        StringToPrint = "SURAT PENGANTAR PENGIRIMAN"
        LineLen = StringToPrint.Length
        Dim spcLen4 As New String(" "c, Math.Round((33 - LineLen)))
        TextToPrint &= spcLen4 & StringToPrint & Environment.NewLine

        StringToPrint = "=================================================="
        LineLen = StringToPrint.Length
        Dim spcLen10 As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen10 & StringToPrint & Environment.NewLine
    End Sub

    Public Sub ItemsToBePrinted1()

        TextToPrint &= " "
        Dim globalLengt As Integer = 0

        Dim NamaPerusahaan As String = tb_5.Text.ToString()
        Dim NomorDO As String = tb_6.Text.ToString()
        Dim Nopol As String = tb_1.Text.ToString()
        Dim NamaSopir As String = tb_2.Text.ToString()
        Dim NamaKernet As String = tb_3.Text.ToString()
        Dim Kapasitas As String = tb_4.Text.ToString()
        Dim Tujuan As String = tb_7.Text.ToString()
        Dim AlamatTujuan As String = tb_8.Text.ToString()
        Dim Barcode As String = tb_9.Text.ToString()
        Dim TanggalPengiriman As String = DateTimePicker1.Text.ToString()
        Dim WaktuPengiriman As String = DateTimePicker2.Text.ToString()

        Dim StringToPrint As String = "No.Pol.Kendaraan      :"
        Dim StringToPrint2 As String = Nopol
        Dim LineLen As String = StringToPrint.Length
        Dim LineLen2 As String = StringToPrint2.Length
        globalLengt = StringToPrint.Length
        Dim spcLen1 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen1b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen1 & StringToPrint & spcLen1b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Shippment No           :"
        StringToPrint2 = IDDistribusi
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen2 As New String(" "c, Math.Round(24 - LineLen))
        Dim spcLen2b As New String(" "c, Math.Round(1))
        TextToPrint &= spcLen2 & StringToPrint & spcLen2b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Nama Pengemudi         :"
        StringToPrint2 = NamaSopir
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen3 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen3b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen3 & StringToPrint & spcLen3b & StringToPrint2 & Environment.NewLine


        StringToPrint = "Tujuan                 :"
        StringToPrint2 = Tujuan
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen4 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen4b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen4 & StringToPrint & spcLen4b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Nomor DO               :"
        StringToPrint2 = NomorDO
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen5b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen5 & StringToPrint & spcLen5b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Kapasitas Tangki       :"
        StringToPrint2 = Kapasitas
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen6 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen6b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen6 & StringToPrint & spcLen6b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Tanggal Pengiriman     :"
        StringToPrint2 = TanggalPengiriman
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen7 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen7b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen7 & StringToPrint & spcLen7b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Waktu Pengiriman       :"
        StringToPrint2 = WaktuPengiriman
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen8 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen8b As New String(" "c, Math.Round((2)))
        TextToPrint &= spcLen8 & StringToPrint & spcLen8b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Nomor Barcode          :"
        StringToPrint2 = Barcode
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen9 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen9b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen9 & StringToPrint & spcLen9b & StringToPrint2 & Environment.NewLine

        StringToPrint = "=================================================="
        LineLen = StringToPrint.Length
        Dim spcLen10 As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen10 & StringToPrint & Environment.NewLine

    End Sub
    Public Sub ItemsToBePrinted2()

        TextToPrint &= " "
        Dim globalLengt As Integer = 0

        Dim NamaPerusahaan As String = tb_5.Text.ToString()
        Dim NomorDO As String = tb_6.Text.ToString()
        Dim Nopol As String = tb_1.Text.ToString()
        Dim NamaSopir As String = tb_2.Text.ToString()
        Dim NamaKernet As String = tb_3.Text.ToString()
        Dim Kapasitas As String = tb_4.Text.ToString()
        Dim Tujuan As String = tb_7.Text.ToString()
        Dim AlamatTujuan As String = tb_8.Text.ToString()
        Dim Barcode As String = tb_9.Text.ToString()
        Dim TanggalPengiriman As String = DateTimePicker1.Text.ToString()
        Dim WaktuPengiriman As String = DateTimePicker2.Text.ToString()

        Dim StringToPrint As String
        Dim StringToPrint2 As String
        Dim LineLen As String
        Dim LineLen2 As String


        StringToPrint = "Tanggal Pengiriman     :"
        StringToPrint2 = TanggalPengiriman
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen7 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen7b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen7 & StringToPrint & spcLen7b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Kapasitas Tangki       :"
        StringToPrint2 = Kapasitas
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen8 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen8b As New String(" "c, Math.Round((2)))
        TextToPrint &= spcLen8 & StringToPrint & spcLen8b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Tujuan Pengiriman    :"
        StringToPrint2 = Tujuan
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen9 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen9b As New String(" "c, Math.Round((2)))
        TextToPrint &= spcLen9 & StringToPrint & spcLen9b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Nomor DO               :"
        StringToPrint2 = NomorDO
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen10 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen10b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen10 & StringToPrint & spcLen10b & StringToPrint2 & Environment.NewLine

        StringToPrint = "Nomor Polisi Kendaraan :"
        StringToPrint2 = Nopol
        LineLen = StringToPrint.Length
        LineLen2 = StringToPrint2.Length()
        globalLengt = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen11 As New String(" "c, Math.Round((24 - LineLen)))
        Dim spcLen11b As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen11 & StringToPrint & spcLen11b & StringToPrint2 & Environment.NewLine

        StringToPrint = "=================================================="
        LineLen = StringToPrint.Length
        Dim spcLen12 As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen12 & StringToPrint & Environment.NewLine

    End Sub

    Public Sub printFooter()

        TextToPrint &= Environment.NewLine & Environment.NewLine & Environment.NewLine & Environment.NewLine
        Dim globalLengt As Integer = 0

        TextToPrint &= Environment.NewLine & Environment.NewLine
        Dim StringToPrint As String = "TERIMA KASIH ATAS KEPERCAYAAN ANDA "
        Dim LineLen As String = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((1)))
        TextToPrint &= spcLen5 & StringToPrint

        StringToPrint = "MENGGUNAKAN PRODUK PERTAMINA"
        LineLen = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen6 As New String(" "c, Math.Round((5)))
        TextToPrint &= Environment.NewLine & spcLen6 & StringToPrint & Environment.NewLine

    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Static currentChar As Integer
        Dim textfont As Font = New Font("Courier New", 10, FontStyle.Bold)

        Dim h, w As Integer
        Dim left, top As Integer
        With PrintDocument1.DefaultPageSettings
            h = 0
            w = 0
            left = 0
            top = 0
        End With


        Dim lines As Integer = CInt(Math.Round(h / 1))
        Dim b As New Rectangle(left, top, w, h)
        Dim format As StringFormat
        format = New StringFormat(StringFormatFlags.LineLimit)
        Dim line, chars As Integer
        Dim appPath As String = Application.StartupPath()
        Dim newImage As Image = Image.FromFile(appPath + "\logoPln.png")
        Dim newImage2 As Image = Image.FromFile(appPath + "\barcode.png")

        ' Create Point for upper-left corner of image.
        Dim ulCorner As New Point(180, 20)
        Dim ulCorner2 As New Point(35, 230)
        If FormMenu.dbOnline = False Then
            ulCorner2 = New Point(25, 230)
        End If

        ' Draw image to screen.
        e.Graphics.DrawImage(newImage, ulCorner)
        e.Graphics.DrawImage(newImage2, ulCorner2)
        'printFooter()
        e.Graphics.MeasureString(Mid(TextToPrint, currentChar + 1), textfont, New SizeF(w, h), format, chars, line)
        e.Graphics.DrawString(TextToPrint.Substring(currentChar, chars), New Font("Courier New", 9, FontStyle.Bold), Brushes.Black, b, format)


        currentChar = currentChar + chars
        If currentChar < TextToPrint.Length Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            currentChar = 0
        End If
    End Sub



    Private Sub btn_9_Click(sender As Object, e As EventArgs) Handles btn_9.Click
        Dim tanggalAwal, tanggalAkhir As String
        bukaDB()
        DateTimePicker5.Format = DateTimePickerFormat.Custom
        DateTimePicker6.Format = DateTimePickerFormat.Custom

        DateTimePicker5.CustomFormat = "dd/MMM/yyyy"
        DateTimePicker6.CustomFormat = "dd/MMM/yyyy"

        tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-dd")
        tanggalAkhir = Format(DateTimePicker6.Value.Date, "yyyy-MM-dd")
        If tanggalAwal = tanggalAkhir Then
            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO,tdatatujuan.namaTujuan,tkendaraan.noPolKendaraan,tdistribusi.Keterangan,tdatauser.NamaUser,tdistribusi.dataBarcode, " +
                        " tdistribusi.tglMuat,tdistribusi.wktMuat,tdistribusi.wktSampai,tdistribusi.tglSampai,tkendaraan.callCenter,tkendaraan.kapasitasTruk , tdistribusi.tempatLoading " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                        " JOIN tdatatujuan ON tdatatujuan.IDTujuan = tdistribusi.IDTujuan " +
                        " JOIN tdatauser ON tdatauser.IDUser = tdistribusi.IDUser WHERE tdistribusi.tglMuat LIKE '%" & tanggalAkhir & "%'"
            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        Else
            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO,tdatatujuan.namaTujuan,tkendaraan.noPolKendaraan,tdistribusi.Keterangan,tdatauser.NamaUser,tdistribusi.dataBarcode, " +
                        " tdistribusi.tglMuat,tdistribusi.wktMuat,tdistribusi.wktSampai,tdistribusi.tglSampai,tkendaraan.callCenter,tkendaraan.kapasitasTruk , tdistribusi.tempatLoading " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                        " JOIN tdatatujuan ON tdatatujuan.IDTujuan = tdistribusi.IDTujuan " +
                        " JOIN tdatauser ON tdatauser.IDUser = tdistribusi.IDUser WHERE tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'"
            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)

        End If
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = (DS.Tables("tdistribusi"))
        DataGridView1.Sort(DataGridView1.Columns(6), ListSortDirection.Descending)
        DataGridView1.Enabled = True
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Transportir"
            .Columns(1).HeaderCell.Value = "Nomor DO"
            .Columns(2).HeaderCell.Value = "Tempat Tujuan"
            .Columns(3).HeaderCell.Value = "No Pol Kendaraan"
            .Columns(4).HeaderCell.Value = "Keterangan"
            .Columns(5).HeaderCell.Value = "User Server"
            .Columns(6).HeaderCell.Value = "Barcode"
            .Columns(7).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(8).HeaderCell.Value = "Waktu Pengiriman"
            .Columns(9).HeaderCell.Value = "Tanggal Sampai"
            .Columns(10).HeaderCell.Value = "Waktu Sampai"
            .Columns(11).HeaderCell.Value = "Call Center"
            .Columns(12).HeaderCell.Value = "Liter"
            .Columns(13).HeaderCell.Value = "Tempat Loading"

        End With
    End Sub

    Private Sub tb_1_TextChanged(sender As Object, e As EventArgs) Handles tb_1.TextChanged

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlWorkSheet2 As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer
        Dim intData() As Integer = {1, 7, 11, 0, 12}


        DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.ActiveSheet, Excel.Worksheet)

        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To 2
                xlWorkSheet.Cells(1, 1) = DataGridView1.Columns(intData(0)).HeaderText

                xlWorkSheet.Cells(1, 2) = DataGridView1.Columns(intData(1)).HeaderText

                xlWorkSheet.Cells(1, 3) = DataGridView1.Columns(intData(2)).HeaderText

                xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(intData(j)).Value

            Next
        Next
        xlWorkSheet.Cells.EntireColumn.AutoFit()

        xlWorkSheet2 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To 4
                xlWorkSheet2.Cells(1, 1) = DataGridView1.Columns(intData(0)).HeaderText

                xlWorkSheet2.Cells(1, 2) = DataGridView1.Columns(intData(1)).HeaderText

                xlWorkSheet2.Cells(1, 3) = DataGridView1.Columns(intData(2)).HeaderText

                xlWorkSheet2.Cells(1, 4) = DataGridView1.Columns(intData(3)).HeaderText

                xlWorkSheet2.Cells(1, 5) = DataGridView1.Columns(intData(4)).HeaderText

                xlWorkSheet2.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(intData(j)).Value

            Next
        Next
        xlWorkSheet2.Cells.EntireColumn.AutoFit()

        Dim appPath As String = Application.StartupPath()
        xlWorkSheet.SaveAs(appPath + "\Excel\Level3.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("You can find the file " + appPath + "\Excel\Level3.xlsx")

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub tb_8_TextChanged(sender As Object, e As EventArgs) Handles tb_8.TextChanged
        If tb_8.Text = "" Then
            btn_4.Enabled = False
            btn_7.Enabled = False
        End If
        
    End Sub
End Class
