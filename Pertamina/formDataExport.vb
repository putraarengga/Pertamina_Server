Imports System.ComponentModel
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core

Public Class formDataExport

    Dim databaru As Boolean
    Dim selectDataBase, vJenisUser, tmpString As String
    Dim indexSatuan, indexKategori, indexLokasi As Integer
    Dim lNamaPerusahaan As New List(Of String)


    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub FormDataUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        databaru = False
        Dim appPath As String = Application.StartupPath()
        'Dim img As New System.Drawing.Icon(appPath + "\drugstore.ico")
        'Me.Icon = img
        PictureBox1.ImageLocation = appPath + ("\icons\export-icon.png")
    End Sub

    Sub IsiGridLv3()
        selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.kapasitasTruk,tkendaraan.namaPerusahaan,tkendaraan.callCenter " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = DS.Tables("tdistribusi")
        DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "Nomor DO"
            .Columns(1).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(2).HeaderCell.Value = "Liter"
            .Columns(3).HeaderCell.Value = "Transportir"
            .Columns(4).HeaderCell.Value = "Call Center"
        End With
        DataGridView1.Enabled = True
    End Sub
    Sub IsiGridLv2()
        selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.kapasitasTruk,tkendaraan.namaPerusahaan,tkendaraan.callCenter, " +
                        " tdistribusi.Keterangan ,tdatauser.NamaLengkap,tjenisuser.jenisUser " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                        " JOIN tdatauser ON tdatauser.IDUser = tdistribusi.IDUserClient " +
                        " JOIN tjenisuser ON tjenisuser.IDJenisUser = tdatauser.IDJenisUser "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = DS.Tables("tdistribusi")
        DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "No DO"
            .Columns(1).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(2).HeaderCell.Value = "Liter"
            .Columns(3).HeaderCell.Value = "Transportir"
            .Columns(4).HeaderCell.Value = "Call Center"
            .Columns(5).HeaderCell.Value = "Status"
            .Columns(6).HeaderCell.Value = "Nama Penerima"
            .Columns(7).HeaderCell.Value = "Unit PLTD"
        End With
        DataGridView1.Enabled = True
    End Sub
    Sub IsiGridLamp()
        selectDataBase = "SELECT tdistribusi.NoDO, tdistribusi.tglMuat," +
                                    " tkendaraan.namaPerusahaan, tkendaraan.kapasitasTruk " +
                                    " FROM tdistribusi " +
                                    " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                                    " WHERE tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED' "
        
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = DS.Tables("tdistribusi")
        DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "No DO"
            .Columns(1).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(2).HeaderCell.Value = "Transportir"
            .Columns(3).HeaderCell.Value = "Liter"
        End With
        DataGridView1.Enabled = True
    End Sub
    Sub Bersih()
        ComboBox1.Text = ""
        ComboBox1.Items.Clear()

    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        databaru = True
        ComboBox1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        DateTimePicker1.Enabled = True
        DateTimePicker5.Enabled = True
        If ComboBox1.SelectedItem = "Level 3" Then
            IsiGridLv3()
        ElseIf ComboBox1.SelectedItem = "Level 2" Then
            IsiGridLv2()
        ElseIf ComboBox1.SelectedItem = "Lamp. TUGS" Then
            IsiGridLamp()
        ElseIf ComboBox1.SelectedItem = "Lamp. TUG 3.4" Then
            IsiGridLamp34()

        End If

    End Sub

    Sub ShowJenisUser()
        selectDataBase = "SELECT JenisUser FROM tjenisuser "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)

        If DT.Rows.Count > 0 Then
            With ComboBox1
                .Items.Clear()
                For i As Integer = 0 To DT.Rows.Count - 1
                    .Items.Add(DT.Rows(i).Item("JenisUser"))
                Next
                .SelectedIndex = -1
            End With
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim tanggalAwal, tanggalAkhir As String
        bukaDB()
        DateTimePicker5.Format = DateTimePickerFormat.Custom
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
        tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-dd")
        tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-dd")
        If tanggalAwal = tanggalAkhir Then
            If ComboBox1.SelectedItem = "Level 3" Then
                selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.kapasitasTruk,tkendaraan.namaPerusahaan,tkendaraan.callCenter " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tdistribusi.tglMuat LIKE '%" & tanggalAkhir & "%'"
               
            ElseIf ComboBox1.SelectedItem = "Level 2" Then
                selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.kapasitasTruk,tkendaraan.namaPerusahaan,tkendaraan.callCenter, " +
                        " tdistribusi.Keterangan ,tdatauser.NamaLengkap,tjenisuser.jenisUser " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                        " JOIN tdatauser ON tdatauser.IDUser = tdistribusi.IDUserClient " +
                        " JOIN tjenisuser ON tjenisuser.IDJenisUser = tdatauser.IDJenisUser WHERE tdistribusi.tglMuat LIKE '%" & tanggalAkhir & "%'"

            ElseIf ComboBox1.SelectedItem = "Lamp. TUGS" Then
                selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.namaPerusahaan,tkendaraan.kapasitasTruk " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tdistribusi.tglMuat LIKE '%" & tanggalAkhir & "%' " +
                        " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "

            ElseIf ComboBox1.SelectedItem = "Lamp. TUG 3.4" Then
                selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.namaPerusahaan,tkendaraan.kapasitasTruk " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tdistribusi.tglMuat LIKE '%" & tanggalAkhir & "%' " +
                        " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "
            End If
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        Else
        If ComboBox1.SelectedItem = "Level 3" Then
            selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.kapasitasTruk,tkendaraan.namaPerusahaan,tkendaraan.callCenter " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'"
        ElseIf ComboBox1.SelectedItem = "Level 2" Then
            selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                    " tkendaraan.kapasitasTruk,tkendaraan.namaPerusahaan,tkendaraan.callCenter, " +
                    " tdistribusi.Keterangan ,tdatauser.NamaLengkap,tjenisuser.jenisUser " +
                    " FROM tdistribusi " +
                    " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                    " JOIN tdatauser ON tdatauser.IDUser = tdistribusi.IDUserClient " +
                    " JOIN tjenisuser ON tjenisuser.IDJenisUser = tdatauser.IDJenisUser WHERE tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'"
        ElseIf ComboBox1.SelectedItem = "Lamp. TUGS" Then
                selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.namaPerusahaan,tkendaraan.kapasitasTruk " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "' " +
                        " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "
            ElseIf ComboBox1.SelectedItem = "Lamp. TUG 3.4" Then
                selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.namaPerusahaan,tkendaraan.kapasitasTruk " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "' " +
                        " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "
        End If
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        End If

        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = (DS.Tables("tdistribusi"))
        DataGridView1.Enabled = True

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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox1.SelectedItem = "Level 3" Then
            ExportLv3()
        ElseIf ComboBox1.SelectedItem = "Level 2" Then
            ExportLv2()
        ElseIf ComboBox1.SelectedItem = "Lamp. TUGS" Then
            'isiGrid()
            'ExportLamp()
            'IsiGridLamp()
            isiGrid()
            GetNamaPerusahaan()
            ExportLamp2()
        ElseIf ComboBox1.SelectedItem = "Lamp. TUG 3.4" Then
            isiGrid()
            GetNamaPerusahaan()
            ExportLampTUG34()
        End If
    End Sub

    Private Sub FormDataUser_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()

        End If
    End Sub
    Sub ExportLv3()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer
        Dim intData() As Integer = {0, 1, 2, 3, 4}


        DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.ActiveSheet, Excel.Worksheet)

        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To 4
                xlWorkSheet.Cells(1, 1) = DataGridView1.Columns(intData(0)).HeaderText

                xlWorkSheet.Cells(1, 2) = DataGridView1.Columns(intData(1)).HeaderText

                xlWorkSheet.Cells(1, 3) = DataGridView1.Columns(intData(2)).HeaderText

                xlWorkSheet.Cells(1, 4) = DataGridView1.Columns(intData(3)).HeaderText

                xlWorkSheet.Cells(1, 5) = DataGridView1.Columns(intData(4)).HeaderText

                xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(intData(j)).Value

            Next
        Next
        xlWorkSheet.Cells.EntireColumn.AutoFit()

        Dim appPath As String = Application.StartupPath()
        xlWorkSheet.SaveAs(appPath + "\Excel\Level3.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("You can find the file " + appPath + "\Excel\Level3.xlsx")
    End Sub
    Sub ExportLv2()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlWorkSheet2 As Excel.Worksheet
        Dim xlWorkSheet3 As Excel.Worksheet
        Dim xlWorkSheet4 As Excel.Worksheet
        Dim xlWorkSheet5 As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        Dim excel2sheet1() As Integer = {1, 2, 5}
        Dim excel2sheet2() As Integer = {1, 3, 4, 5}
        Dim excel2sheet3() As Integer = {1, 2, 6, 7, 5}
        Dim excel2sheet4() As Integer = {0, 1, 5}
        Dim excel2sheet5() As Integer = {0, 2, 1, 5}

        DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.ActiveSheet, Excel.Worksheet)

        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To 2
                xlWorkSheet.Cells(1, 1) = DataGridView1.Columns(excel2sheet1(0)).HeaderText

                xlWorkSheet.Cells(1, 2) = DataGridView1.Columns(excel2sheet1(1)).HeaderText

                xlWorkSheet.Cells(1, 3) = DataGridView1.Columns(excel2sheet1(2)).HeaderText

                xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(excel2sheet1(j)).Value

            Next
        Next
        xlWorkSheet.Cells.EntireColumn.AutoFit()
        xlWorkSheet2 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )

        xlWorkSheet2 = xlWorkBook.Sheets("sheet2")
        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To 3
                xlWorkSheet2.Cells(1, 1) = DataGridView1.Columns(excel2sheet2(0)).HeaderText

                xlWorkSheet2.Cells(1, 2) = DataGridView1.Columns(excel2sheet2(1)).HeaderText

                xlWorkSheet2.Cells(1, 3) = DataGridView1.Columns(excel2sheet2(2)).HeaderText

                xlWorkSheet2.Cells(1, 4) = DataGridView1.Columns(excel2sheet2(3)).HeaderText

                xlWorkSheet2.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(excel2sheet2(j)).Value

            Next
        Next
        xlWorkSheet2.Cells.EntireColumn.AutoFit()

        xlWorkSheet3 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        xlWorkSheet3 = xlWorkBook.Sheets("sheet3")
        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To 4
                xlWorkSheet3.Cells(1, 1) = DataGridView1.Columns(excel2sheet3(0)).HeaderText

                xlWorkSheet3.Cells(1, 2) = DataGridView1.Columns(excel2sheet3(1)).HeaderText

                xlWorkSheet3.Cells(1, 3) = DataGridView1.Columns(excel2sheet3(2)).HeaderText

                xlWorkSheet3.Cells(1, 4) = DataGridView1.Columns(excel2sheet3(3)).HeaderText

                xlWorkSheet3.Cells(1, 5) = DataGridView1.Columns(excel2sheet3(4)).HeaderText

                xlWorkSheet3.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(excel2sheet3(j)).Value

            Next
        Next
        xlWorkSheet3.Cells.EntireColumn.AutoFit()

        xlWorkSheet4 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        xlWorkSheet4 = xlWorkBook.Sheets("sheet4")
        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To 2
                xlWorkSheet4.Cells(1, 1) = DataGridView1.Columns(excel2sheet4(0)).HeaderText

                xlWorkSheet4.Cells(1, 2) = DataGridView1.Columns(excel2sheet4(1)).HeaderText

                xlWorkSheet4.Cells(1, 3) = DataGridView1.Columns(excel2sheet4(2)).HeaderText

                xlWorkSheet4.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(excel2sheet4(j)).Value

            Next
        Next
        xlWorkSheet4.Cells.EntireColumn.AutoFit()

        xlWorkSheet5 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        xlWorkSheet5 = xlWorkBook.Sheets("sheet5")
        For i = 0 To DataGridView1.RowCount - 1
            For j = 0 To 3
                xlWorkSheet5.Cells(1, 1) = DataGridView1.Columns(excel2sheet5(0)).HeaderText

                xlWorkSheet5.Cells(1, 2) = DataGridView1.Columns(excel2sheet5(1)).HeaderText

                xlWorkSheet5.Cells(1, 3) = DataGridView1.Columns(excel2sheet5(2)).HeaderText

                xlWorkSheet5.Cells(1, 4) = DataGridView1.Columns(excel2sheet5(3)).HeaderText

                xlWorkSheet5.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(excel2sheet5(j)).Value

            Next
        Next
        xlWorkSheet5.Cells.EntireColumn.AutoFit()
        Dim appPath As String = Application.StartupPath()
        xlWorkSheet.SaveAs(appPath + "\Excel\Level2.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("You can find the file " + appPath + "\Excel\Level2.xlsx")
    End Sub
    Sub isiGrid()
        Dim tanggalAwal, tanggalAkhir As String
        Dim duplicateDictionary As New Dictionary(Of Integer, Integer) 'value, count
        lNamaPerusahaan.Clear()

        bukaDB()
        DateTimePicker5.Format = DateTimePickerFormat.Custom
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
        tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-mm-dd")
        tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-mm-dd")

        selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                        " tkendaraan.namaPerusahaan,tkendaraan.kapasitasTruk " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'"
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = (DS.Tables("tdistribusi"))
        DataGridView1.Enabled = True

        DataGridView1.Sort(DataGridView1.Columns(2), ListSortDirection.Ascending)


    End Sub
    Sub GetNamaPerusahaan()
        Dim tanggalAwal, tanggalAkhir As String
        Dim duplicateDictionary As New Dictionary(Of Integer, Integer) 'value, count
        lNamaPerusahaan.Clear()

        bukaDB()
        DateTimePicker5.Format = DateTimePickerFormat.Custom
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
        tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-dd")
        tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-dd")

        selectDataBase = "SELECT DISTINCT tkendaraan.namaPerusahaan " +
                        " FROM tdistribusi " +
                        " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'"
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DA.Fill(DT)
        Dim i As Integer
        For i = 0 To DT.Rows.Count - 1
            lNamaPerusahaan.Add(DT.Rows(i).Item("namaPerusahaan"))
        Next

        Label3.Text = lNamaPerusahaan.Count
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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        isiGrid()
        GetNamaPerusahaan()

        ExportLamp2()

    End Sub

    Sub ExportLamp2()
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlWorkSheet2 As Excel.Worksheet
        Dim xlWorkSheet3 As Excel.Worksheet
        Dim xlWorkSheet4 As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer
        Dim excel3sheet1() As Integer = {0, 1, 2, 3, 4, 5}
        Dim tanggalAwal, tanggalAkhir As String
        Dim total, xx As Integer
        Dim x As Integer = 0
        Dim rowData As Integer = 0
        Dim rowAwal As Integer = 0
        Dim rowAkhir As Integer = 0
        Dim range As Excel.Range
        Dim r As Excel.Range
        Dim fileName As String
        Dim pictureWidth As Integer
        Dim pictureHeight As Integer
        Dim shape As Excel.Shape
        Dim appPath As String = Application.StartupPath()
        fileName = appPath + "\logoPLNcolor.png"


        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.ActiveSheet, Excel.Worksheet)
        total = 0
        For xx = 0 To lNamaPerusahaan.Count - 1
            bukaDB()
            DateTimePicker5.Format = DateTimePickerFormat.Custom
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
            tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-1")
            tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-7")

            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO, " +
                                    " tdistribusi.tglMuat,tkendaraan.kapasitasTruk " +
                                    " FROM tdistribusi " +
                                    " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tkendaraan.namaPerusahaan = '" & lNamaPerusahaan.Item(xx).ToString & "' " +
                                    " AND tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "' " +
                                    " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "
            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tdistribusi")
            DataGridView1.DataSource = (DS.Tables("tdistribusi"))
            DataGridView1.Enabled = True
            DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)

            If DataGridView1.RowCount <> 0 Then


                r = xlWorkSheet.Cells(rowData + 3, 2)


                shape = xlWorkSheet.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      r.Left, r.Top, pictureWidth, pictureHeight)

                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                shape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)

                rowData = rowData + 1
                xlWorkSheet.Cells(rowData, 2) = "PT. PLN ( Persero )"
                xlWorkSheet.Range(xlWorkSheet.Cells(rowData, 3), xlWorkSheet.Cells(rowData, 6)).Merge()
                xlWorkSheet.Cells(rowData, 3) = "LAMPIRAN TUG. 3 No. 340/LOG.00.01/BBM.SOLAR PL. WNA/SP2B/2016"

                rowData = rowData + 1
                xlWorkSheet.Cells(rowData, 2) = "Sektor Papua & Papua Barat"
                xlWorkSheet.Cells(rowData, 3) = "TANGGAL. 01  s/d  7 - " + DateTimePicker5.Value.Date.Month.ToString + " -  2016"
                rowData = rowData + 1
                xlWorkSheet.Range(xlWorkSheet.Cells(rowData, 3), xlWorkSheet.Cells(rowData, 4)).Merge()
                xlWorkSheet.Cells(rowData, 3) = "REKAPITULASI PENERIMAAN BBM/SOLAR PLTD WAENA"
                rowData = rowData + 1
                xlWorkSheet.Range(xlWorkSheet.Cells(rowData, 3), xlWorkSheet.Cells(rowData, 4)).Merge()
                xlWorkSheet.Cells(rowData, 3) = "Periode        : I Tgl.  1 s/d  7 - " + DateTimePicker5.Value.Date.Month.ToString + " -  2016"
                rowData = rowData + 1
                xlWorkSheet.Range(xlWorkSheet.Cells(rowData, 3), xlWorkSheet.Cells(rowData, 4)).Merge()
                xlWorkSheet.Cells(rowData, 3) = "Lampiran     : " + DataGridView1.Rows(0).Cells(0).Value.ToString

                rowData = rowData + 2
                rowAwal = rowData

                total = 0
                For i = 0 To DataGridView1.RowCount - 1
                    For j = 0 To 5
                        x = i + 1
                        xlWorkSheet.Cells(rowData, 1) = "No"

                        xlWorkSheet.Cells(rowData, 2) = DataGridView1.Columns(excel3sheet1(0)).HeaderText

                        xlWorkSheet.Cells(rowData, 3) = DataGridView1.Columns(excel3sheet1(1)).HeaderText

                        xlWorkSheet.Cells(rowData, 4) = DataGridView1.Columns(excel3sheet1(2)).HeaderText

                        xlWorkSheet.Cells(rowData, 5) = DataGridView1.Columns(excel3sheet1(3)).HeaderText

                        xlWorkSheet.Cells(rowData, 6) = "Keterangan"

                        If j = 0 Then
                            xlWorkSheet.Cells(i + rowData + 1, j + 1) = x.ToString

                        ElseIf j = 5 Then

                        Else
                            xlWorkSheet.Cells(i + rowData + 1, j + 1) = DataGridView1.Rows(i).Cells(excel3sheet1(j - 1)).Value
                        End If


                    Next
                    total += DataGridView1.Rows(i).Cells(3).Value
                Next
                If i <= 40 Then
                    rowData = rowData + 40
                Else
                    rowData = rowData + i
                End If
                rowAkhir = rowData
                xlWorkSheet.Range(xlWorkSheet.Cells(rowData, 1), xlWorkSheet.Cells(rowData, 4)).Merge()
                xlWorkSheet.Cells(rowData, 1) = "JUMLAH"
                xlWorkSheet.Range("A" + rowData.ToString).VerticalAlignment = Excel.Constants.xlCenter
                xlWorkSheet.Cells(rowData, 5) = total.ToString
                rowData = rowData + 2
                xlWorkSheet.Cells(rowData, 4) = "Yang membuat,	"
                rowData = rowData + 1
                xlWorkSheet.Cells(rowData, 4) = "Jr. OFFICER LOGISTIK	"
                rowData = rowData + 4
                xlWorkSheet.Cells(rowData, 4) = "Y U S U F	"
                xlWorkSheet.Cells.EntireColumn.AutoFit()
                range = xlWorkSheet.Range("A" + rowAwal.ToString, "F" + rowAkhir.ToString)

                Dim borders As Excel.Borders = range.Borders
                'Set the thi lines style.
                borders.LineStyle = Excel.XlLineStyle.xlContinuous
                borders.Weight = 2.0R

                rowData += 3
            End If

        Next
        '===================================================================================================================
        xlWorkSheet2 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        rowData = 0
        total = 0
        For xx = 0 To lNamaPerusahaan.Count - 1
            bukaDB()
            DateTimePicker5.Format = DateTimePickerFormat.Custom
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
            tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-8")
            tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-14")

            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO, " +
                                    " tdistribusi.tglMuat,tkendaraan.kapasitasTruk " +
                                    " FROM tdistribusi " +
                                    " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tkendaraan.namaPerusahaan = '" & lNamaPerusahaan.Item(xx).ToString & "' AND tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'"
            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tdistribusi")
            DataGridView1.DataSource = (DS.Tables("tdistribusi"))
            DataGridView1.Enabled = True
            DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)

            If DataGridView1.RowCount <> 0 Then
                r = xlWorkSheet2.Cells(rowData + 3, 2)


                shape = xlWorkSheet2.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      r.Left, r.Top, pictureWidth, pictureHeight)

                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                shape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)

                rowData = rowData + 1
                xlWorkSheet2.Cells(rowData, 2) = "PT. PLN ( Persero )"
                xlWorkSheet2.Range(xlWorkSheet2.Cells(rowData, 3), xlWorkSheet2.Cells(rowData, 6)).Merge()
                xlWorkSheet2.Cells(rowData, 3) = "LAMPIRAN TUG. 3 No. 340/LOG.00.01/BBM.SOLAR PL. WNA/SP2B/2016"
                rowData = rowData + 1
                xlWorkSheet2.Cells(rowData, 2) = "Sektor Papua & Papua Barat"
                xlWorkSheet2.Cells(rowData, 3) = "TANGGAL. 08  s/d  14  - " + DateTimePicker5.Value.Date.Month.ToString + " -  2016"
                rowData = rowData + 1
                xlWorkSheet2.Range(xlWorkSheet2.Cells(rowData, 3), xlWorkSheet2.Cells(rowData, 4)).Merge()
                xlWorkSheet2.Cells(rowData, 3) = "REKAPITULASI PENERIMAAN BBM/SOLAR PLTD WAENA"
                rowData = rowData + 1
                xlWorkSheet2.Range(xlWorkSheet2.Cells(rowData, 3), xlWorkSheet2.Cells(rowData, 4)).Merge()
                xlWorkSheet2.Cells(rowData, 3) = "Periode        : II Tgl.  08 s/d  14  - " + DateTimePicker5.Value.Date.Month.ToString + " -  2016"
                rowData = rowData + 1
                xlWorkSheet2.Range(xlWorkSheet2.Cells(rowData, 3), xlWorkSheet2.Cells(rowData, 4)).Merge()
                xlWorkSheet2.Cells(rowData, 3) = "Lampiran     : " + DataGridView1.Rows(0).Cells(0).Value.ToString

                rowData = rowData + 2
                rowAwal = rowData

                total = 0
                For i = 0 To DataGridView1.RowCount - 1
                    For j = 0 To 5
                        x = i + 1
                        xlWorkSheet2.Cells(rowData, 1) = "No"

                        xlWorkSheet2.Cells(rowData, 2) = DataGridView1.Columns(excel3sheet1(0)).HeaderText

                        xlWorkSheet2.Cells(rowData, 3) = DataGridView1.Columns(excel3sheet1(1)).HeaderText

                        xlWorkSheet2.Cells(rowData, 4) = DataGridView1.Columns(excel3sheet1(2)).HeaderText

                        xlWorkSheet2.Cells(rowData, 5) = DataGridView1.Columns(excel3sheet1(3)).HeaderText

                        xlWorkSheet2.Cells(rowData, 6) = "Keterangan"

                        If j = 0 Then
                            xlWorkSheet2.Cells(i + rowData + 1, j + 1) = x.ToString

                        ElseIf j = 5 Then

                        Else
                            xlWorkSheet2.Cells(i + rowData + 1, j + 1) = DataGridView1.Rows(i).Cells(excel3sheet1(j - 1)).Value
                        End If


                    Next
                    total += DataGridView1.Rows(i).Cells(3).Value
                Next
                If i <= 40 Then
                    rowData = rowData + 40
                Else
                    rowData = rowData + i
                End If
                rowAkhir = rowData
                xlWorkSheet2.Range(xlWorkSheet2.Cells(rowData, 1), xlWorkSheet2.Cells(rowData, 4)).Merge()
                xlWorkSheet2.Cells(rowData, 1) = "JUMLAH"
                xlWorkSheet2.Range("A" + rowData.ToString).VerticalAlignment = Excel.Constants.xlCenter
                xlWorkSheet2.Cells(rowData, 5) = total.ToString
                rowData = rowData + 2
                xlWorkSheet2.Cells(rowData, 4) = "Yang membuat,	"
                rowData = rowData + 1
                xlWorkSheet2.Cells(rowData, 4) = "Jr. OFFICER LOGISTIK	"
                rowData = rowData + 4
                xlWorkSheet2.Cells(rowData, 4) = "Y U S U F	"
                xlWorkSheet2.Cells.EntireColumn.AutoFit()
                range = xlWorkSheet2.Range("A" + rowAwal.ToString, "F" + rowAkhir.ToString)

                Dim borders As Excel.Borders = range.Borders
                'Set the thi lines style.
                borders.LineStyle = Excel.XlLineStyle.xlContinuous
                borders.Weight = 2.0R

                rowData += 3
            End If

        Next

        '===================================================================================================================
        xlWorkSheet3 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        rowData = 0
        total = 0
        For xx = 0 To lNamaPerusahaan.Count - 1
            bukaDB()
            DateTimePicker5.Format = DateTimePickerFormat.Custom
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
            tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-15")
            tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-22")

            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO, " +
                                    " tdistribusi.tglMuat,tkendaraan.kapasitasTruk " +
                                    " FROM tdistribusi " +
                                    " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tkendaraan.namaPerusahaan = '" & lNamaPerusahaan.Item(xx).ToString & "' AND tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "'"
            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tdistribusi")
            DataGridView1.DataSource = (DS.Tables("tdistribusi"))
            DataGridView1.Enabled = True
            DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)

            If DataGridView1.RowCount <> 0 Then
                r = xlWorkSheet3.Cells(rowData + 3, 2)


                shape = xlWorkSheet3.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      r.Left, r.Top, pictureWidth, pictureHeight)

                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                shape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)

                rowData = rowData + 1
                xlWorkSheet3.Cells(rowData, 2) = "PT. PLN ( Persero )"
                xlWorkSheet3.Range(xlWorkSheet3.Cells(rowData, 3), xlWorkSheet3.Cells(rowData, 6)).Merge()
                xlWorkSheet3.Cells(rowData, 3) = "LAMPIRAN TUG. 3 No. 340/LOG.00.01/BBM.SOLAR PL. WNA/SP2B/2016"
                rowData = rowData + 1
                xlWorkSheet3.Cells(rowData, 2) = "Sektor Papua & Papua Barat"
                xlWorkSheet3.Cells(rowData, 3) = "TANGGAL. 15  s/d  22  - " + DateTimePicker5.Value.Date.Month.ToString + " -  2016"
                rowData = rowData + 1
                xlWorkSheet3.Range(xlWorkSheet3.Cells(rowData, 3), xlWorkSheet3.Cells(rowData, 4)).Merge()
                xlWorkSheet3.Cells(rowData, 3) = "REKAPITULASI PENERIMAAN BBM/SOLAR PLTD WAENA"
                rowData = rowData + 1
                xlWorkSheet3.Range(xlWorkSheet3.Cells(rowData, 3), xlWorkSheet3.Cells(rowData, 4)).Merge()
                xlWorkSheet3.Cells(rowData, 3) = "Periode        : III Tgl.  15 s/d  22  - " + DateTimePicker5.Value.Date.Month.ToString + " -  2016"
                rowData = rowData + 1
                xlWorkSheet3.Range(xlWorkSheet3.Cells(rowData, 3), xlWorkSheet3.Cells(rowData, 4)).Merge()
                xlWorkSheet3.Cells(rowData, 3) = "Lampiran     : " + DataGridView1.Rows(0).Cells(0).Value.ToString

                rowData = rowData + 2
                rowAwal = rowData

                total = 0
                For i = 0 To DataGridView1.RowCount - 1
                    For j = 0 To 5
                        x = i + 1
                        xlWorkSheet3.Cells(rowData, 1) = "No"

                        xlWorkSheet3.Cells(rowData, 2) = DataGridView1.Columns(excel3sheet1(0)).HeaderText

                        xlWorkSheet3.Cells(rowData, 3) = DataGridView1.Columns(excel3sheet1(1)).HeaderText

                        xlWorkSheet3.Cells(rowData, 4) = DataGridView1.Columns(excel3sheet1(2)).HeaderText

                        xlWorkSheet3.Cells(rowData, 5) = DataGridView1.Columns(excel3sheet1(3)).HeaderText

                        xlWorkSheet3.Cells(rowData, 6) = "Keterangan"

                        If j = 0 Then
                            xlWorkSheet3.Cells(i + rowData + 1, j + 1) = x.ToString

                        ElseIf j = 5 Then

                        Else
                            xlWorkSheet3.Cells(i + rowData + 1, j + 1) = DataGridView1.Rows(i).Cells(excel3sheet1(j - 1)).Value
                        End If


                    Next
                    total += DataGridView1.Rows(i).Cells(3).Value
                Next
                If i <= 40 Then
                    rowData = rowData + 40
                Else
                    rowData = rowData + i
                End If
                rowAkhir = rowData
                xlWorkSheet3.Range(xlWorkSheet3.Cells(rowData, 1), xlWorkSheet3.Cells(rowData, 4)).Merge()
                xlWorkSheet3.Cells(rowData, 1) = "JUMLAH"
                xlWorkSheet3.Range("A" + rowData.ToString).VerticalAlignment = Excel.Constants.xlCenter
                xlWorkSheet3.Cells(rowData, 5) = total.ToString
                rowData = rowData + 2
                xlWorkSheet3.Cells(rowData, 4) = "Yang membuat,	"
                rowData = rowData + 1
                xlWorkSheet3.Cells(rowData, 4) = "Jr. OFFICER LOGISTIK	"
                rowData = rowData + 4
                xlWorkSheet3.Cells(rowData, 4) = "Y U S U F	"
                xlWorkSheet3.Cells.EntireColumn.AutoFit()
                range = xlWorkSheet3.Range("A" + rowAwal.ToString, "F" + rowAkhir.ToString)

                Dim borders As Excel.Borders = range.Borders
                'Set the thi lines style.
                borders.LineStyle = Excel.XlLineStyle.xlContinuous
                borders.Weight = 2.0R

                rowData += 3

            End If

           
        Next
        '===================================================================================================================
        xlWorkSheet4 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        rowData = 0
        total = 0
        For xx = 0 To lNamaPerusahaan.Count - 1
            bukaDB()
            DateTimePicker5.Format = DateTimePickerFormat.Custom
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
            tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-22")
            tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-31")

            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO, " +
                                    " tdistribusi.tglMuat,tkendaraan.kapasitasTruk " +
                                    " FROM tdistribusi " +
                                    " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tkendaraan.namaPerusahaan = '" & lNamaPerusahaan.Item(xx).ToString & "' " +
                                    " AND tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "' " +
                                    " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "

            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tdistribusi")
            DataGridView1.DataSource = (DS.Tables("tdistribusi"))
            DataGridView1.Enabled = True
            DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)

            If DataGridView1.RowCount <> 0 Then
                r = xlWorkSheet4.Cells(rowData + 3, 2)


                shape = xlWorkSheet4.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      r.Left, r.Top, pictureWidth, pictureHeight)

                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                shape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)

                rowData = rowData + 1
                xlWorkSheet4.Cells(rowData, 2) = "PT. PLN ( Persero )"
                xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 3), xlWorkSheet4.Cells(rowData, 6)).Merge()
                xlWorkSheet4.Cells(rowData, 3) = "LAMPIRAN TUG. 3 No. 340/LOG.00.01/BBM.SOLAR PL. WNA/SP2B/2016"
                rowData = rowData + 1
                xlWorkSheet4.Cells(rowData, 2) = "Sektor Papua & Papua Barat"
                xlWorkSheet4.Cells(rowData, 3) = "TANGGAL. 23  s/d  31 - " + DateTimePicker5.Value.Date.Month.ToString + " -  2016"
                rowData = rowData + 1
                xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 3), xlWorkSheet4.Cells(rowData, 4)).Merge()
                xlWorkSheet4.Cells(rowData, 3) = "REKAPITULASI PENERIMAAN BBM/SOLAR PLTD WAENA"
                rowData = rowData + 1
                xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 3), xlWorkSheet4.Cells(rowData, 4)).Merge()
                xlWorkSheet4.Cells(rowData, 3) = "Periode        : IV Tgl.  23 s/d  31  - " + DateTimePicker5.Value.Date.Month.ToString + " -  2016"
                rowData = rowData + 1
                xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 3), xlWorkSheet4.Cells(rowData, 4)).Merge()
                xlWorkSheet4.Cells(rowData, 3) = "Lampiran     : " + DataGridView1.Rows(0).Cells(0).Value.ToString

                rowData = rowData + 2
                rowAwal = rowData

                total = 0
                For i = 0 To DataGridView1.RowCount - 1
                    For j = 0 To 5
                        x = i + 1
                        xlWorkSheet4.Cells(rowData, 1) = "No"

                        xlWorkSheet4.Cells(rowData, 2) = DataGridView1.Columns(excel3sheet1(0)).HeaderText

                        xlWorkSheet4.Cells(rowData, 3) = DataGridView1.Columns(excel3sheet1(1)).HeaderText

                        xlWorkSheet4.Cells(rowData, 4) = DataGridView1.Columns(excel3sheet1(2)).HeaderText

                        xlWorkSheet4.Cells(rowData, 5) = DataGridView1.Columns(excel3sheet1(3)).HeaderText

                        xlWorkSheet4.Cells(rowData, 6) = "Keterangan"

                        If j = 0 Then
                            xlWorkSheet4.Cells(i + rowData + 1, j + 1) = x.ToString

                        ElseIf j = 5 Then

                        Else
                            xlWorkSheet4.Cells(i + rowData + 1, j + 1) = DataGridView1.Rows(i).Cells(excel3sheet1(j - 1)).Value
                        End If


                    Next
                    total += DataGridView1.Rows(i).Cells(3).Value
                Next
                If i <= 40 Then
                    rowData = rowData + 40
                Else
                    rowData = rowData + i
                End If
                rowAkhir = rowData
                xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData, 4)).Merge()
                xlWorkSheet4.Cells(rowData, 1) = "JUMLAH"
                xlWorkSheet4.Range("A" + rowData.ToString).VerticalAlignment = Excel.Constants.xlCenter
                xlWorkSheet4.Cells(rowData, 5) = total.ToString
                rowData = rowData + 2
                xlWorkSheet4.Cells(rowData, 4) = "Yang membuat,	"
                rowData = rowData + 1
                xlWorkSheet4.Cells(rowData, 4) = "Jr. OFFICER LOGISTIK	"
                rowData = rowData + 4
                xlWorkSheet4.Cells(rowData, 4) = "Y U S U F	"
                xlWorkSheet4.Cells.EntireColumn.AutoFit()
                range = xlWorkSheet4.Range("A" + rowAwal.ToString, "F" + rowAkhir.ToString)

                Dim borders As Excel.Borders = range.Borders
                'Set the thi lines style.
                borders.LineStyle = Excel.XlLineStyle.xlContinuous
                borders.Weight = 2.0R

                rowData += 3
            End If

        Next

        tanggalAkhir = Format(DateTimePicker1.Value.Date, "MMMM")
        xlWorkSheet.SaveAs(appPath + "\Excel\Lamp_TUGS_" + tanggalAkhir + ".xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("You can find the file " + appPath + "\Excel\Lamp_TUGS" + tanggalAkhir + ".xlsx")
    End Sub

    Sub IsiGridLamp34()
        selectDataBase = "SELECT tdistribusi.NoDO,tdistribusi.tglMuat, " +
                       " tkendaraan.namaPerusahaan,tkendaraan.kapasitasTruk " +
                       " FROM tdistribusi " +
                       " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan " +
                       " WHERE tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        DataGridView1.DataSource = DS.Tables("tdistribusi")
        DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)
        With DataGridView1
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "No DO"
            .Columns(1).HeaderCell.Value = "Tanggal Pengiriman"
            .Columns(2).HeaderCell.Value = "Transportir"
            .Columns(3).HeaderCell.Value = "Liter"
        End With
        DataGridView1.Enabled = True
    End Sub
    Sub ExportLampTUG34()

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet, xlWorkSheet2, xlWorkSheet3, xlWorkSheet4 As Excel.Worksheet
        Dim xlWorkSheet5, xlWorkSheet6, xlWorkSheet7, xlWorkSheet8 As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer
        Dim excel4sheet1() As Integer = {0, 1, 2, 3, 4, 5}
        Dim tanggalAwal, tanggalAkhir As String
        Dim total, xx As Integer
        Dim x As Integer = 0
        Dim rowData As Integer = 0
        Dim rowAwal As Integer = 0
        Dim rowAkhir As Integer = 0
        Dim range As Excel.Range
        Dim r As Excel.Range
        Dim fileName As String
        Dim pictureWidth As Integer
        Dim pictureHeight As Integer
        Dim shape As Excel.Shape
        Dim appPath As String = Application.StartupPath()
        fileName = appPath + "\logoPLNcolor.png"


        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.ActiveSheet, Excel.Worksheet)

        exportLampTUGS3Sheet(xlWorkSheet, rowData, 1, 7)
        xlWorkSheet2 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        exportLampTUGS3Sheet(xlWorkSheet2, rowData, 8, 14)
        xlWorkSheet3 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        exportLampTUGS3Sheet(xlWorkSheet3, rowData, 15, 21)
        xlWorkSheet4 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        exportLampTUGS3Sheet(xlWorkSheet4, rowData, 22, 31)


        xlWorkSheet5 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        exportLampTUGS4Sheet(xlWorkSheet5, rowData, 1, 7)
        xlWorkSheet6 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        exportLampTUGS4Sheet(xlWorkSheet6, rowData, 8, 14)
        xlWorkSheet7 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        exportLampTUGS4Sheet(xlWorkSheet7, rowData, 15, 21)
        xlWorkSheet8 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )
        exportLampTUGS4Sheet(xlWorkSheet8, rowData, 22, 31)


        'xlWorkSheet4.Range("A1").ColumnWidth = 9.29
        'xlWorkSheet4.Range("B1").ColumnWidth = 1
        'xlWorkSheet4.Range("C1").ColumnWidth = 33.43
        'xlWorkSheet4.Range("D1").ColumnWidth = 1
        'xlWorkSheet4.Range("E1").ColumnWidth = 3
        'xlWorkSheet4.Range("F1").ColumnWidth = 9.43
        'xlWorkSheet4.Range("G1").ColumnWidth = 1.57
        'xlWorkSheet4.Range("H1").ColumnWidth = 7
        'xlWorkSheet4.Range("I1").ColumnWidth = 1.29
        'xlWorkSheet4.Range("J1").ColumnWidth = 8.86
        'xlWorkSheet4.Range("K1").ColumnWidth = 0.75
        'xlWorkSheet4.Range("L1").ColumnWidth = 0.92
        'xlWorkSheet4.Range("M1").ColumnWidth = 4
        'xlWorkSheet4.Range("N1").ColumnWidth = 8.43
        'xlWorkSheet4.Range("O1").ColumnWidth = 9
        'xlWorkSheet4.Range("P1").ColumnWidth = 1.29
        'xlWorkSheet4.Range("Q1").ColumnWidth = 3.86
        'xlWorkSheet4.Range("R1").ColumnWidth = 6.57
        'xlWorkSheet4.Range("S1").ColumnWidth = 11.29
        'xlWorkSheet4.Range("T1").ColumnWidth = 8.57
        'xlWorkSheet4.Range("U1").ColumnWidth = 4.71
        'xlWorkSheet4.Range("V1").ColumnWidth = 2.57
        'xlWorkSheet4.Range("W1").ColumnWidth = 4.86
        'xlWorkSheet4.Range("X1").ColumnWidth = 11.43
        'xlWorkSheet4.Range("Y1").ColumnWidth = 1.43
        'xlWorkSheet4.Range("Z1").ColumnWidth = 8.43
        'xlWorkSheet4.Range("AA1").ColumnWidth = 8.43


        'rowData = 0
        'total = 0
        'For xx = 0 To lNamaPerusahaan.Count - 1
        '    bukaDB()
        '    DateTimePicker5.Format = DateTimePickerFormat.Custom
        '    DateTimePicker1.Format = DateTimePickerFormat.Custom
        '    'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
        '    tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-22")
        '    tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-31")

        '    selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO, " +
        '                            " tdistribusi.tglMuat,tkendaraan.kapasitasTruk " +
        '                            " FROM tdistribusi " +
        '                            " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tkendaraan.namaPerusahaan = '" & lNamaPerusahaan.Item(xx).ToString & "' " +
        '                            " AND tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "' " +
        '                            " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "

        '    DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        '    DS = New DataSet
        '    DS.Clear()
        '    DA.Fill(DS, "tdistribusi")
        '    DataGridView1.DataSource = (DS.Tables("tdistribusi"))
        '    DataGridView1.Enabled = True
        '    DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)

        '    If DataGridView1.RowCount <> 0 Then
        '        r = xlWorkSheet4.Cells(rowData + 2, 1)


        '        shape = xlWorkSheet4.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue,
        '                                              Microsoft.Office.Core.MsoTriState.msoTrue,
        '                                              r.Left, r.Top, pictureWidth, pictureHeight)

        '        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
        '        shape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 4.5

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 18
        '        xlWorkSheet4.Cells(rowData, 2) = "PT. PLN (PERSERO) WILAYAH PAPUA"
        '        xlWorkSheet4.Cells(rowData, 24) = "TUG. 3"
        '        xlWorkSheet4.Cells(rowData, 26) = "MERAH"
        '        xlWorkSheet4.Cells(rowData, 27) = "1. Untuk TU Keuangan Persediaan"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Cells(rowData, 2) = "SEKTOR PAPUA & PAPUA BARAT"
        '        xlWorkSheet4.Cells(rowData, 26) = "BIRU"
        '        xlWorkSheet4.Cells(rowData, 27) = "2. Untuk T.U.G"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Cells(rowData, 2) = "BAGIAN LOGISTIK"
        '        xlWorkSheet4.Cells(rowData, 26) = "HIJAU"
        '        xlWorkSheet4.Cells(rowData, 27) = "3. Untuk Perbekalan"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Cells(rowData, 23) = "No."
        '        xlWorkSheet4.Cells(rowData, 24) = "1339"
        '        xlWorkSheet4.Cells(rowData, 26) = "KUNING"
        '        xlWorkSheet4.Cells(rowData, 27) = "4. Untuk Pengiriman Barang"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData, 26)).Merge()
        '        xlWorkSheet4.Cells(rowData, 1) = "BON PENERIMAAN BARANG-BARANG GUDANG"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 21
        '        xlWorkSheet4.Cells(rowData, 6) = "  Nomor :"
        '        xlWorkSheet4.Cells(rowData, 8) = "339"
        '        xlWorkSheet4.Cells(rowData, 9) = "/"
        '        xlWorkSheet4.Cells(rowData, 10) = "LOG.00.01"
        '        xlWorkSheet4.Cells(rowData, 11) = "/"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 12), xlWorkSheet4.Cells(rowData, 18)).Merge()
        '        xlWorkSheet4.Cells(rowData, 12) = "BBM. SOLAR  PL. WNA / SP2B / 2016"
        '        xlWorkSheet4.Cells(rowData, 20) = "1. Untuk Pengiriman Barang"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 12.75
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 3), xlWorkSheet4.Cells(rowData, 4)).Merge()
        '        xlWorkSheet4.Cells(rowData, 20) = "P L N"
        '        xlWorkSheet4.Cells(rowData, 21) = ":"
        '        xlWorkSheet4.Cells(rowData, 22) = "WILAYAH PAPUA"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 12.75
        '        xlWorkSheet4.Cells(rowData, 20) = "SEKTOR"
        '        xlWorkSheet4.Cells(rowData, 21) = ":"
        '        xlWorkSheet4.Cells(rowData, 22) = "PAPUA & PAPUA BARAT"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 6.75

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 3

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 8.25

        '        rowData = rowData + 1
        '        xlWorkSheet4.Cells(rowData, 1) = "Diterima Tgl."
        '        xlWorkSheet4.Cells(rowData, 2) = " :"
        '        DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
        '        tanggalAkhir = Format(DateTimePicker1.Value.Date, "dd MMMM yyyy")
        '        xlWorkSheet4.Cells(rowData, 3) = "'01 s.d " + tanggalAkhir
        '        xlWorkSheet4.Cells(rowData, 10) = "Pembelian ditempat, lihat Faktur/Bukti Kas No."
        '        xlWorkSheet4.Cells(rowData, 16) = " :"
        '        xlWorkSheet4.Cells(rowData, 18) = " …………………………………………………………………………………….."

        '        rowData = rowData + 1
        '        xlWorkSheet4.Cells(rowData, 1) = "Dari"
        '        xlWorkSheet4.Cells(rowData, 2) = " :"
        '        xlWorkSheet4.Cells(rowData, 3) = "DEPOT PERTAMINA ( PT.WIRA SEMBADA PERKASA )"
        '        xlWorkSheet4.Cells(rowData, 10) = "Diterima Bon Pengiriman No."
        '        xlWorkSheet4.Cells(rowData, 16) = " :"
        '        xlWorkSheet4.Cells(rowData, 18) = "  ……………………………………………."
        '        xlWorkSheet4.Cells(rowData, 21) = "Tgl.  :"
        '        xlWorkSheet4.Cells(rowData, 22) = " …………………………………"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Cells(rowData, 1) = "Dengan"
        '        xlWorkSheet4.Cells(rowData, 2) = " :"
        '        xlWorkSheet4.Cells(rowData, 10) = "Menurut Surat Pesanan No."
        '        xlWorkSheet4.Cells(rowData, 16) = " :"
        '        xlWorkSheet4.Cells(rowData, 18) = "219/KIT.04.03/WP2B/2016"
        '        xlWorkSheet4.Cells(rowData, 21) = "Tgl.  :"
        '        xlWorkSheet4.Cells(rowData, 22) = " …………………………………"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 3.75

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData, 2)).Merge()
        '        xlWorkSheet4.Cells(rowData, 1) = "No."
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 3), xlWorkSheet4.Cells(rowData, 8)).Merge()
        '        xlWorkSheet4.Cells(rowData, 3) = " N A M A  B A R A N G"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 9), xlWorkSheet4.Cells(rowData, 13)).Merge()
        '        xlWorkSheet4.Cells(rowData, 9) = "No."
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 14), xlWorkSheet4.Cells(rowData + 1, 14)).Merge()
        '        xlWorkSheet4.Cells(rowData, 14) = "Satuan"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 15), xlWorkSheet4.Cells(rowData + 1, 15)).Merge()
        '        xlWorkSheet4.Cells(rowData, 15) = "Banyaknya"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 16), xlWorkSheet4.Cells(rowData + 1, 18)).Merge()
        '        xlWorkSheet4.Cells(rowData, 16) = "Keterangan"
        '        xlWorkSheet4.Cells(rowData, 19) = "Harga Satuan"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 20), xlWorkSheet4.Cells(rowData + 1, 24)).Merge()
        '        xlWorkSheet4.Cells(rowData, 20) = "Jumlah"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData, 2)).Merge()
        '        xlWorkSheet4.Cells(rowData, 1) = "Urut / Tgl."
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 3), xlWorkSheet4.Cells(rowData, 8)).Merge()
        '        xlWorkSheet4.Cells(rowData, 3) = "(ditulis selengkap-lengkapnya)"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 9), xlWorkSheet4.Cells(rowData, 13)).Merge()
        '        xlWorkSheet4.Cells(rowData, 9) = "Normalisasi"
        '        xlWorkSheet4.Cells(rowData, 19) = "Rp"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 4.5

        '        total = 0
        '        For i = 0 To DataGridView1.RowCount - 1
        '            total += DataGridView1.Rows(i).Cells(3).Value
        '        Next

        '        rowData = rowData + 2
        '        rowAwal = rowData
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData + 1, 1)).Merge()

        '        DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
        '        tanggalAkhir = Format(DateTimePicker1.Value.Date, "dd / MM / yyyy")
        '        xlWorkSheet4.Cells(rowData, 1) = "'01 s.d " + tanggalAkhir
        '        xlWorkSheet4.Cells(rowData, 3) = "DO/FAKTUR TERLAMPIR SEBANYAK :"
        '        xlWorkSheet4.Cells(rowData, 4) = ":"
        '        xlWorkSheet4.Cells(rowData, 5) = DataGridView1.RowCount.ToString
        '        xlWorkSheet4.Cells(rowData, 6) = "DO/FAKTUR"
        '        xlWorkSheet4.Cells(rowData, 7) = "="
        '        xlWorkSheet4.Cells(rowData, 8) = total.ToString
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 9), xlWorkSheet4.Cells(rowData, 13)).Merge()
        '        xlWorkSheet4.Cells(rowData, 9) = "0008010001				"
        '        xlWorkSheet4.Cells(rowData, 14) = "Liter"
        '        xlWorkSheet4.Cells(rowData, 15) = total.ToString
        '        xlWorkSheet4.Cells(rowData, 17) = DataGridView1.RowCount.ToString
        '        xlWorkSheet4.Cells(rowData, 18) = "Faktur"

        '        xlWorkSheet4.Cells(rowData + 2, 22) = "-"
        '        rowData = rowData + 9
        '        rowAkhir = rowData

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 4.5

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 5), xlWorkSheet4.Cells(rowData, 6)).Merge()
        '        xlWorkSheet4.Cells(rowData, 5) = "Total ……."
        '        xlWorkSheet4.Cells(rowData, 8) = total.ToString

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 4.5

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 8.25

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData, 3)).Merge()
        '        xlWorkSheet4.Cells(rowData, 1) = "Nota No. ……………………….."
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 5), xlWorkSheet4.Cells(rowData, 6)).Merge()
        '        xlWorkSheet4.Cells(rowData, 5) = " Kode Perkiraan : "
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 7), xlWorkSheet4.Cells(rowData, 8)).Merge()
        '        xlWorkSheet4.Cells(rowData, 7) = "613200202"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 9), xlWorkSheet4.Cells(rowData, 13)).Merge()
        '        xlWorkSheet4.Cells(rowData, 8) = " Catatan-catatan  :"
        '        xlWorkSheet4.Cells(rowData, 14) = "UNIT PENERIMA : PLTD WAENA"
        '        xlWorkSheet4.Cells(rowData, 20) = "Rp."
        '        xlWorkSheet4.Cells(rowData, 22) = "-"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 9

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).RowHeight = 7.5

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData, 4)).Merge()
        '        xlWorkSheet4.Cells(rowData, 1) = "Mengetahui,"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 9), xlWorkSheet4.Cells(rowData, 17)).Merge()
        '        xlWorkSheet4.Cells(rowData, 9) = "Diperiksa,"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 20), xlWorkSheet4.Cells(rowData, 24)).Merge()
        '        xlWorkSheet4.Cells(rowData, 20) = "dibuat,"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData, 4)).Merge()
        '        xlWorkSheet4.Cells(rowData, 1) = "ASMAN SDM, KEU & Adm"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 9), xlWorkSheet4.Cells(rowData, 17)).Merge()
        '        xlWorkSheet4.Cells(rowData, 9) = "SUPERVISOR LOGISTIK"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 20), xlWorkSheet4.Cells(rowData, 24)).Merge()
        '        xlWorkSheet4.Cells(rowData, 20) = "Jr. OFFICER LOGISTIK"

        '        rowData = rowData + 3
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 1), xlWorkSheet4.Cells(rowData, 4)).Merge()
        '        xlWorkSheet4.Cells(rowData, 1) = "BUDI SUTARMANTO"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 9), xlWorkSheet4.Cells(rowData, 17)).Merge()
        '        xlWorkSheet4.Cells(rowData, 9) = "TARMIZI MANILET"
        '        xlWorkSheet4.Range(xlWorkSheet4.Cells(rowData, 20), xlWorkSheet4.Cells(rowData, 24)).Merge()
        '        xlWorkSheet4.Cells(rowData, 20) = "Y U S U F"

        '        rowData = rowData + 1
        '        xlWorkSheet4.Range("A" + rowData.ToString).VerticalAlignment = Excel.Constants.xlCenter
        '        'xlWorkSheet4.Cells.EntireColumn.AutoFit()
        '        range = xlWorkSheet4.Range("A" + rowAwal.ToString, "F" + rowAkhir.ToString)

        '        'Dim borders As Excel.Borders = range.Borders
        '        'Set the thi lines style.
        '        'borders.LineStyle = Excel.XlLineStyle.xlContinuous
        '        'borders.Weight = 2.0R

        '        rowData += 7
        '    End If
        'Next

        tanggalAkhir = Format(DateTimePicker1.Value.Date, "MMMM")
        xlWorkSheet.SaveAs(appPath + "\Excel\Lamp_TUG34_" + tanggalAkhir + ".xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("You can find the file " + appPath + "\Excel\Lamp_TUG34_" + tanggalAkhir + ".xlsx")
    End Sub
    Sub exportLampTUGS3Sheet(ByVal worksheet As Excel.Worksheet, ByVal rowData As Integer, ByVal tglAwal As Integer, ByVal tglAkhir As Integer)
        Dim total As Integer = 0
        Dim tanggalAwal, tanggalAkhir As String
        Dim r As Excel.Range
        Dim fileName As String
        Dim pictureWidth As Integer
        Dim pictureHeight As Integer
        Dim shape As Excel.Shape

        Dim rowAwal As Integer = 0
        Dim rowAkhir As Integer = 0
        Dim range As Excel.Range

        Dim appPath As String = Application.StartupPath()
        fileName = appPath + "\logoPLNcolor.png"

        worksheet.Range("A1").ColumnWidth = 9.29
        worksheet.Range("B1").ColumnWidth = 1
        worksheet.Range("C1").ColumnWidth = 33.43
        worksheet.Range("D1").ColumnWidth = 1
        worksheet.Range("E1").ColumnWidth = 3
        worksheet.Range("F1").ColumnWidth = 9.43
        worksheet.Range("G1").ColumnWidth = 1.57
        worksheet.Range("H1").ColumnWidth = 7
        worksheet.Range("I1").ColumnWidth = 1.29
        worksheet.Range("J1").ColumnWidth = 8.86
        worksheet.Range("K1").ColumnWidth = 0.75
        worksheet.Range("L1").ColumnWidth = 0.92
        worksheet.Range("M1").ColumnWidth = 4
        worksheet.Range("N1").ColumnWidth = 8.43
        worksheet.Range("O1").ColumnWidth = 9
        worksheet.Range("P1").ColumnWidth = 1.29
        worksheet.Range("Q1").ColumnWidth = 3.86
        worksheet.Range("R1").ColumnWidth = 6.57
        worksheet.Range("S1").ColumnWidth = 11.29
        worksheet.Range("T1").ColumnWidth = 8.57
        worksheet.Range("U1").ColumnWidth = 4.71
        worksheet.Range("V1").ColumnWidth = 2.57
        worksheet.Range("W1").ColumnWidth = 4.86
        worksheet.Range("X1").ColumnWidth = 11.43
        worksheet.Range("Y1").ColumnWidth = 1.43
        worksheet.Range("Z1").ColumnWidth = 8.43
        worksheet.Range("AA1").ColumnWidth = 8.43


        rowData = 0
        total = 0
        For xx = 0 To lNamaPerusahaan.Count - 1
            bukaDB()
            DateTimePicker5.Format = DateTimePickerFormat.Custom
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
            tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-" + tglAwal.ToString)
            tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-" + tglAkhir.ToString)

            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO, " +
                                    " tdistribusi.tglMuat,tkendaraan.kapasitasTruk " +
                                    " FROM tdistribusi " +
                                    " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tkendaraan.namaPerusahaan = '" & lNamaPerusahaan.Item(xx).ToString & "' " +
                                    " AND tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "' " +
                                    " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "

            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tdistribusi")
            DataGridView1.DataSource = (DS.Tables("tdistribusi"))
            DataGridView1.Enabled = True
            DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)

            If DataGridView1.RowCount <> 0 Then
                r = worksheet.Cells(rowData + 2, 1)


                shape = worksheet.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      r.Left, r.Top, pictureWidth, pictureHeight)

                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                shape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 4.5

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 18
                worksheet.Cells(rowData, 2) = "PT. PLN (PERSERO) WILAYAH PAPUA"
                worksheet.Cells(rowData, 24) = "TUG. 3"
                worksheet.Cells(rowData, 26) = "MERAH"
                worksheet.Cells(rowData, 27) = "1. Untuk TU Keuangan Persediaan"

                rowData = rowData + 1
                worksheet.Cells(rowData, 2) = "SEKTOR PAPUA & PAPUA BARAT"
                worksheet.Cells(rowData, 26) = "BIRU"
                worksheet.Cells(rowData, 27) = "2. Untuk T.U.G"

                rowData = rowData + 1
                worksheet.Cells(rowData, 2) = "BAGIAN LOGISTIK"
                worksheet.Cells(rowData, 26) = "HIJAU"
                worksheet.Cells(rowData, 27) = "3. Untuk Perbekalan"

                rowData = rowData + 1
                worksheet.Cells(rowData, 23) = "No."
                worksheet.Cells(rowData, 24) = "1339"
                worksheet.Cells(rowData, 26) = "KUNING"
                worksheet.Cells(rowData, 27) = "4. Untuk Pengiriman Barang"

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 26)).Merge()
                worksheet.Cells(rowData, 1) = "BON PENERIMAAN BARANG-BARANG GUDANG"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 21
                worksheet.Cells(rowData, 6) = "  Nomor :"
                worksheet.Cells(rowData, 8) = "339"
                worksheet.Cells(rowData, 9) = "/"
                worksheet.Cells(rowData, 10) = "LOG.00.01"
                worksheet.Cells(rowData, 11) = "/"
                worksheet.Range(worksheet.Cells(rowData, 12), worksheet.Cells(rowData, 18)).Merge()
                worksheet.Cells(rowData, 12) = "BBM. SOLAR  PL. WNA / SP2B / 2016"
                worksheet.Cells(rowData, 20) = "1. Untuk Pengiriman Barang"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 12.75
                worksheet.Range(worksheet.Cells(rowData, 3), worksheet.Cells(rowData, 4)).Merge()
                worksheet.Cells(rowData, 20) = "P L N"
                worksheet.Cells(rowData, 21) = ":"
                worksheet.Cells(rowData, 22) = "WILAYAH PAPUA"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 12.75
                worksheet.Cells(rowData, 20) = "SEKTOR"
                worksheet.Cells(rowData, 21) = ":"
                worksheet.Cells(rowData, 22) = "PAPUA & PAPUA BARAT"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 6.75

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 3

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 8.25

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "Diterima Tgl."
                worksheet.Cells(rowData, 2) = " :"
                DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
                tanggalAkhir = Format(DateTimePicker1.Value.Date, "dd MMMM yyyy")
                worksheet.Cells(rowData, 3) = "'01 s.d " + tanggalAkhir
                worksheet.Cells(rowData, 10) = "Pembelian ditempat, lihat Faktur/Bukti Kas No."
                worksheet.Cells(rowData, 16) = " :"
                worksheet.Cells(rowData, 18) = " …………………………………………………………………………………….."

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "Dari"
                worksheet.Cells(rowData, 2) = " :"
                worksheet.Cells(rowData, 3) = "DEPOT PERTAMINA ( " + lNamaPerusahaan.Item(xx).ToString + " )"
                worksheet.Cells(rowData, 10) = "Diterima Bon Pengiriman No."
                worksheet.Cells(rowData, 16) = " :"
                worksheet.Cells(rowData, 18) = "  ……………………………………………."
                worksheet.Cells(rowData, 21) = "Tgl.  :"
                worksheet.Cells(rowData, 22) = " …………………………………"

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "Dengan"
                worksheet.Cells(rowData, 2) = " :"
                worksheet.Cells(rowData, 10) = "Menurut Surat Pesanan No."
                worksheet.Cells(rowData, 16) = " :"
                worksheet.Cells(rowData, 18) = "219/KIT.04.03/WP2B/2016"
                worksheet.Cells(rowData, 21) = "Tgl.  :"
                worksheet.Cells(rowData, 22) = " …………………………………"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 3.75

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 2)).Merge()
                worksheet.Cells(rowData, 1) = "No."
                worksheet.Range(worksheet.Cells(rowData, 3), worksheet.Cells(rowData, 8)).Merge()
                worksheet.Cells(rowData, 3) = " N A M A  B A R A N G"
                worksheet.Range(worksheet.Cells(rowData, 9), worksheet.Cells(rowData, 13)).Merge()
                worksheet.Cells(rowData, 9) = "No."
                worksheet.Range(worksheet.Cells(rowData, 14), worksheet.Cells(rowData + 1, 14)).Merge()
                worksheet.Cells(rowData, 14) = "Satuan"
                worksheet.Range(worksheet.Cells(rowData, 15), worksheet.Cells(rowData + 1, 15)).Merge()
                worksheet.Cells(rowData, 15) = "Banyaknya"
                worksheet.Range(worksheet.Cells(rowData, 16), worksheet.Cells(rowData + 1, 18)).Merge()
                worksheet.Cells(rowData, 16) = "Keterangan"
                worksheet.Cells(rowData, 19) = "Harga Satuan"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData + 1, 24)).Merge()
                worksheet.Cells(rowData, 20) = "Jumlah"

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 2)).Merge()
                worksheet.Cells(rowData, 1) = "Urut / Tgl."
                worksheet.Range(worksheet.Cells(rowData, 3), worksheet.Cells(rowData, 8)).Merge()
                worksheet.Cells(rowData, 3) = "(ditulis selengkap-lengkapnya)"
                worksheet.Range(worksheet.Cells(rowData, 9), worksheet.Cells(rowData, 13)).Merge()
                worksheet.Cells(rowData, 9) = "Normalisasi"
                worksheet.Cells(rowData, 19) = "Rp"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 4.5

                total = 0
                For i = 0 To DataGridView1.RowCount - 1
                    total += DataGridView1.Rows(i).Cells(3).Value
                Next

                rowData = rowData + 2
                rowAwal = rowData
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData + 1, 1)).Merge()

                DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
                tanggalAkhir = Format(DateTimePicker1.Value.Date, "dd / MM / yyyy")
                worksheet.Cells(rowData, 1) = "'01 s.d " + tanggalAkhir
                worksheet.Cells(rowData, 3) = "DO/FAKTUR TERLAMPIR SEBANYAK :"
                worksheet.Cells(rowData, 4) = ":"
                worksheet.Cells(rowData, 5) = DataGridView1.RowCount.ToString
                worksheet.Cells(rowData, 6) = "DO/FAKTUR"
                worksheet.Cells(rowData, 7) = "="
                worksheet.Cells(rowData, 8) = total.ToString
                worksheet.Range(worksheet.Cells(rowData, 9), worksheet.Cells(rowData, 13)).Merge()
                worksheet.Cells(rowData, 9) = "0008010001				"
                worksheet.Cells(rowData, 14) = "Liter"
                worksheet.Cells(rowData, 15) = total.ToString
                worksheet.Cells(rowData, 17) = DataGridView1.RowCount.ToString
                worksheet.Cells(rowData, 18) = "Faktur"

                worksheet.Cells(rowData + 2, 22) = "-"
                rowData = rowData + 9
                rowAkhir = rowData

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 4.5

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 5), worksheet.Cells(rowData, 6)).Merge()
                worksheet.Cells(rowData, 5) = "Total ……."
                worksheet.Cells(rowData, 8) = total.ToString

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 4.5

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 8.25

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 3)).Merge()
                worksheet.Cells(rowData, 1) = "Nota No. ……………………….."
                worksheet.Range(worksheet.Cells(rowData, 5), worksheet.Cells(rowData, 6)).Merge()
                worksheet.Cells(rowData, 5) = " Kode Perkiraan : "
                worksheet.Range(worksheet.Cells(rowData, 7), worksheet.Cells(rowData, 8)).Merge()
                worksheet.Cells(rowData, 7) = "613200202"
                worksheet.Range(worksheet.Cells(rowData, 9), worksheet.Cells(rowData, 13)).Merge()
                worksheet.Cells(rowData, 8) = " Catatan-catatan  :"
                worksheet.Cells(rowData, 14) = "UNIT PENERIMA : PLTD WAENA"
                worksheet.Cells(rowData, 20) = "Rp."
                worksheet.Cells(rowData, 22) = "-"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 9

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 7.5

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 4)).Merge()
                worksheet.Cells(rowData, 1) = "Mengetahui,"
                worksheet.Range(worksheet.Cells(rowData, 9), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 9) = "Diperiksa,"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData, 24)).Merge()
                worksheet.Cells(rowData, 20) = "dibuat,"

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 4)).Merge()
                worksheet.Cells(rowData, 1) = "ASMAN SDM, KEU & Adm"
                worksheet.Range(worksheet.Cells(rowData, 9), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 9) = "SUPERVISOR LOGISTIK"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData, 24)).Merge()
                worksheet.Cells(rowData, 20) = "Jr. OFFICER LOGISTIK"

                rowData = rowData + 3
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 4)).Merge()
                worksheet.Cells(rowData, 1) = "BUDI SUTARMANTO"
                worksheet.Range(worksheet.Cells(rowData, 9), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 9) = "TARMIZI MANILET"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData, 24)).Merge()
                worksheet.Cells(rowData, 20) = "Y U S U F"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).VerticalAlignment = Excel.Constants.xlCenter
                'worksheet.Cells.EntireColumn.AutoFit()
                range = worksheet.Range("A" + rowAwal.ToString, "F" + rowAkhir.ToString)

                'Dim borders As Excel.Borders = range.Borders
                'Set the thi lines style.
                'borders.LineStyle = Excel.XlLineStyle.xlContinuous
                'borders.Weight = 2.0R

                rowData += 7
            End If
        Next

    End Sub

    Sub exportLampTUGS4Sheet(ByVal worksheet As Excel.Worksheet, ByVal rowData As Integer, ByVal tglAwal As Integer, ByVal tglAkhir As Integer)
        Dim total As Integer = 0
        Dim tanggalAwal, tanggalAkhir As String
        Dim r As Excel.Range
        Dim fileName As String
        Dim pictureWidth As Integer
        Dim pictureHeight As Integer
        Dim shape As Excel.Shape

        Dim rowAwal As Integer = 0
        Dim rowAkhir As Integer = 0
        Dim range As Excel.Range

        Dim appPath As String = Application.StartupPath()
        fileName = appPath + "\logoPLNcolor.png"

        worksheet.Range("A1").ColumnWidth = 6.57
        worksheet.Range("B1").ColumnWidth = 2.71
        worksheet.Range("C1").ColumnWidth = 1.43
        worksheet.Range("D1").ColumnWidth = 4.86
        worksheet.Range("E1").ColumnWidth = 2.43
        worksheet.Range("F1").ColumnWidth = 1
        worksheet.Range("G1").ColumnWidth = 5.14
        worksheet.Range("H1").ColumnWidth = 0.75
        worksheet.Range("I1").ColumnWidth = 2.43
        worksheet.Range("J1").ColumnWidth = 1.57
        worksheet.Range("K1").ColumnWidth = 3
        worksheet.Range("L1").ColumnWidth = 3
        worksheet.Range("M1").ColumnWidth = 0.75
        worksheet.Range("N1").ColumnWidth = 14.86
        worksheet.Range("O1").ColumnWidth = 6.43
        worksheet.Range("P1").ColumnWidth = 0.83
        worksheet.Range("Q1").ColumnWidth = 2.57
        worksheet.Range("R1").ColumnWidth = 1.57
        worksheet.Range("S1").ColumnWidth = 2.57
        worksheet.Range("T1").ColumnWidth = 5.43
        worksheet.Range("U1").ColumnWidth = 4.43
        worksheet.Range("V1").ColumnWidth = 1.57
        worksheet.Range("W1").ColumnWidth = 9.71
        worksheet.Range("X1").ColumnWidth = 1.43
        worksheet.Range("Y1").ColumnWidth = 8.43
        worksheet.Range("Z1").ColumnWidth = 8.43
        worksheet.Range("AA1").ColumnWidth = 0.75


        rowData = 0
        total = 0
        For xx = 0 To lNamaPerusahaan.Count - 1
            bukaDB()
            DateTimePicker5.Format = DateTimePickerFormat.Custom
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            'DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
            tanggalAwal = Format(DateTimePicker5.Value.Date, "yyyy-MM-" + tglAwal.ToString)
            tanggalAkhir = Format(DateTimePicker1.Value.Date, "yyyy-MM-" + tglAkhir.ToString)

            selectDataBase = "SELECT tkendaraan.namaPerusahaan,tdistribusi.NoDO, " +
                                    " tdistribusi.tglMuat,tkendaraan.kapasitasTruk " +
                                    " FROM tdistribusi " +
                                    " JOIN tkendaraan ON tkendaraan.IDKendaraan = tdistribusi.IDKendaraan WHERE tkendaraan.namaPerusahaan = '" & lNamaPerusahaan.Item(xx).ToString & "' " +
                                    " AND tdistribusi.tglMuat BETWEEN '" & tanggalAwal & "' AND '" & tanggalAkhir & "' " +
                                    " AND (tdistribusi.keterangan = 'ACCEPTED' OR tdistribusi.keterangan = 'REJECTED') "

            DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
            DS = New DataSet
            DS.Clear()
            DA.Fill(DS, "tdistribusi")
            DataGridView1.DataSource = (DS.Tables("tdistribusi"))
            DataGridView1.Enabled = True
            DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)

            If DataGridView1.RowCount <> 0 Then
                r = worksheet.Cells(rowData + 2, 1)


                shape = worksheet.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      Microsoft.Office.Core.MsoTriState.msoTrue,
                                                      r.Left, r.Top, pictureWidth, pictureHeight)

                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                shape.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 4.5

                rowData = rowData + 1
                worksheet.Cells(rowData, 2) = "PT. PLN (PERSERO) WILAYAH PAPUA"
                

                rowData = rowData + 1
                worksheet.Cells(rowData, 2) = "WILAYAH PAPUA & PAPUA BARAT"
                worksheet.Cells(rowData, 17) = "'1. Untuk  :"
                worksheet.Cells(rowData, 20) = "Pengiriman Barang"
                worksheet.Cells(rowData, 25) = "PUTIH"
                worksheet.Cells(rowData, 26) = "1. Untuk  :"
                worksheet.Cells(rowData, 28) = "Pengiriman Barang"

                rowData = rowData + 1
                worksheet.Cells(rowData, 2) = "SEKTOR PAPUA & PAPUA BARAT"
                worksheet.Cells(rowData, 28) = "Yang Berkepentingan"

                rowData = rowData + 1
                worksheet.Cells(rowData, 25) = "KUNING"
                worksheet.Cells(rowData, 26) = "2. Untuk  :"
                worksheet.Cells(rowData, 28) = "T.U.K.G"

                rowData = rowData + 1
                worksheet.Cells(rowData, 25) = "BIRU"
                worksheet.Cells(rowData, 26) = "3. Untuk  :"
                worksheet.Cells(rowData, 28) = "T.U.G"

                rowData = rowData + 1
                worksheet.Cells(rowData, 25) = "HIJAU"
                worksheet.Cells(rowData, 26) = "4. Untuk  :"
                worksheet.Cells(rowData, 28) = "T.U. Perbekalan"


                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 26)).Merge()
                worksheet.Range("A" + rowData.ToString).RowHeight = 21
                worksheet.Cells(rowData, 1) = "BERITA ACARA PEMERIKSAAN BARANG-BARANG / SPARE PARTS"

                rowData = rowData + 1

                worksheet.Range(worksheet.Cells(rowData, 5), worksheet.Cells(rowData, 7)).Merge()
                worksheet.Cells(rowData, 5) = "  Nomor :"
                worksheet.Range(worksheet.Cells(rowData, 8), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 8) = "339"
                worksheet.Cells(rowData, 10) = "/"

                worksheet.Range(worksheet.Cells(rowData, 11), worksheet.Cells(rowData, 12)).Merge()
                worksheet.Cells(rowData, 11) = "LOG.00.01"
                worksheet.Cells(rowData, 13) = "/"
                worksheet.Range(worksheet.Cells(rowData, 14), worksheet.Cells(rowData, 19)).Merge()
                worksheet.Cells(rowData, 14) = "BBM. SOLAR  PL. WNA / SP2B / 2016"

                rowData = rowData + 3
                worksheet.Cells(rowData, 1) = "Pada tgl."
                worksheet.Cells(rowData, 2) = ":"
                DateTimePicker1.Value = DateSerial(DateTimePicker5.Value.Year, DateTimePicker5.Value.Month + 1, 0)
                tanggalAkhir = Format(DateTimePicker1.Value.Date, "dd MMMM yyyy")
                worksheet.Cells(rowData, 3) = "'01 s.d " + tanggalAkhir
                worksheet.Range(worksheet.Cells(rowData, 13), worksheet.Cells(rowData, 15)).Merge()
                worksheet.Cells(rowData, 13) = "Para pemeriksa terdiri dari  :"

                rowData = rowData + 2
                worksheet.Range("A" + rowData.ToString).RowHeight = 6

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "NO"
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 2) = "N A M A"
                worksheet.Range(worksheet.Cells(rowData, 10), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 10) = "J A B A T A N"
                worksheet.Range(worksheet.Cells(rowData, 17), worksheet.Cells(rowData, 23)).Merge()
                worksheet.Cells(rowData, 17) = "TANDA TANGAN"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 6

                rowData = rowData + 3
                worksheet.Cells(rowData, 1) = "1"
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 2) = "RAMDAN SIRFEFA"
                worksheet.Range(worksheet.Cells(rowData, 10), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 10) = "(KETUA)"
                worksheet.Cells(rowData, 19) = "1"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData, 21)).Merge()
                worksheet.Cells(rowData, 20) = " ………."

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "2"
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 2) = "Y U S U F"
                worksheet.Range(worksheet.Cells(rowData, 10), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 10) = "(SEKRETARIS)	
                worksheet.Cells(rowData, 22) = "2"
                worksheet.Cells(rowData, 23) = " ………."

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "3"
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 2) = "ZAINAL ARIFIN"
                worksheet.Range(worksheet.Cells(rowData, 10), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 10) = "(ANGGOTA)"
                worksheet.Cells(rowData, 19) = "3"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData, 21)).Merge()
                worksheet.Cells(rowData, 20) = " ………."

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "4"
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 2) = "ISHAK SANTOS WABES"
                worksheet.Range(worksheet.Cells(rowData, 10), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 10) = "(ANGGOTA)"
                worksheet.Cells(rowData, 22) = "4"
                worksheet.Cells(rowData, 23) = " ………."

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "5"
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 2) = "MAX S. IGUGE"
                worksheet.Range(worksheet.Cells(rowData, 10), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 10) = "(ANGGOTA)"
                worksheet.Cells(rowData, 19) = "5"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData, 21)).Merge()
                worksheet.Cells(rowData, 20) = " ………."

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "6"
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 2) = "BAHKTIAR"
                worksheet.Range(worksheet.Cells(rowData, 10), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 10) = "(ANGGOTA)"
                worksheet.Cells(rowData, 22) = "6"
                worksheet.Cells(rowData, 23) = " ………."

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "7"
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 2) = "SRI SUTARTA"
                worksheet.Range(worksheet.Cells(rowData, 10), worksheet.Cells(rowData, 17)).Merge()
                worksheet.Cells(rowData, 10) = "(ANGGOTA)"
                worksheet.Cells(rowData, 19) = "7"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData, 21)).Merge()
                worksheet.Cells(rowData, 20) = " ………."

                rowData = rowData + 3
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 21)).Merge()
                worksheet.Cells(rowData, 1) = "Telah mengadakan pemeriksaan atas barang-barang / spare parts milik PT. PLN (Persero) Wilayah Papua dan Papua Barat"

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 11)).Merge()
                worksheet.Cells(rowData, 1) = "Sektor KIT Jayapura yang berada/diterima dari :  "
                worksheet.Range(worksheet.Cells(rowData, 12), worksheet.Cells(rowData, 16)).Merge()
                worksheet.Cells(rowData, 12) = "DEPOT PERTAMINA ( " + lNamaPerusahaan.Item(xx).ToString + " )"
                worksheet.Range(worksheet.Cells(rowData, 17), worksheet.Cells(rowData, 18)).Merge()
                worksheet.Cells(rowData, 17) = "TANGGAL"
                worksheet.Cells(rowData, 19) = ":"
                worksheet.Cells(rowData, 20) = "'01 s.d " + tanggalAkhir

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 7)).Merge()
                worksheet.Cells(rowData, 1) = "Menurut surat pesanan / BP. No. :"
                worksheet.Cells(rowData, 8) = "219/KIT.04.03/WP2B/2016"
                worksheet.Cells(rowData, 17) = "TANGGAL"
                worksheet.Cells(rowData, 19) = ":"
                worksheet.Cells(rowData, 20) = " ………."


                rowData = rowData + 2
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 6)).Merge()
                worksheet.Cells(rowData, 2) = "GUDANG"
                worksheet.Range(worksheet.Cells(rowData, 7), worksheet.Cells(rowData, 23)).Merge()
                worksheet.Cells(rowData, 7) = "PT. PLN (PERSERO) WIL. PAPUA DAN PAPUA BARAT SEKTOR PAPUA & PAPUA BARAT"

                rowData = rowData + 2
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 6)).Merge()
                worksheet.Cells(rowData, 1) = "Dan menyatakan sebagai berikut :"

                rowData = rowData + 2
                worksheet.Range("A" + rowData.ToString).RowHeight = 5.25

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "No."
                worksheet.Range(worksheet.Cells(rowData, 2), worksheet.Cells(rowData + 1, 14)).Merge()
                worksheet.Cells(rowData, 2) = "NAMA BARANG / SPARE PARTS"
                worksheet.Range(worksheet.Cells(rowData, 15), worksheet.Cells(rowData, 16)).Merge()
                worksheet.Cells(rowData, 15) = "No. "
                worksheet.Range(worksheet.Cells(rowData, 17), worksheet.Cells(rowData + 1, 18)).Merge()
                worksheet.Cells(rowData, 17) = "Satuan"
                worksheet.Range(worksheet.Cells(rowData, 19), worksheet.Cells(rowData + 1, 20)).Merge()
                worksheet.Cells(rowData, 19) = "Banyaknya*)"
                worksheet.Range(worksheet.Cells(rowData, 21), worksheet.Cells(rowData + 1, 22)).Merge()
                worksheet.Cells(rowData, 21) = "Catatan"

                rowData = rowData + 1
                worksheet.Cells(rowData, 1) = "Urut :"
                worksheet.Range(worksheet.Cells(rowData, 15), worksheet.Cells(rowData, 16)).Merge()
                worksheet.Cells(rowData, 15) = "Normalisasi"

                rowData = rowData + 1
                worksheet.Range("A" + rowData.ToString).RowHeight = 5.25


                total = 0
                For i = 0 To DataGridView1.RowCount - 1
                    total += DataGridView1.Rows(i).Cells(3).Value
                Next

                rowData = rowData + 3
                worksheet.Range(worksheet.Cells(rowData, 3), worksheet.Cells(rowData, 5)).Merge()
                worksheet.Cells(rowData, 3) = "BBM SOLAR"
                worksheet.Range(worksheet.Cells(rowData, 15), worksheet.Cells(rowData, 16)).Merge()
                worksheet.Cells(rowData, 15) = "0008010001	"
                worksheet.Range(worksheet.Cells(rowData, 17), worksheet.Cells(rowData, 19)).Merge()
                worksheet.Cells(rowData, 17) = "Liter"
                worksheet.Range(worksheet.Cells(rowData, 20), worksheet.Cells(rowData, 21)).Merge()
                worksheet.Cells(rowData, 20) = total.ToString

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 3), worksheet.Cells(rowData, 8)).Merge()
                worksheet.Cells(rowData, 15) = "Barang tersebut dalam TUG 3 :"

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 3), worksheet.Cells(rowData, 4)).Merge()
                worksheet.Cells(rowData, 3) = "Nomor   :"
                worksheet.Cells(rowData, 5) = "339"
                worksheet.Cells(rowData, 6) = "/"
                worksheet.Cells(rowData, 7) = "LOG.00.01"
                worksheet.Cells(rowData, 8) = "/"
                worksheet.Range(worksheet.Cells(rowData, 9), worksheet.Cells(rowData, 14)).Merge()
                worksheet.Cells(rowData, 9) = "BBM. SOLAR  PL. WNA / SP2B / 2016"

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 3), worksheet.Cells(rowData, 4)).Merge()
                worksheet.Cells(rowData, 3) = "Tanggal :"
                worksheet.Range(worksheet.Cells(rowData, 5), worksheet.Cells(rowData, 9)).Merge()
                worksheet.Cells(rowData, 5) = "'01 s.d " + tanggalAkhir


                rowData = rowData + 5
                worksheet.Range(worksheet.Cells(rowData, 1), worksheet.Cells(rowData, 7)).Merge()
                worksheet.Cells(rowData, 1) = "No. Kode Perkiraan  : ………."
                worksheet.Range(worksheet.Cells(rowData, 13), worksheet.Cells(rowData, 14)).Merge()
                worksheet.Cells(rowData, 13) = "No. Perintah Kerja :"
                worksheet.Range(worksheet.Cells(rowData, 21), worksheet.Cells(rowData, 22)).Merge()
                worksheet.Cells(rowData, 21) = "Fungsi :"
                worksheet.Cells(rowData, 23) = "Pembangkitan"

                rowData = rowData + 3
                worksheet.Range(worksheet.Cells(rowData, 15), worksheet.Cells(rowData, 23)).Merge()
                worksheet.Cells(rowData, 15) = "PT. PLN (PERSERO) WILAYAH PAPUA DAN PAPUA BARAT"

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 15), worksheet.Cells(rowData, 23)).Merge()
                worksheet.Cells(rowData, 15) = "SEKTOR PAPUA DAN PAPUA BARAT"

                rowData = rowData + 1
                worksheet.Range(worksheet.Cells(rowData, 15), worksheet.Cells(rowData, 23)).Merge()
                worksheet.Cells(rowData, 15) = "MANAJER"

                rowData = rowData + 5
                worksheet.Range(worksheet.Cells(rowData, 15), worksheet.Cells(rowData, 23)).Merge()
                worksheet.Cells(rowData, 15) = "PAUL KIRING KALOH, ST"

                rowData += 10
            End If
        Next

    End Sub
End Class