Imports Spire.Barcode
Imports System.Net.NetworkInformation
Imports System.ComponentModel
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports MySql.Data.MySqlClient


Public Class FormMenu
    Shared Property loggedIn As Integer
    Shared Property statusLog As Integer
    Shared Property idUser As Integer
    Shared Property dbOnline As Boolean
    Shared Property idJenisUser As Integer
    Shared Property lastIDDistribusi As Double
    Shared Property lastIDDistribusiOffline As Double
    Shared Property dateDBOnline As String
    Shared Property timeDBOnline As String
    Shared Property arrName As New List(Of String)
    Shared Property arrValue As New List(Of String)


    Dim excel3() As Integer = {1, 7, 11, 0, 12}
    Dim excel2sheet1() As Integer = {7, 12, 4}
    Dim excel2sheet2() As Integer = {7, 0, 12, 4}
    Dim excel2sheet3() As Integer = {7, 12, 12, 4}

    Dim selectDataBase As String
    Dim simpan As String
    Dim format As String = "yyyy-MM-dd"


    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BarcodeSettings.ApplyKey("KTWS5-S17CF-B3LKE-FXT34-DVRUH")
        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
        Me.Width = screenWidth
        Me.Height = screenHeight
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\icons\IDSF.ico")
        Me.Icon = img
        Me.BackgroundImage = System.Drawing.Image.FromFile(appPath + "\background.jpg")
        Me.BackgroundImageLayout = ImageLayout.Stretch
        PictureBox1.ImageLocation = appPath + ("\icons\Button-Blank-Red-icon.png")

        Dim sData() As String
        Label2.Text = "Checking Connection ..."
        Using sr As New StreamReader("Setting.csv")
            While Not sr.EndOfStream
                sData = sr.ReadLine().Split(","c)

                arrName.Add(sData(0).Trim())
                arrValue.Add(sData(1).Trim())
            End While
        End Using


        
        loggedIn = 0
        TransaksiToolStripMenuItem.Enabled = False
    End Sub

    Private Sub NetworkSettingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NetworkSettingToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Form2.MdiParent = Me
            Form2.Show()
            Form2.Focus()

        End If
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        Frm_Main.Close()
        FormLogin.MdiParent = Me
        FormLogin.Show()
        FormLogin.Focus()
        Dim appPath As String = Application.StartupPath()
        PictureBox1.ImageLocation = appPath + ("\icons\Button-Blank-Red-icon.png")
        Lbl_connection.Text = "DISCONNECT"
        Lbl_User.Text = ""
        Lbl_JenisUser.Text = ""
        LogoutToolStripMenuItem.Text = "Login"
    End Sub

    Private Sub TransaksiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransaksiToolStripMenuItem.Click

        Frm_Main.MdiParent = Me
        Frm_Main.Show()
        Frm_Main.Focus()
        FormMenu.loggedIn = 1
    End Sub

    Private Sub ServerControlToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ServerControlToolStripMenuItem.Click

    End Sub

    Private Sub DataUserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataUserToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Frm_Main.MdiParent = Me
            Frm_Main.Show()
            'Frm_Main.Focus()

            formDataUser.MdiParent = Me
            formDataUser.Show()
            formDataUser.Focus()

        End If
    End Sub

    Private Sub DataKendaraanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataKendaraanToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Frm_Main.MdiParent = Me
            Frm_Main.Show()
            'Frm_Main.Focus()

            formDataKendaraan.MdiParent = Me
            formDataKendaraan.Show()
            formDataKendaraan.Focus()

        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Lbl_Date.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub

    

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub DataJenisUserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataJenisUserToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Frm_Main.MdiParent = Me
            Frm_Main.Show()
            'Frm_Main.Focus()

            formDataJenisUser.MdiParent = Me
            formDataJenisUser.Show()
            formDataJenisUser.Focus()

        End If
    End Sub

    Private Sub DataTujuanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataTujuanToolStripMenuItem.Click
        If (FormMenu.loggedIn = 1) Then
            Frm_Main.MdiParent = Me
            Frm_Main.Show()
            'Frm_Main.Focus()

            formDataTujuan.MdiParent = Me
            formDataTujuan.Show()
            formDataTujuan.Focus()

        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim host As String = FormMenu.arrValue(1) ' use any other machine name
        Dim pingreq As Ping = New Ping()
        Try
            Dim rep As PingReply = pingreq.Send(host)
            If rep.Status = IPStatus.Success Then
                Label2.Text = "Ping = " + rep.RoundtripTime.ToString
                If (Lbl_connection.Text = "DISCONNECT") Then

                    FormMenu.dbOnline = True
                    Label2.Text = "Synchronizing Data"

                    CompareDataBase()
                    'show data 


                    FormLogin.MdiParent = Me
                    FormLogin.Show()
                    FormLogin.Focus()
                    Timer2.Enabled = False
                    Label2.Text = "Online Mode"

                End If
            Else
                Label2.Text = "Server Connection Problem!"
                FormDataBaseOffline.Show()
                Timer2.Enabled = False
            End If
        Catch ex As Exception
            Label2.Text = "Server Connection Problem!"
            FormDataBaseOffline.Show()
            Timer2.Enabled = False
        End Try
    End Sub
    Private Sub CompareDataBase()
        GetLastUpdateID()
        GetLastUpdateIDOffline()
        'ada data baru
        If lastIDDistribusiOffline > lastIDDistribusi Then
            ' save data to list

            For i = lastIDDistribusi To lastIDDistribusiOffline
                lastIDDistribusi += 1
                selectDataBase = "SELECT * FROM tdistribusi WHERE IDDistribusi ='" & lastIDDistribusi & "' "
                FormMenu.dbOnline = False
                bukaDB()
                DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
                DS = New DataSet
                DT = New DataTable
                DS.Clear()
                DA.Fill(DT)
                If DT.Rows.Count > 0 Then
                    Dim tglMuat As DateTime = Convert.ToDateTime(DT.Rows(0).Item("tglMuat").ToString)
                    simpan = "INSERT INTO tdistribusi(IDDistribusi,IDKendaraan,IDUser,IDTujuan, " +
                            "NoDO,wktMuat,wktSampai,tglMuat,tglSampai, " +
                            " dataBarcode,Keterangan, tempatLoading, IDUserClient) " +
                         "VALUES ('" & DT.Rows(0).Item("IDDistribusi").ToString & "', '" & DT.Rows(0).Item("IDKendaraan").ToString & "', '" & DT.Rows(0).Item("IDUser").ToString & "', '" & DT.Rows(0).Item("IDTujuan").ToString & "', " +
                            " '" & DT.Rows(0).Item("NoDO").ToString & "', '" & DT.Rows(0).Item("wktMuat").ToString & "', '" & DT.Rows(0).Item("wktSampai").ToString & "', '" & tglMuat.ToString(format) & "',  '" & DT.Rows(0).Item("tglSampai").ToString & "', " +
                            " '" & DT.Rows(0).Item("dataBarcode").ToString & "', '" & DT.Rows(0).Item("Keterangan").ToString & "',  '" & DT.Rows(0).Item("tempatLoading").ToString & "',  '" & DT.Rows(0).Item("IDUserClient").ToString & "')"
                    FormMenu.dbOnline = True
                    jalankansql(simpan)
                End If

                lastIDDistribusi += 1
            Next


            ' update data to server


        End If
        'ada update data

    End Sub
    Private Sub GetLastUpdateID()
        selectDataBase = "SELECT * FROM tdistribusi ORDER BY IDDistribusi DESC LIMIT 1"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            lastIDDistribusi = DT.Rows(0).Item("IDDistribusi")
        End If
        konek.Close()
        konek.Dispose()
    End Sub
    Private Sub GetLastUpdateIDOffline()
        selectDataBase = "SELECT * FROM tdistribusi ORDER BY IDDistribusi DESC LIMIT 1"
        bukaDBoffline()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            lastIDDistribusiOffline = DT.Rows(0).Item("IDDistribusi")
        End If
        konek.Close()
    End Sub
    Sub SynchronizingServer()

    End Sub

    Sub IsiGrid()

        selectDataBase = "SELECT * FROM attachedTable " +
                            "WHERE col1 NOT IN( SELECT lt.col1 FROM localTable as lt)"
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tdistribusi")
        'DataGridView1.DataSource = DS.Tables("tdistribusi")
        'DataGridView1.Sort(DataGridView1.Columns(6), ListSortDirection.Descending)
        'With DataGridView1
        '    .RowHeadersVisible = False
        '    .Columns(0).HeaderCell.Value = "Transportir"
        '    .Columns(1).HeaderCell.Value = "Nomor DO"
        '    .Columns(2).HeaderCell.Value = "Tempat Tujuan"
        '    .Columns(3).HeaderCell.Value = "No Pol Kendaraan"
        '    .Columns(4).HeaderCell.Value = "Keterangan"
        '    .Columns(5).HeaderCell.Value = "User Server"
        '    .Columns(6).HeaderCell.Value = "Barcode"
        '    .Columns(7).HeaderCell.Value = "Tanggal Pengiriman"
        '    .Columns(8).HeaderCell.Value = "Waktu Pengiriman"
        '    .Columns(9).HeaderCell.Value = "Tanggal Sampai"
        '    .Columns(10).HeaderCell.Value = "Waktu Sampai"
        '    .Columns(11).HeaderCell.Value = "Call Center"
        '    .Columns(12).HeaderCell.Value = "Liter"
        '    .Columns(13).HeaderCell.Value = "Tempat Loading"

        'End With
    End Sub

    Private Sub Level3ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer


        Frm_Main.DataGridView1.Sort(Frm_Main.DataGridView1.Columns(0), ListSortDirection.Ascending)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")


        'For i = 0 To DataGridView1.RowCount - 2
        '    For j = 0 To DataGridView1.ColumnCount - 1
        '        For k As Integer = 1 To DataGridView1.Columns.Count
        '            xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
        '            xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
        '        Next
        '    Next
        'Next
        For i = 0 To Frm_Main.DataGridView1.RowCount - 1
            For j = 0 To 4
                xlWorkSheet.Cells(1, 1) = Frm_Main.DataGridView1.Columns(excel3(0)).HeaderText

                xlWorkSheet.Cells(1, 2) = Frm_Main.DataGridView1.Columns(excel3(1)).HeaderText

                xlWorkSheet.Cells(1, 3) = Frm_Main.DataGridView1.Columns(excel3(2)).HeaderText

                xlWorkSheet.Cells(1, 4) = Frm_Main.DataGridView1.Columns(excel3(3)).HeaderText

                xlWorkSheet.Cells(1, 5) = Frm_Main.DataGridView1.Columns(excel3(4)).HeaderText

                xlWorkSheet.Cells(i + 2, j + 1) = Frm_Main.DataGridView1.Rows(i).Cells(excel3(j)).Value

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

    Private Sub Level21ToolStripMenuItem_Click(sender As Object, e As EventArgs)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlWorkSheet2 As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer


        Frm_Main.DataGridView1.Sort(Frm_Main.DataGridView1.Columns(0), ListSortDirection.Ascending)
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.ActiveSheet, Excel.Worksheet)


        For i = 0 To Frm_Main.DataGridView1.RowCount - 1
            For j = 0 To 2
                xlWorkSheet.Cells(1, 1) = Frm_Main.DataGridView1.Columns(excel2sheet1(0)).HeaderText

                xlWorkSheet.Cells(1, 2) = Frm_Main.DataGridView1.Columns(excel2sheet1(1)).HeaderText

                xlWorkSheet.Cells(1, 3) = Frm_Main.DataGridView1.Columns(excel2sheet1(2)).HeaderText

                xlWorkSheet.Cells(i + 2, j + 1) = Frm_Main.DataGridView1.Rows(i).Cells(excel2sheet1(j)).Value

            Next
        Next
        xlWorkSheet.Cells.EntireColumn.AutoFit()
        xlWorkSheet2 = xlWorkBook.Worksheets.Add(, xlWorkSheet, , )

        xlWorkSheet = xlWorkBook.Sheets("sheet2")
        For i = 0 To Frm_Main.DataGridView1.RowCount - 1
            For j = 0 To 3
                xlWorkSheet.Cells(1, 1) = Frm_Main.DataGridView1.Columns(excel2sheet2(0)).HeaderText

                xlWorkSheet.Cells(1, 2) = Frm_Main.DataGridView1.Columns(excel2sheet2(1)).HeaderText

                xlWorkSheet.Cells(1, 3) = Frm_Main.DataGridView1.Columns(excel2sheet2(2)).HeaderText

                xlWorkSheet.Cells(1, 4) = Frm_Main.DataGridView1.Columns(excel2sheet2(3)).HeaderText

                xlWorkSheet.Cells(i + 2, j + 1) = Frm_Main.DataGridView1.Rows(i).Cells(excel2sheet2(j)).Value

            Next
        Next
        xlWorkSheet2.Cells.EntireColumn.AutoFit()
        Dim appPath As String = Application.StartupPath()
        xlWorkSheet.SaveAs(appPath + "\Excel\Level2.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("You can find the file " + appPath + "\Excel\Level2.xlsx")

    End Sub

    Private Sub ExportDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportDataToolStripMenuItem.Click
        formDataExport.MdiParent = Me
        formDataExport.Show()
        formDataExport.Focus()
    End Sub

    Private Sub FormMenu_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        Dim pesan As String
        If (FormMenu.dbOnline = True) Then
            pesan = MsgBox("Update local database from Server?", MsgBoxStyle.YesNo, "IDSF")
            If pesan = MsgBoxResult.No Then
                Exit Sub
            End If

            Dim appPath As String = Application.StartupPath()
            Dim file As String = appPath + "\" + FormMenu.arrValue(8)

            Dim connect As MySqlConnection

            connect = New MySqlConnection
            If (FormMenu.dbOnline = True) Then
                connect.ConnectionString = "server=103.247.8.246;user=admin_idsf;password=123456;database=idsf;Convert Zero Datetime=True;"
            Else
                connect.ConnectionString = "server=localhost;userid=root;password=r7pqv6s6Xc9QbZKK;database=idsf"
            End If

            Dim cmd As MySqlCommand
            Dim mb As MySqlBackup
            cmd = New MySqlCommand
            mb = New MySqlBackup(cmd)
            cmd.Connection = connect
            connect.Open()
            mb.ExportInfo.AddCreateDatabase = True
            mb.ExportInfo.ExportTableStructure = True
            mb.ExportInfo.ExportRows = True
            Try
                mb.ExportToFile(file)
            Catch ex As Exception

            End Try

            connect.Close()

            FormMenu.dbOnline = False
            If (FormMenu.dbOnline = True) Then
                connect.ConnectionString = "server=" + FormMenu.arrValue(1) + ";userid=admin_idsf;password=123456;database=idsf"
            Else
                connect.ConnectionString = "server=localhost;userid=root;pwd=r7pqv6s6Xc9QbZKK;database=idsf;"
            End If
            cmd.Connection = connect
            connect.Open()
            mb.ImportFromFile(file)
            connect.Close()

        End If

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
End Class