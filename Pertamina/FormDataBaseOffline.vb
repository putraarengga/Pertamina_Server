Imports System.Data.OleDb
Imports MySql.Data.MySqlClient
Imports System.Net

Public Class FormDataBaseOffline
    Dim connect As MySqlConnection
    Dim command As MySqlCommand
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim simpan As String
    Dim tmpAbsensi, tmpDate, tmpTime As String
    Dim countAbsensi As Integer
    Dim tmpCount As Integer
    Dim selectDataBase As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        FormMenu.Label2.Text = "Checking Connection ..."
        FormMenu.Timer2.Enabled = True
        Close()

    End Sub
    Private Sub GetIDUser()
        selectDataBase = "SELECT * FROM tdatauser WHERE NamaLengkap='" & FormMenu.Lbl_User.Text & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            FormMenu.idUser = DT.Rows(0).Item("IDUser")
        End If
    End Sub
    Private Sub GetJenisUser()
        selectDataBase = "SELECT * FROM tjenisuser WHERE IDJenisUser='" & FormMenu.idJenisUser & "' "
        bukaDB()
        DA = New Odbc.OdbcDataAdapter(selectDataBase, konek)
        DS = New DataSet
        DT = New DataTable
        DS.Clear()
        DA.Fill(DT)
        If DT.Rows.Count > 0 Then
            FormMenu.Lbl_JenisUser.Text = DT.Rows(0).Item("JenisUser")
        End If
    End Sub
    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim appPath As String = Application.StartupPath()
        Dim img As New System.Drawing.Icon(appPath + "\icons\IDSF.ico")
        Me.Icon = img
        Me.BackgroundImage = System.Drawing.Image.FromFile(appPath + "\black.jpg")
        Me.BackgroundImageLayout = ImageLayout.Stretch
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
            'MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub

    Private Sub periksasql(ByVal sQL As String)
        Dim objcmd As New System.Data.Odbc.OdbcCommand
        bukaDB()
        Try
            objcmd.Connection = konek
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQL
            countAbsensi = Convert.ToInt16(objcmd.ExecuteScalar())
            objcmd.Dispose()
            'MsgBox("Data sudah disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak bisa menyimpan data ke server" & ex.Message)
        End Try
    End Sub
    Public Shared Function CheckForInternetConnection() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://" + FormMenu.arrValue(1))
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub TB_Password_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            btn_login.PerformClick()
        End If
    End Sub

    Private Sub TB_Nama_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            btn_login.PerformClick()
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        FormMenu.Label2.Text = "Offline Mode"
        FormMenu.dbOnline = False
        FormLogin.Show()
        FormLogin.Focus()
        Close()
    End Sub
End Class
