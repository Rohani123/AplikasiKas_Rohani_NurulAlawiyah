Imports System.Data.Odbc
Public Class Form1
    Dim saldoSekarang As Integer

    Sub TampilGrid()
        bukakoneksi()

        DA = New OdbcDataAdapter("select * from tbl_kas ", CONN)
        DS = New DataSet
        DA.Fill(DS, "tbl_kas")
        DataGridView1.DataSource = DS.Tables("tbl_kas")

        tutupkoneksi()
    End Sub

    Sub getSaldoSekarang()
        bukakoneksi()

        CMD = New OdbcCommand("select * from tbl_kas order by id_kas desc limit 1", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            lblSaldoUtama.Text = "0"
        Else
            lblSaldoUtama.Text = RD.Item("saldo_sekarang")
            saldoSekarang = RD.Item("saldo_sekarang")
        End If

        tutupkoneksi()
    End Sub

    Sub MunculCombo()
        comboJenis.Items.Add("Masuk")
        comboJenis.Items.Add("Keluar")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        TampilGrid()
        MunculCombo()
        getSaldoSekarang()
        txtID.Focus()
    End Sub

    Sub KosongkanData()
        txtID.Text = ""
        comboJenis.Text = ""
        txtJumlah.Text = ""
        txtKet.Text = ""
    End Sub

    Private Sub btnSimpan_Click(sender As Object, e As EventArgs) Handles btnSimpan.Click
        If txtID.Text = "" Or dateKas.Text = "" Or comboJenis.Text = "" Or txtJumlah.Text = "" Or txtKet.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
        Else
            bukakoneksi()

            If comboJenis.Text = "Masuk" Then
                saldoSekarang = saldoSekarang + txtJumlah.Text
                Dim simpan As String = "insert into tbl_kas values ('" & txtID.Text & "','" & dateKas.Text & "','" & comboJenis.Text & "','" & txtJumlah.Text & "','" & saldoSekarang &
                    "','" & txtKet.Text & "')"

                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Input data berhasil")
                lblSaldoUtama.Text = saldoSekarang

            ElseIf comboJenis.Text = "Keluar" Then
                If saldoSekarang < txtJumlah.Text Then
                    MsgBox("Saldo Tidak Cukup!")
                Else
                    saldoSekarang = saldoSekarang - txtJumlah.Text
                    lblSaldoUtama.Text = saldoSekarang
                    Dim simpan As String = "insert into tbl_kas values ('" & txtID.Text & "','" & dateKas.Text & "','" & comboJenis.Text & "','" & txtJumlah.Text & "','" & saldoSekarang &
                    "','" & txtKet.Text & "')"

                    CMD = New OdbcCommand(simpan, CONN)
                    CMD.ExecuteNonQuery()
                    MsgBox("Input data berhasil")

                End If
            End If

            TampilGrid()
            KosongkanData()
            tutupkoneksi()
        End If
    End Sub

    Private Sub txtID_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtID.KeyPress
        txtID.MaxLength = 6

        If e.KeyChar = Chr(13) Then
            bukakoneksi()
            CMD = New OdbcCommand("select * from tbl_kas where id ='" & txtID.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("ID tidak ada, Silahkan coba lagi!")
                txtID.Focus()
            Else
                txtID.Text = RD.Item("id")
                dateKas.Text = RD.Item("tanggal")
                comboJenis.Text = RD.Item("jenis")
                txtJumlah.Text = RD.Item("jumlah")
                txtKet.Text = RD.Item("keterangan")
                txtID.Focus()
            End If
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        bukakoneksi()
        Dim edit As String = "update tbl_kas set
        tanggal='" & dateKas.Text & "',
        jenis='" & comboJenis.Text & "',
        jumlah='" & txtJumlah.Text & "',
        keterangan='" & txtKet.Text & "'
        where id ='" & txtID.Text & "'"

        CMD = New OdbcCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil diUpdate")
        TampilGrid()
        KosongkanData()
        tutupkoneksi()
    End Sub

    Private Sub btnHapus_Click(sender As Object, e As EventArgs) Handles btnHapus.Click
        If txtID.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan dihapus dengan Masukan ID dan Enter")
        Else
            If MessageBox.Show("Yakin akan dihapus..? ", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) Then
                bukakoneksi()
                Dim hapus As String = "delete from tbl_kas where id ='" & txtID.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)

                TampilGrid()
                KosongkanData()
                tutupkoneksi()
            End If
        End If
    End Sub
End Class
