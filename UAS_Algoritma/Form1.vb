Imports System
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms

Public Class Form1
    Inherits Form

    Private dgvNilai As DataGridView
    Private btnLoad As Button
    Private btnProses As Button
    Private ofd As OpenFileDialog

    Private lblRata As Label
    Private lblLulus As Label
    Private lblTidakLulus As Label

    Private txtRata As TextBox
    Private txtLulus As TextBox
    Private txtTidakLulus As TextBox

    Public Sub New()
        InitializeComponent()
    End Sub

    '=========================
    ' INISIALISASI KOMPONEN UI
    '=========================
    Private Sub InitializeComponent()
        Me.Text = "UAS Algoritma - Pengolah Nilai Mahasiswa"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(900, 560)
        Me.MinimumSize = New Size(760, 500)

        ' --- OpenFileDialog ---
        ofd = New OpenFileDialog()
        ofd.Title = "Pilih file data (CSV/TXT)"
        ofd.Filter = "CSV File (*.csv)|*.csv|Text File (*.txt)|*.txt|All Files (*.*)|*.*"

        ' --- Buttons ---
        btnLoad = New Button()
        btnLoad.Name = "btnLoad"
        btnLoad.Text = "Load File"
        btnLoad.Width = 120

        btnProses = New Button()
        btnProses.Name = "btnProses"
        btnProses.Text = "Proses"
        btnProses.Width = 120

        AddHandler btnLoad.Click, AddressOf btnLoad_Click
        AddHandler btnProses.Click, AddressOf btnProses_Click

        Dim panelTop As New FlowLayoutPanel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 48
        panelTop.Padding = New Padding(10, 10, 10, 0)
        panelTop.FlowDirection = FlowDirection.LeftToRight
        panelTop.WrapContents = False
        panelTop.Controls.Add(btnLoad)
        panelTop.Controls.Add(btnProses)

        ' --- DataGridView ---
        dgvNilai = New DataGridView()
        dgvNilai.Name = "dgvNilai"
        dgvNilai.Dock = DockStyle.Fill
        dgvNilai.AllowUserToAddRows = False
        dgvNilai.AllowUserToDeleteRows = False
        dgvNilai.ReadOnly = True
        dgvNilai.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvNilai.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        ' --- Summary panel ---
        lblRata = New Label()
        lblRata.Text = "Rata-rata:"
        lblRata.AutoSize = True

        txtRata = New TextBox()
        txtRata.Name = "txtRata"
        txtRata.Width = 120
        txtRata.ReadOnly = True

        lblLulus = New Label()
        lblLulus.Text = "Jumlah Lulus:"
        lblLulus.AutoSize = True

        txtLulus = New TextBox()
        txtLulus.Name = "txtLulus"
        txtLulus.Width = 120
        txtLulus.ReadOnly = True

        lblTidakLulus = New Label()
        lblTidakLulus.Text = "Jumlah Tidak Lulus:"
        lblTidakLulus.AutoSize = True

        txtTidakLulus = New TextBox()
        txtTidakLulus.Name = "txtTidakLulus"
        txtTidakLulus.Width = 120
        txtTidakLulus.ReadOnly = True

        Dim panelBottom As New FlowLayoutPanel()
        panelBottom.Dock = DockStyle.Bottom
        panelBottom.Height = 60
        panelBottom.Padding = New Padding(10, 10, 10, 10)
        panelBottom.FlowDirection = FlowDirection.LeftToRight
        panelBottom.WrapContents = True

        panelBottom.Controls.Add(lblRata)
        panelBottom.Controls.Add(txtRata)
        panelBottom.Controls.Add(New Label() With {.Width = 20})
        panelBottom.Controls.Add(lblLulus)
        panelBottom.Controls.Add(txtLulus)
        panelBottom.Controls.Add(New Label() With {.Width = 20})
        panelBottom.Controls.Add(lblTidakLulus)
        panelBottom.Controls.Add(txtTidakLulus)

        ' --- Add to form ---
        Me.Controls.Add(dgvNilai)
        Me.Controls.Add(panelBottom)
        Me.Controls.Add(panelTop)
    End Sub

    '=========================
    ' EVENT: LOAD FILE
    '=========================
    Private Sub btnLoad_Click(sender As Object, e As EventArgs)
        LoadDataDariFile()
    End Sub

    '=========================
    ' EVENT: PROSES DATA
    '=========================
    Private Sub btnProses_Click(sender As Object, e As EventArgs)
        ProsesNilaiDanRingkasan()
    End Sub

    '=========================================
    ' PROCEDURE 1: MEMBACA FILE (CSV/TXT)
    '=========================================
    Private Sub LoadDataDariFile()
        Try
            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            dgvNilai.Columns.Clear()
            dgvNilai.Rows.Clear()

            dgvNilai.Columns.Add("NIM", "NIM")
            dgvNilai.Columns.Add("Nama", "Nama")
            dgvNilai.Columns.Add("Nilai", "Nilai")
            dgvNilai.Columns.Add("Grade", "Grade")
            dgvNilai.Columns.Add("Status", "Status")

            Dim lines() As String = File.ReadAllLines(ofd.FileName)

            ' PERULANGAN: baca baris file
            For i As Integer = 0 To lines.Length - 1
                Dim line As String = lines(i).Trim()
                If line = "" Then Continue For

                ' Skip header jika ada
                If i = 0 AndAlso line.ToLower().Contains("nim") AndAlso line.ToLower().Contains("nilai") Then
                    Continue For
                End If

                Dim parts() As String = line.Split(","c)
                If parts.Length < 3 Then
                    parts = line.Split(";"c)
                End If

                If parts.Length >= 3 Then
                    Dim nim As String = parts(0).Trim()
                    Dim nama As String = parts(1).Trim()
                    Dim nilaiStr As String = parts(2).Trim()

                    dgvNilai.Rows.Add(nim, nama, nilaiStr, "", "")
                End If
            Next

            MessageBox.Show("Data berhasil dimuat.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Gagal membaca file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    '================================================
    ' PROCEDURE 2: PROSES (GRADE, STATUS, RINGKASAN)
    '================================================
    Private Sub ProsesNilaiDanRingkasan()
        If dgvNilai.Rows.Count = 0 Then
            MessageBox.Show("Data masih kosong. Silakan Load File terlebih dahulu.", "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim total As Double = 0
        Dim jumlah As Integer = 0
        Dim lulus As Integer = 0
        Dim tidakLulus As Integer = 0

        ' PERULANGAN: proses tiap baris
        For Each row As DataGridViewRow In dgvNilai.Rows
            If row.IsNewRow Then Continue For

            Dim nilai As Double
            If Double.TryParse(Convert.ToString(row.Cells("Nilai").Value), nilai) Then
                Dim grade As String = HitungGrade(nilai)
                Dim status As String

                ' PERCABANGAN (If)
                If nilai >= 60 Then
                    status = "Lulus"
                    lulus += 1
                Else
                    status = "Tidak Lulus"
                    tidakLulus += 1
                End If

                row.Cells("Grade").Value = grade
                row.Cells("Status").Value = status

                total += nilai
                jumlah += 1
            Else
                row.Cells("Grade").Value = "-"
                row.Cells("Status").Value = "Nilai tidak valid"
            End If
        Next

        Dim rata As Double = 0
        If jumlah > 0 Then rata = total / jumlah

        txtRata.Text = rata.ToString("0.00")
        txtLulus.Text = lulus.ToString()
        txtTidakLulus.Text = tidakLulus.ToString()
    End Sub

    '=========================
    ' FUNCTION: HITUNG GRADE
    '=========================
    Private Function HitungGrade(nilai As Double) As String
        ' PERCABANGAN (Select Case)
        Select Case nilai
            Case Is >= 85
                Return "A"
            Case Is >= 75
                Return "B"
            Case Is >= 60
                Return "C"
            Case Is >= 50
                Return "D"
            Case Else
                Return "E"
        End Select
    End Function

End Class
