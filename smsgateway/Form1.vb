Imports System.Management

Public Class Form1
    Dim rdata As String = ""
    Public Function ModemsConnected() As String
        Dim modems As String = ""

        Try
            Dim cari As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_POTSModem")
            For Each query As ManagementObject In cari.Get()
                If query("Status") = "OK" Then
                    modems = modems & (query("AttachedTo") & " - " & query("Description") & "***")
                End If
            Next
        Catch ex As ManagementException
            Return ""
        End Try
        Return modems
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ports() As String

        ports = Split(ModemsConnected(), "***")
        For i As Integer = 0 To ports.Length - 2
            ComboBox1.Items.Add(ports(i))
        Next

        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 0
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim comport As String

        comport = "COM1"
        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 0
            Dim pot = ComboBox1.Text.Split("-")
            comport = Trim(pot(0))
        End If

        Try
            With SerialPort1
                .PortName = comport
                .BaudRate = 9600
                .Parity = IO.Ports.Parity.None
                .DataBits = 8
                .StopBits = IO.Ports.StopBits.One
                .Handshake = IO.Ports.Handshake.None
                .ReceivedBytesThreshold = 1
                .NewLine = vbCr
                .ReadTimeout = 1000
                .Open()
            End With

            If SerialPort1.IsOpen Then
                Label1.Text = "Modem terkoneksi"
            Else
                Label1.Text = "Modem tidak terkoneksi"
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            With SerialPort1
                .Write("at+cmgf=1" & vbCrLf)
                Threading.Thread.Sleep(1000)
                .Write("at+cmgs=" & Chr(34) & TextBox1.Text & Chr(34) & vbCrLf)
                .Write(TextBox2.Text & Chr(26))
                Threading.Thread.Sleep(1000)
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SerialPort1_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Dim data As String = ""
        Dim nbyte As Integer = SerialPort1.BytesToRead
        For i As Integer = 1 To nbyte
            data &= Chr(SerialPort1.ReadChar)
        Next
        rdata &= data
    End Sub
End Class
