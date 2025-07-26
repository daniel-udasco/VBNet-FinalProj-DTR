Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports ZXing
Imports AForge.Video
Imports AForge.Video.DirectShow
Imports System.Drawing.Drawing2D


Public Class Form1

    Dim camera As FilterInfoCollection
    Dim video As VideoCaptureDevice

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadDTRTable()

        camera = New FilterInfoCollection(FilterCategory.VideoInputDevice)
        For Each cam As FilterInfo In camera
            ComboBox1.Items.Add(cam.Name)
        Next
        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 0
        End If
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        video = New VideoCaptureDevice(camera(ComboBox1.SelectedIndex).MonikerString)
        AddHandler video.NewFrame, AddressOf Capture
        video.Start()
        Timer1.Start()
    End Sub

    Private Sub Capture(sender As Object, eventArgs As NewFrameEventArgs)
        PictureBox1.Image = DirectCast(eventArgs.Frame.Clone(), Bitmap)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If PictureBox1.Image IsNot Nothing Then
            Dim reader As New BarcodeReader()
            reader.Options.TryHarder = True
            Dim result As Result = reader.Decode(DirectCast(PictureBox1.Image, Bitmap))

            If result IsNot Nothing Then
                Dim scannedID As String = result.Text.Trim()
                TextBox1.Text = scannedID

                Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\CHUWI\Desktop\VBNet_FinalProject\Database\DTR.accdb"
                Try
                    Using conn As New OleDbConnection(connStr)
                        conn.Open()

                        Dim cmd As New OleDbCommand("SELECT Name FROM tblCode WHERE TRIM(TUPVID) = ?", conn)
                        cmd.Parameters.AddWithValue("?", scannedID)
                        Dim readerDB As OleDbDataReader = cmd.ExecuteReader()

                        If readerDB.Read() Then
                            TextBox2.Text = readerDB("Name").ToString()
                        Else
                            TextBox2.Text = "Unknown ID"
                        End If
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Database error: " & ex.Message)
                    TextBox2.Text = "Connection Error"
                End Try

                Timer1.Stop()
            End If
        End If
    End Sub

    Private Sub btnTimeIn_Click(sender As Object, e As EventArgs) Handles btnTimeIn.Click
        Dim scannedID As String = TextBox1.Text.Trim()
        Dim name As String = TextBox2.Text.Trim()
        If scannedID = "" Or name = "" Or name = "Unknown ID" Then
            MessageBox.Show("Invalid scan. Please scan a valid QR code.")
            Exit Sub
        End If

        Dim currentDate As String = DateTime.Now.ToString("yyyy-MM-dd")
        Dim currentTimeIn As String = DateTime.Now.ToString("hh:mm:ss tt")
        Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\CHUWI\Desktop\VBNet_FinalProject\Database\DTR.accdb"

        Using conn As New OleDbConnection(connStr)
            conn.Open()

            Dim checkCmd As New OleDbCommand("SELECT COUNT(*) FROM tblDTR WHERE TUPVID = ? AND [Date] = ?", conn)
            checkCmd.Parameters.AddWithValue("?", scannedID)
            checkCmd.Parameters.AddWithValue("?", currentDate)
            Dim exists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

            If exists > 0 Then
                MessageBox.Show("Already timed in today.")
                Exit Sub
            End If

            Dim insertCmd As New OleDbCommand("INSERT INTO tblDTR (TUPVID, Name, [Date], TimeIn) VALUES (?, ?, ?, ?)", conn)
            insertCmd.Parameters.AddWithValue("?", scannedID)
            insertCmd.Parameters.AddWithValue("?", name)
            insertCmd.Parameters.AddWithValue("?", currentDate)
            insertCmd.Parameters.AddWithValue("?", currentTimeIn)
            insertCmd.ExecuteNonQuery()
        End Using

        LoadDTRTable()
        DataGridView1.Refresh()

        MessageBox.Show("Time In recorded.")

    End Sub

    Private Sub btnTimeOut_Click(sender As Object, e As EventArgs) Handles btnTimeOut.Click
        Dim scannedID As String = TextBox1.Text.Trim()
        Dim currentDate As String = DateTime.Now.ToString("yyyy-MM-dd")
        Dim currentTimeOut As String = DateTime.Now.ToString("hh:mm:ss tt")
        Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\CHUWI\Desktop\VBNet_FinalProject\Database\DTR.accdb"

        Using conn As New OleDbConnection(connStr)
            conn.Open()

            Dim updateCmd As New OleDbCommand("UPDATE tblDTR SET TimeOut = ? WHERE TUPVID = ? AND [Date] = ?", conn)
            updateCmd.Parameters.AddWithValue("?", currentTimeOut)
            updateCmd.Parameters.AddWithValue("?", scannedID)
            updateCmd.Parameters.AddWithValue("?", currentDate)
            updateCmd.ExecuteNonQuery()

            Dim selectCmd As New OleDbCommand("SELECT ID, TimeIn, TimeOut FROM tblDTR WHERE TUPVID = ? AND [Date] = ?", conn)
            selectCmd.Parameters.AddWithValue("?", scannedID)
            selectCmd.Parameters.AddWithValue("?", currentDate)

            Dim reader As OleDbDataReader = selectCmd.ExecuteReader()
            If reader.Read() Then
                Dim id As Integer = Convert.ToInt32(reader("ID"))
                Dim timeIn As DateTime = Convert.ToDateTime(reader("TimeIn"))
                Dim timeOut As DateTime = Convert.ToDateTime(reader("TimeOut"))

                Dim totalHours As Double = Math.Round((timeOut - timeIn).TotalHours, 2)
                reader.Close()

                Dim updateHoursCmd As New OleDbCommand("UPDATE tblDTR SET TotalHours = ? WHERE ID = ?", conn)
                updateHoursCmd.Parameters.AddWithValue("?", totalHours)
                updateHoursCmd.Parameters.AddWithValue("?", id)
                updateHoursCmd.ExecuteNonQuery()
            Else
                reader.Close()
            End If
        End Using

        MessageBox.Show("Time Out recorded.")
        LoadDTRTable()
    End Sub


    Private Sub LoadDTRTable()
        Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\CHUWI\Desktop\VBNet_FinalProject\Database\DTR.accdb"
        Using conn As New OleDbConnection(connStr)
            Dim query As String = "SELECT Name, TUPVID, Format(TimeIn, 'hh:nn AM/PM') AS TimeIn, Format(TimeOut, 'hh:nn AM/PM') AS TimeOut, 
                               Round(TotalHours, 2) AS TotalHours, Format([Date], 'dd/mm/yyyy') AS [Date] 
                               FROM tblDTR ORDER BY [Date] ASC, TimeIn ASC"

            Dim adapter As New OleDbDataAdapter(query, conn)
            Dim table As New DataTable()
            adapter.Fill(table)

            With DataGridView1
                .Columns.Clear()
                .AutoGenerateColumns = False

                .Columns.Add("TUPVID", "TUPVID")
                .Columns.Add("Name", "Name")
                .Columns.Add("Date", "Date")
                .Columns.Add("TimeIn", "TimeIn")
                .Columns.Add("TimeOut", "TimeOut")
                .Columns.Add("TotalHours", "TotalHours")

                .Columns("TUPVID").DataPropertyName = "TUPVID"
                .Columns("Name").DataPropertyName = "Name"
                .Columns("Date").DataPropertyName = "Date"
                .Columns("TimeIn").DataPropertyName = "TimeIn"
                .Columns("TimeOut").DataPropertyName = "TimeOut"
                .Columns("TotalHours").DataPropertyName = "TotalHours"

                .Columns(0).Width = 85
                .Columns(1).Width = 130
                .Columns(2).Width = 70
                .Columns(3).Width = 70
                .Columns(4).Width = 70
                .Columns(5).Width = 100
            End With

            DataGridView1.DataSource = table
        End Using
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If video IsNot Nothing AndAlso video.IsRunning Then
            video.SignalToStop()
            video.WaitForStop()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles GenerateReport.Click
        Dim result = MessageBox.Show("Do you want to export the DTR table to Excel?", "Export", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.No Then Exit Sub

        Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\CHUWI\Desktop\VBNet_FinalProject\Database\DTR.accdb"
        Dim dt As New DataTable()

        Try
            Using conn As New OleDbConnection(connStr)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT * FROM tblDTR", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                adapter.Fill(dt)
            End Using

            Dim excelApp As New Excel.Application
            Dim workbook As Excel.Workbook = excelApp.Workbooks.Add()
            Dim worksheet As Excel.Worksheet = workbook.Sheets(1)
            worksheet.Name = "TUPV_DTR Report"

            For col = 0 To dt.Columns.Count - 1
                worksheet.Cells(1, col + 1).Value = dt.Columns(col).ColumnName
            Next

            For row = 0 To dt.Rows.Count - 1
                Dim dateValue As DateTime
                Dim timeInValue As DateTime
                Dim timeOutValue As DateTime

                For col = 0 To dt.Columns.Count - 1
                    Dim columnName = dt.Columns(col).ColumnName
                    Dim value = dt.Rows(row)(col)

                    Select Case columnName
                        Case "Date"
                            If DateTime.TryParse(value.ToString(), dateValue) Then
                                worksheet.Cells(row + 2, col + 1).Value = dateValue.ToString("MM/dd/yyyy")
                            End If
                        Case "TimeIn", "TimeOut"
                            If DateTime.TryParse(value.ToString(), timeInValue) Then
                                worksheet.Cells(row + 2, col + 1).Value = Convert.ToDateTime(value).ToString("hh:mm:ss tt")
                            End If
                        Case "TotalHours"
                            worksheet.Cells(row + 2, col + 1).Value = Math.Round(Convert.ToDouble(value), 2)
                        Case Else
                            worksheet.Cells(row + 2, col + 1).Value = value.ToString()
                    End Select
                Next
            Next


            worksheet.Columns.AutoFit()

            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Excel Files|*.xlsx"
            saveDialog.Title = "Save DTR Report"
            saveDialog.FileName = "DTR_Report.xlsx"

            If saveDialog.ShowDialog() = DialogResult.OK Then
                workbook.SaveAs(saveDialog.FileName)
                MessageBox.Show("Report exported successfully to Excel.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            workbook.Close(False)
            excelApp.Quit()
            ReleaseObject(worksheet)
            ReleaseObject(workbook)
            ReleaseObject(excelApp)

        Catch ex As Exception
            MessageBox.Show("Export failed: " & ex.Message)
        End Try
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub RefreshBtn_Click(sender As Object, e As EventArgs) Handles RefreshBtn.Click
        Try

            LoadDTRTable()
            MessageBox.Show("Table refreshed successfully.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Failed to refresh table: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        Dim color1 As Color = Color.DarkRed
        Dim color2 As Color = Color.FloralWhite


        Dim brush As New LinearGradientBrush(Me.ClientRectangle, color1, color2, LinearGradientMode.BackwardDiagonal)

        e.Graphics.FillRectangle(brush, Me.ClientRectangle)
    End Sub
End Class
