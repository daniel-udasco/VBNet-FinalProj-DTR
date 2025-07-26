Imports System.Data.OleDb

Public Class DatabaseHelper
    Private connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\CHUWI\Desktop\VBNet_FinalProject\Database\DTR.accdb"
    Private conn As New OleDbConnection(connStr)

    Public Function LoadNameDictionary() As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)
        Try
            conn.Open()
            Dim cmd As New OleDbCommand("SELECT [TUPVID], [Name] FROM tblCode", conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            While reader.Read()
                dict(reader("TUPVID").ToString()) = reader("Name").ToString()
            End While
            reader.Close()
        Catch ex As Exception
            Throw New Exception("LoadNameDictionary error: " & ex.Message)
        Finally
            conn.Close()
        End Try
        Return dict
    End Function

    Public Sub InsertTimeIn(name As String, [date] As String, timeIn As String)
        Try
            conn.Open()
            Dim cmd As New OleDbCommand("INSERT INTO tblDTR ([Name], [Date], [TimeIn]) VALUES (?, ?, ?)", conn)
            cmd.Parameters.AddWithValue("?", name)
            cmd.Parameters.AddWithValue("?", [date])
            cmd.Parameters.AddWithValue("?", timeIn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("InsertTimeIn error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub UpdateTimeOut(name As String, [date] As String, timeOut As String)
        Try
            conn.Open()
            Dim cmd As New OleDbCommand("UPDATE tblDTR SET [TimeOut]=? WHERE [Name]=? AND [Date]=? AND [TimeOut] IS NULL", conn)
            cmd.Parameters.AddWithValue("?", timeOut)
            cmd.Parameters.AddWithValue("?", name)
            cmd.Parameters.AddWithValue("?", [date])
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("UpdateTimeOut error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub UpdateTotalHours(name As String, [date] As String, totalHours As Double)
        Try
            conn.Open()
            Dim cmd As New OleDbCommand("UPDATE tblDTR SET [TotalHours]=? WHERE [Name]=? AND [Date]=?", conn)
            cmd.Parameters.AddWithValue("?", totalHours)
            cmd.Parameters.AddWithValue("?", name)
            cmd.Parameters.AddWithValue("?", [date])
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("UpdateTotalHours error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub
End Class
