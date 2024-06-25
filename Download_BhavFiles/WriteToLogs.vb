Imports System.IO

Public Class WriteToLogs
    Public Sub CaptureLogs(path As String, Texts As String, Optional Types As String = "INFO")

        If (System.IO.Directory.Exists(path) = False) Then
            System.IO.Directory.CreateDirectory(path)
        End If

        If (System.IO.Directory.Exists(path) = False) Then
            System.IO.Directory.CreateDirectory(path)
        End If


        Dim strFileName As String

        If (Types.ToUpper <> "INFO") Then
            strFileName = path & "err.txt"
            'Dim writer As StreamWriter = New StreamWriter(path & "\err.txt", True)
            'writer.WriteLine(Texts)
            'My.Computer.FileSystem.WriteAllText(path & "\err.txt", path & vbCrLf, True)
            'writer.Close()
        Else
            'My.Computer.FileSystem.WriteAllText(path & "\info.txt", path & vbCrLf, True)
            strFileName = path & "info.txt"
        End If

        ' Dim strFile As String = String.Format("C:\ErrorLog_{0}.txt", DateTime.Today.ToString("dd-MMM-yyyy"))
        File.AppendAllText(strFileName, String.Format(DateTime.Now & " " & Texts & Environment.NewLine, DateTime.Now, Environment.NewLine))


    End Sub



    Public Sub CaptureInsertsLogs(path As String, Texts As String)

        If (System.IO.Directory.Exists(path) = False) Then
            System.IO.Directory.CreateDirectory(path)
        End If

        If (System.IO.Directory.Exists(path) = False) Then
            System.IO.Directory.CreateDirectory(path)
        End If

        Dim strFileName As String
        strFileName = path & "Inserts.txt"
        ' Dim strFile As String = String.Format("C:\ErrorLog_{0}.txt", DateTime.Today.ToString("dd-MMM-yyyy"))
        File.AppendAllText(strFileName, String.Format(DateTime.Now & " " & Texts & Environment.NewLine, DateTime.Now, Environment.NewLine))
    End Sub


    Public Sub CaptureUpdateLogs(path As String, Texts As String)

        If (System.IO.Directory.Exists(path) = False) Then
            System.IO.Directory.CreateDirectory(path)
        End If

        If (System.IO.Directory.Exists(path) = False) Then
            System.IO.Directory.CreateDirectory(path)
        End If

        Dim strFileName As String
        strFileName = path & "Update.txt"
        ' Dim strFile As String = String.Format("C:\ErrorLog_{0}.txt", DateTime.Today.ToString("dd-MMM-yyyy"))
        File.AppendAllText(strFileName, String.Format(DateTime.Now & " " & Texts & Environment.NewLine, DateTime.Now, Environment.NewLine))


    End Sub

End Class
