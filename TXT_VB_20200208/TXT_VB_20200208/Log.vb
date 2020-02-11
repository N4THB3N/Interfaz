Imports System.IO
Public Class Log
    Public Shared Sub Append_To_Log(ByVal p_LogContent As String)
        Dim v_Path As String = AppDomain.CurrentDomain.BaseDirectory & "\Logs\" & Now.Date.ToString("yyyyMM") & "\"

        Dim v_FileName As String = "DHL5_IFX_GBI_Log_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"

        ' If directory does not exists (first day of month), then create it.
        ' MGamarrro 25-Mar-2019
        '
        If Not IO.Directory.Exists(v_Path) Then

            IO.Directory.CreateDirectory(v_Path)
        End If

        Dim v_FileStream As FileStream = File.Open(v_Path & v_FileName, FileMode.Append)
        Dim v_StreamWriter As New StreamWriter(v_FileStream)

        v_StreamWriter.WriteLine(String.Format("{1}  {0} ", p_LogContent, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")))
        v_StreamWriter.Flush()
        v_StreamWriter.Close()
        v_FileStream.Close()
    End Sub
End Class
