Public Class TXT_VB
    Public Shared Sub Append_To_Log(ByVal p_LogContent As String)
        Dim v_Path As String = AppDomain.CurrentDomain.BaseDirectory & "\Logs\" & Now.Date.ToString("yyyyMM") & "\"

        Dim v_FileName As String = "TXT_VB_Log_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"


    End Sub
End Class
