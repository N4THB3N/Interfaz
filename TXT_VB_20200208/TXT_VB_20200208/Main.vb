Imports System.Configuration
Module Main
    Private o_TXT_VB As New TXT_VB_Class
    Private v_Receipt_ID As String = ""
    Private v_utc_date As DateTime
    Dim v_Running_Mode = ConfigurationManager.AppSettings("Running_Mode")
    Private v_Station_Cd As String = String.Empty
    Private v_GMT As String = String.Empty
    Private v_Data_Access As New Data_Access_Class
    Private o_VB_TXT As New TXT_VB_Class
    Private v_Ctry_Cd As String = String.Empty
    Private v_Resend As String = ConfigurationManager.AppSettings("Resend_YN")


    Private v_To_ID As String = ConfigurationManager.AppSettings("To_Receipt_ID")
    Private v_From_ID As String
    Private v_Days_To_Get As String
    Private v_To_Dt As String
    Private v_From_Dt As String

    Sub Main()
        Try
            If v_Running_Mode.ToLower.Equals("manual") Then
                v_From_Dt = ConfigurationManager.AppSettings("From_Dt")
                v_To_Dt = ConfigurationManager.AppSettings("To_Dt")
            ElseIf v_Running_Mode.ToLower.Equals("automatic") Then
                v_To_Dt = Date.Now.AddDays(v_Days_To_Get).ToString("yyyy-MM-dd")
                v_From_Dt = Date.Now.AddDays(v_Days_To_Get).ToString("yyyy-MM-dd")
            End If

            v_utc_date = DateTime.UtcNow.ToString("yyyy-MM-dd")
            v_utc_date.AddHours(-4)


        Catch ex As Exception

        End Try
    End Sub
End Module