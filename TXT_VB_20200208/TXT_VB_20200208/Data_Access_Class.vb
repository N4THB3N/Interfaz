Imports System.Configuration
Imports System.Data.SqlClient
Imports TXT_VB_20200208.Log

Public Class Data_Access_Class
    Private Const v_SimpleQuote As String = "'"
    Private vDB As New DB_SQLServer
    Private v_DoubleQuotes As String = "'"
    Private vReader As SqlDataReader
    Private vResult As String = String.Empty
    Private vCommand As New SqlCommand

    Public Function Get_Receipts_Select_For_GBI(ByVal p_Station_ID As Integer, p_From_Dt As String, ByVal p_To_Dt As String, ByVal p_From_Num_Docto As Integer, ByVal p_To_Num_Docto As Integer, ByVal p_Resend As String) As DataSet
        Dim v_DataSet As DataSet

        Dim v_Query As String = "EXEC usp_Receipts_Select4GBI '" & p_Station_ID & "' , '" & p_From_Dt & "' ,'" & p_To_Dt & "' , " & p_From_Num_Docto & " , " & p_To_Num_Docto & " , '" & p_Resend & "'"
        Append_To_Log(v_Query)

        v_DataSet = vDB.ReturnDataSet(v_Query)

        vDB.CloseDb()


    End Function

End Class
