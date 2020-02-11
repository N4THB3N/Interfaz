Imports System.Data.SqlClient
Imports System.Configuration
Public Class DB_SQLServer
    Private vCon As SqlConnection
    Private vCom As SqlCommand
    Private vReader As IDataReader

    Private Function OpenDb() As Boolean
        vCon = New SqlConnection
        Try
            If vCon.State = ConnectionState.Open Then
                vCon.Close()
            End If
            vCon.ConnectionString = ConfigurationManager.ConnectionStrings("TXT_ConStr").ToString
            vCon.Open()
            OpenDb = True
        Catch ex As Exception
            OpenDb = False
        End Try
    End Function

    Public Function ReturnDataSet(ByVal pQuery As String) As DataSet
        vCom = New SqlCommand
        Dim vDataSet As New DataSet
        Dim vDataAdapter As New SqlDataAdapter

        If OpenDb() Then
            vCom.Connection = vCon
            vCom.CommandType = CommandType.Text
            vCom.CommandText = pQuery
            vCom.CommandTimeout = 3600
            vDataAdapter = New SqlDataAdapter(pQuery, vCon)
            vDataAdapter.Fill(vDataSet)
            CloseDb()
            Return vDataSet
        End If
        Return Nothing
    End Function

    Public Function CloseDb() As Boolean
        Try
            ' Checks if there is an open data reader to the database
            If Not vReader Is Nothing Then
                If Not vReader.IsClosed Then
                    vReader.Close()
                End If
            End If


            ' Closes the connection
            If vCon.State = ConnectionState.Open Then
                vCon.Close()
            End If

            CloseDb = True
        Catch ex As Exception
            CloseDb = False
            ' Error engine code
        End Try
    End Function


End Class
