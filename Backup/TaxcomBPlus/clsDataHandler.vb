Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Win32

Public Class clsDataHandler
    Dim strConn As String = ""
    Dim strServer As String
    Dim strUsername As String
    Dim strPassword As String

    Public Sub New(ByVal strAppType As String)

        If Not strAppType = "" Then
            InitialiseConnection(strAppType)
        End If
    End Sub

    Private Sub InitialiseConnection(ByVal strAppType As String)

        Dim strDatabaseName As String = ""
        Dim strContents As String = ""
        Dim objReader As StreamReader

        Try
            Me.GetServerInfo()
            If strAppType = "B" Or strAppType = "BE" Or strAppType = "M" Then
                strContents = Application.StartupPath & "..\TaxcomB.ini" 'after make exe run this
            ElseIf strAppType = "P" Or strAppType = "CP30" Then
                strContents = Application.StartupPath & "..\TaxcomP.ini" 'after make exe run this
            End If
            objReader = New StreamReader(strContents)
            strDatabaseName = objReader.ReadToEnd()
            objReader.Close()
            'strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strContents
            strConn = "Server=" & strServer & ";Database=" & strDatabaseName & ";User Id=" & strUsername & ";Password=" & strPassword & ";MultipleActiveResultSets=True;"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Function GetServerInfo() As Dictionary(Of String, String)
        Dim dic As New Dictionary(Of String, String)
        strServer = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\TAXOFFICE\", "value1", "")
        strUsername = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\TAXOFFICE\", "value2", "")
        strPassword = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\TAXOFFICE\", "value3", "")
        Return dic
    End Function

    Public Function GetDataReader(ByVal strQuery As String) As SqlDataReader
        Dim dr As SqlDataReader
        Dim cmd As New SqlCommand
        With cmd
            ' Create a Connection object
            .Connection = New SqlConnection(strConn)
            .Connection.Open()
            .CommandText = strQuery
            dr = .ExecuteReader(CommandBehavior.CloseConnection)
        End With
        'If Not dr.HasRows Or dr.RecordsAffected <= 0 Then
        '    dr = Nothing
        'End If
        Return dr
    End Function

    Public Function GetDataReader1(ByVal strQuery As String) As SqlDataReader
        Dim dr As SqlDataReader
        Dim cmd As New SqlCommand
        With cmd
            ' Create a Connection object
            InitialiseConnection("B")
            .Connection = New SqlConnection(strConn)
            .Connection.Open()
            .CommandText = strQuery
            dr = .ExecuteReader(CommandBehavior.CloseConnection)
        End With
        Return dr
    End Function

    Public Function GetData(ByVal strQuery As String) As DataSet

        Dim ds As New DataSet
        Dim dataConnection As New SqlConnection
        dataConnection.ConnectionString = strConn
        Try
            Dim cmd As New SqlCommand(strQuery, dataConnection)
            'If prmOleDb IsNot Nothing Then
            'For Each prmOle As sqlparameter In prmOleDb
            'If prmOle IsNot Nothing Then cmd.Parameters.Add(prmOle)
            '    Next
            'End If
            Dim da As New SqlDataAdapter(cmd)
            da.Fill(ds)

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            If dataConnection.State = ConnectionState.Open Then dataConnection.Close()
        End Try

        Return ds
    End Function

    Public Function GetData(ByVal strQuery As String, ByVal ParamArray prmOleDb As IDataParameter()) As DataSet

        Dim ds As New DataSet
        Dim dataConnection As New SqlConnection
        dataConnection.ConnectionString = strConn
        Try
            Dim cmd As New SqlCommand(strQuery, dataConnection)

            If prmOleDb IsNot Nothing Then
                For Each prmOle As SqlParameter In prmOleDb
                    If prmOle IsNot Nothing Then cmd.Parameters.Add(prmOle)
                Next
            End If
            Dim da As New SqlDataAdapter(cmd)
            da.Fill(ds)

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            If dataConnection.State = ConnectionState.Open Then dataConnection.Close()
        End Try

        Return ds
    End Function

    Public Function Execute(ByVal strSQL As String) As Integer
        Dim objConn As New SqlConnection(strConn)
        Dim cmd As SqlCommand
        Dim intAffectedRow As Integer

        Try
            cmd = New SqlCommand
            cmd.Connection = objConn
            objConn.Open()
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        Finally
            objConn.Close()
        End Try

        Return intAffectedRow
    End Function

    Public ReadOnly Property sqlConnection()
        Get
            Return strConn
        End Get
    End Property
End Class
