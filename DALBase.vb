'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Imports System.Configuration
Imports System.Data.OleDb


Public Class DALBase
    Implements IDisposable


#Region " Class Level Variables "
    'Class level variables that are available to classes that instantiate me
    Public SQL As String

    Public Connection As OleDbConnection
    Public Command As OleDbCommand
    Public DataAdapter As OleDbDataAdapter
    Public DataReader As OleDbDataReader

#End Region

#Region " Constructor and Destructor "
    Public Sub New()
        'Get the connection settings from the application configuration file
        Dim objAppSettings As Specialized.NameValueCollection
        objAppSettings = ConfigurationSettings.AppSettings()
        'Build the SQL connection string and initialize the Connection object
        Connection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=D:\Documents and Settings\megatron\My Documents\db1.mdb;")
        objAppSettings = Nothing
    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose
        If Not DataReader Is Nothing Then
            DataReader.Close()
            DataReader = Nothing
        End If
        If Not DataAdapter Is Nothing Then
            DataAdapter.Dispose()
            DataAdapter = Nothing
        End If
        If Not Command Is Nothing Then
            Command.Dispose()
            Command = Nothing
        End If
        If Not Connection Is Nothing Then
            Connection.Close()
            Connection.Dispose()
            Connection = Nothing
        End If
    End Sub
#End Region

#Region " Public Procedures & Functions "

    Public Sub OpenConnection()
        Try
            Connection.Open()
        Catch OleDbExceptionErr As OleDbException
            Throw New System.Exception(OleDbExceptionErr.Message, _
                OleDbExceptionErr.InnerException)
        Catch InvalidOperationExceptionErr As InvalidOperationException
            Throw New System.Exception(InvalidOperationExceptionErr.Message, _
                InvalidOperationExceptionErr.InnerException)
        End Try
    End Sub

    Public Sub CloseConnection()
        Connection.Close()
    End Sub

    Public Sub InitializeCommand()
        If Command Is Nothing Then
            Try
                Command = New OleDbCommand(SQL, Connection)
                'See if this is a stored procedure
                If Not SQL.ToUpper.StartsWith("SELECT ") _
                    And Not SQL.ToUpper.StartsWith("INSERT ") _
                    And Not SQL.ToUpper.StartsWith("UPDATE ") _
                    And Not SQL.ToUpper.StartsWith("DELETE ") Then
                    Command.CommandType = CommandType.StoredProcedure
                End If
            Catch OleDbExceptionErr As OleDbException
                Throw New System.Exception(OleDbExceptionErr.Message, _
                    OleDbExceptionErr.InnerException)
            End Try
        End If
    End Sub

    Public Sub AddParameter(ByVal Name As String, ByVal Type As OleDbType, _
        ByVal Size As Integer, ByVal Value As Object)
        Try
            Command.Parameters.Add(Name, Type, Size).Value = Value
        Catch OleDbExceptionErr As OleDbException
            Throw New System.Exception(OleDbExceptionErr.Message, _
                OleDbExceptionErr.InnerException)
        End Try
    End Sub

    Public Sub InitializeDataAdapter()
        Try
            DataAdapter = New OleDbDataAdapter
            DataAdapter.SelectCommand = Command
        Catch OleDbExceptionErr As OleDbException
            Throw New System.Exception(OleDbExceptionErr.Message, _
            OleDbExceptionErr.InnerException)
        End Try
    End Sub

    Public Sub FillDataSet(ByRef oDataSet As DataSet, ByVal TableName As String)
        Try
            InitializeCommand()
            InitializeDataAdapter()
            DataAdapter.Fill(oDataSet, TableName)
        Catch OleDbExceptionErr As OleDbException
            Throw New System.Exception(OleDbExceptionErr.Message, _
                OleDbExceptionErr.InnerException)
        Finally
            Command.Dispose()
            Command = Nothing
            DataAdapter.Dispose()
            DataAdapter = Nothing
        End Try
    End Sub

    Public Function ExecuteStoredProcedure() As Integer
        Try
            OpenConnection()
            ExecuteStoredProcedure = Command.ExecuteNonQuery()
        Catch ExceptionErr As Exception
            Throw New System.Exception(ExceptionErr.Message, _
                ExceptionErr.InnerException)
        Finally
            CloseConnection()
        End Try
    End Function
#End Region

End Class
