'To work with ADO.NET
Imports System.Data
'The DataProvider for SQL Server.
Imports System.Data.SqlClient

'By declaring the module as Friend all its functions are
'available for other methods in the solution.
Friend Module MData

    Function Retrieve_Data_With_DataReader() As ArrayList

        'SQL query in use.
        Const sSqlQuery As String = _
            "SELECT CompanyName AS Company " & _
            "FROM Customers " & _
            "ORDER BY CompanyName;"

        'Connection string in use.
        Const sConnection As String = _
            "Data Source=PED\SQLEXPRESS;" & _
            "Initial Catalog=Northwind;" & _
            "Integrated Security=True"

        'Declare and initialize the connection.
        Dim sqlCon As New SqlConnection(connectionString:= _
                                        sConnection)
        'Declare and initialize the command.
        Dim sqlCmd As New SqlCommand(cmdText:=sSqlQuery, _
                                     connection:=sqlCon)
        'Define the command type.
        sqlCmd.CommandType = CommandType.Text

        'Explicit open the connection.
        sqlCon.Open()

        'Populate the DataReader with data and 
        'explicit close the connection.
        Dim sqlDataReader As SqlDataReader = _
        sqlCmd.ExecuteReader(behavior:= _
                             CommandBehavior.CloseConnection)

        'Variable for keeping track of number of rows in the 
        'DataReader.
        Dim iRecordCounter As Integer = Nothing

        'Get the number of columns in the DataReader.
        Dim iColumnsCount As Integer = sqlDataReader.FieldCount

        'Declare and instantiate the ArrayList.
        Dim DataArrLst As New ArrayList

        'Check to see that it consists of, at least, has one 
        'record. 
        If sqlDataReader.HasRows Then

            'Iterate through the collection of records.
            While sqlDataReader.Read

                For iRecordCounter = 0 To iColumnsCount - 1

                    'Add data to the ArrayList's variable.
                    DataArrLst.Add(sqlDataReader.Item _
                                  (iRecordCounter).ToString)

                Next iRecordCounter

            End While
        End If

        'Cleaning up by disposing objects, close and 
        'release variables.
        sqlCmd.Dispose()
        sqlCmd = Nothing

        sqlDataReader.Close()
        sqlDataReader = Nothing

        sqlCon.Close()
        sqlCon.Dispose()
        sqlCon = Nothing

        'Send the list to the calling method.
        Return DataArrLst

    End Function

    Function Retrieve_Data_With_DataSet() As DataTable

        'SQL query in use.
        Const sSqlQuery As String = _
            "SELECT CompanyName AS Company " & _
            "FROM Customers " & _
            "ORDER BY CompanyName;"

        'Connection string in use.
        Const sConnection As String = _
            "Data Source=PED\SQLEXPRESS;" & _
            "Initial Catalog=Northwind;" & _
            "Integrated Security=True"

        'Declare the connection variable.
        Dim SqlCon As SqlConnection = Nothing

        'Declare the DataAdapter variable.
        Dim SqlAdp As SqlDataAdapter = Nothing

        'Declare and initialize a new empty DataSet.
        Dim SqlDataSet As New DataSet

        Try
            'Initialize the connection.
            SqlCon = New SqlConnection(connectionString:= _
                                       sConnection)
            'Initialize the DataAdapter.
            SqlAdp = New SqlDataAdapter(selectCommandText:= _
                                         sSqlQuery, _
                                         selectConnection:= _
                                         SqlCon)

            'Fill the DataSet.
            SqlAdp.Fill(dataSet:=SqlDataSet, srcTable:="PED")

            'Return the datatable.
            Return SqlDataSet.Tables(0)

        Catch Sqlex As SqlException
            'Exception handling for the communication with
            'the SQL Server Database.

            'Tell it to the calling method.
            Return Nothing

        Finally

            'Releases all resources the variable has consumed from
            'the memory.
            SqlDataSet.Dispose()

            'Release the reference the variable holds and 
            'prepare it to be collected by the Garbage Collector
            '(GC) when it comes around.
            SqlDataSet = Nothing

            SqlCon.Dispose()
            SqlCon = Nothing

            SqlAdp.Dispose()
            SqlAdp = Nothing

        End Try

    End Function


End Module
