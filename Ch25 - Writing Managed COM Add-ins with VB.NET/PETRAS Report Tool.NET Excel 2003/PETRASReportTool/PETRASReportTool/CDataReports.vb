'
' Description:  This module contains all the required functions to
'               populate the controls in the main Windows Form 
'               with data.
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Imported namespaces
'To use the SQL Server client class library.
Imports System.Data.SqlClient

Public Class CDataReports

#Region "Module-wide variables."

    'A customized error message.
    Private msMESSAGESQLERROR As String = "Please report the issue to " + vbNewLine + _
                                         "the Database Administrator."

#End Region

#Region "Properties"

    'Property for the connection string.
    Private ReadOnly Property msConnection() As String
        Get
            msConnection = My.Settings.SQLConnection.ToString()
            Return msConnection
        End Get
    End Property

#End Region

#Region "Connection."

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function instantiates a new connection to the
    '           database.
    '
    ' Arguments:    None
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/13/08      Dennis Wallentin    Ch25    Initial version

    Friend Function sqlCreate_Connection() As SqlConnection

        Return New SqlConnection(connectionString:=Me.msConnection)

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This is the procedure that tries to access the 
    '           SQL Server database PETRA. If true then a connection
    '           pooling is created for the application.
    '
    ' Arguments:    None
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/13/08      Dennis Wallentin    Ch25    Initial version

    Friend Function bConnect_Database() As Boolean

        'Customized error message.
        Const sMESSAGECONNECTIONERROR As String = "Unable to connect to the database." + vbNewLine + _
                                                  "If you are unable to access the " + vbNewLine + _
                                                  "database for a longer time please " + vbNewLine + _
                                                  "report it to the Database Administrator."

        'The maximum number of tries we should do to connect to the database
        'before we give up.
        Const iMAXACCESSTIMES As Integer = 3

        'The variable for keeping track on how many tries we have done.
        Dim iCountAccessTries As Integer = 0

        'The connection variable.
        Dim SqlCon As SqlConnection = Nothing

        'Variable that holds the value to be returned. 
        Dim bStatusConnection As Boolean = Nothing

        Do While bStatusConnection = Nothing

            Try
                'Instantiate a new connection.
                SqlCon = Me.sqlCreate_Connection()

                'If we cannot open the connection it will generate an exception.
                'If we establish the connection then this connection will be added
                'to the database connection pooling.
                With SqlCon
                    .Open()
                    .Close()
                End With

                'Set the flag to OK.
                bStatusConnection = True

            Catch SqlExc As SqlException

                'To keep record of number of tries.
                iCountAccessTries += 1

                'If we cannot access the database and has reached the maximum
                'number then it's time to cancel the action.
                If iCountAccessTries = iMAXACCESSTIMES Then
                    MessageBox.Show(text:=sMESSAGECONNECTIONERROR, _
                                    caption:=swsCaption, _
                                    buttons:=MessageBoxButtons.OK, _
                                    icon:=MessageBoxIcon.Stop)

                    'Things didn't worked out as we expected so we set the boolean 
                    'value to false.
                    bStatusConnection = False
                    Exit Do
                End If

            End Try
        Loop

        If bStatusConnection = False Then
            'Release all resources consumed by the variable from the
            'memory.
            SqlCon.Dispose()

            'Prepare the object for GC.
            SqlCon = Nothing

        End If

        'Inform the calling procedure about the outcome.
        Return bStatusConnection

    End Function

#End Region

#Region "Retrieve the data from the database."

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function retrieves the clients's names from
    '           the database who actually have registrated ongoing 
    '           projects.
    '
    ' Arguments:    None
    '
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/13/08      Dennis Wallentin    Ch25    Initial version

    Friend Function dtGet_Clients() As DataTable

        'Select only clients that have running projects.
        Dim sSqlQuestion As String = "SELECT ClientID, ClientName FROM Clients " + _
                                     "WHERE ClientID IN (SELECT DISTINCT ClientID FROM Projects) " + _
                                     "ORDER BY ClientName;"


        'The variable for the SQL adapter in use.
        Dim adpSql As SqlDataAdapter = Nothing

        'Declare and instantiate a new dataset.
        Dim dsData As New DataSet

        Try
            'Instantiate a new SQL adapter. 
            adpSql = New SqlDataAdapter(selectCommandText:=sSqlQuestion, _
                                        selectConnection:=Me.sqlCreate_Connection)
            'Fill the dataset with data.
            adpSql.Fill(dataSet:=dsData, srcTable:="PETRA")

            'Name the retrieved datatable.
            dsData.Tables(0).TableName = "Clients"

            'Return the datatable.
            Return dsData.Tables("Clients")


        Catch SqlExc As SqlException

            'Show a customized message for any exceptions related 
            'to the SQL adapter. 
            MessageBox.Show(text:=msMESSAGESQLERROR, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

            'Things didn't worked out as we expected so we set the datatable
            'to nothing and return it.
            Return Nothing

        Finally

            'Releases all resources the variable has consumed from
            'the memory.
            dsData.Dispose()
            'Prepare the object for GC.
            dsData = Nothing

            adpSql.Dispose()
            adpSql = Nothing

        End Try

    End Function

    Friend Function dtGetProjects_Client(ByVal iClient As Integer) As DataTable

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Comments: This function retrieves the list of projects's names from
        '           the database based on the selected client's ID.
        '
        ' Arguments:    lnClient    The Client's ID in the database.
        '
        ' Date          Developers          Chap    Action
        ' --------------------------------------------------------------
        ' 09/13/08      Dennis Wallentin    Ch25    Initial version


        Dim sSqlQuestion As String = "SELECT ProjectID, ClientID, ProjectName From Projects " + _
                                     "WHERE ClientID =" + iClient.ToString() + _
                                     " ORDER BY ProjectName;"

        'The variable for the SQL adapter in use.
        Dim adpSql As SqlDataAdapter = Nothing

        'Declare and instantiate a new dataset.
        Dim dsData As New DataSet


        Try
            'Instantiate a new SQL adapter 
            adpSql = New SqlDataAdapter(selectCommandText:=sSqlQuestion, _
                                        selectConnection:=Me.sqlCreate_Connection)

            'Fill the dataset with data.
            adpSql.Fill(dataSet:=dsData, srcTable:="PETRA")

            'Name the retrieved datatable.
            dsData.Tables(0).TableName = "Projects"

            'Return the datatable.
            Return dsData.Tables("Projects")

        Catch sqlExc As SqlException

            'Show a customized message for any exceptions related 
            'to the SQL adapter.
            MessageBox.Show(text:=msMESSAGESQLERROR, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

            'Things didn't worked out as we expected so we set the datatable
            'to nothing and return it.
            Return Nothing

        Finally

            'Release all resources consumed by the variable from the
            'memory.
            dsData.Dispose()
            'Prepare the object for GC.
            dsData = Nothing

            adpSql.Dispose()
            adpSql = Nothing

        End Try

    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function aquires a datatable with data from the
    '           database.
    '
    ' Arguments:    sSQLQuery      The SQL statement to be executed.
    '               dtStartDate    The selected start date for the report.
    '               dtEndDate      The selected end date for the report.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/13/08      Dennis Wallentin    Ch25    Initial version

    Private Function dtGeneral_Report(ByVal sSQLQuery As String, _
                                      ByVal dtStartDate As Date, _
                                      ByVal dtEndDate As Date) As DataTable


        'In order to build up the parameterized SQL statement we need a SQL 
        'command variable.
        Dim cmdSql As SqlCommand = Nothing

        'The variable for the SQL adapter in use.
        Dim adpSql As SqlDataAdapter = Nothing

        'Declare and instantiate a new dataset.
        Dim dsData As New DataSet


        Try
            'Instantiate a new SQL command.
            cmdSql = New SqlCommand

            'Create the parameterized SQL statement and create a connection
            'to the database.
            With cmdSql
                .CommandText = sSQLQuery
                .Parameters.Add(parameterName:="@dtStartdate", _
                                sqlDbType:=SqlDbType.DateTime).Value = dtStartDate.Date
                .Parameters.Add(parameterName:="@dtEndDate", _
                                sqlDbType:=SqlDbType.DateTime).Value = dtEndDate.Date
                .Connection = Me.sqlCreate_Connection()
            End With

            'Instantiate a new SQL adapter based on the SQL command and the created
            'parameterized SQL statement.
            adpSql = New SqlDataAdapter(selectCommandText:=sSQLQuery, _
                                        selectConnection:=cmdSql.Connection)

            'Fill the dataset.
            With adpSql
                .SelectCommand = cmdSql
                .Fill(dataSet:=dsData, srcTable:="PETRA")
            End With

            'Check to see the number of rows (i e records) retrieved.
            If dsData.Tables(0).Rows.Count >= 1 Then

                'Create a name for the datatable.
                dsData.Tables(0).TableName = "General"

                'Return the datatable.
                Return dsData.Tables("General")

            Else
                'If no records exist then return nothing.
                Return Nothing

            End If


        Catch Sqlexc As SqlException

            'Show a customized message for any exceptions related 
            'to the SQL adapter.
            MessageBox.Show(text:=msMESSAGESQLERROR + _
                            Sqlexc.Message.ToString(), _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

            'Things didn't worked out as we expected so we set the datatable
            'to nothing and return it.
            Return Nothing

        Finally

            'Release all resources consumed by the variable from the
            'memory.
            dsData.Dispose()
            'Prepare the object for GC.
            dsData = Nothing

            adpSql.Dispose()
            adpSql = Nothing

            cmdSql.Dispose()
            cmdSql = Nothing

        End Try

    End Function

#End Region

#Region "All SQL statements."

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function returns all activties related to one
    '           selected project in a datatable.
    '
    ' Arguments:    lnProjID    The selected project ID in the database.
    '               dtStart     The selected start date for the report.
    '               dtEnd       The selected end date for the report.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/13/08      Dennis Wallentin    Ch25    Initial version

    Friend Function dtActivities(ByVal lnProjID As Long, ByVal dtStart As Date, _
                                 ByVal dtEnd As Date) As DataTable


        'SQL statement where only activities are selected.
        Dim sSqlQuestion As String = _
               "SELECT MAX(Clients.ClientID) AS ClientID, " + _
               "MAX(Clients.ClientName) AS ClientName, " + _
               "MAX(Projects.ProjectID) AS ProjectID, " + _
               "MAX(Projects.ProjectName) AS ProjectName, " + _
               "Activities.ActivityName, SUM(BillableHours.Hours) AS Hours, " + _
               "SUM(Activities.Rate*BillableHours.Hours) AS Revenue FROM Activities, " + _
               "BillableHours, Clients, Projects WHERE " + _
               "Clients.ClientID=Projects.ClientID AND " + _
               "Projects.ProjectID = BillableHours.ProjectID AND " + _
               "BillableHours.ActivityID = Activities.ActivityID AND " + _
               "BillableHours.ProjectID = " + lnProjID.ToString()


        sSqlQuestion += " AND BillableHours.DateWorked >= @dtStartDate " + _
                        "AND BillableHours.DateWorked <= @dtEndDate "

        sSqlQuestion += "GROUP BY Activities.ActivityName;"


        Return Me.dtGeneral_Report(sSQLQuery:=sSqlQuestion, dtStartDate:=dtStart, dtEndDate:=dtEnd)


    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function returns all consultants related to one
    '           selected project in a datatable.
    '
    ' Arguments:    lnProjID    The selected project ID in the database.
    '               dtStart     The selected start date for the report.
    '               dtEnd       The selected end date for the report.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/13/08      Dennis Wallentin    Ch25    Initial version


    Friend Function dtConsultants(ByVal lnProjId As Long, ByVal dtStart As Date, _
                                  ByVal dtEnd As Date) As DataTable

        'SQL statement where only consultants are selected.
        Dim sSqlQuestion As String = _
               "SELECT MAX(Clients.ClientID) AS ClientID," + _
               "MAX(Clients.ClientName) AS ClientName, " + _
               "MAX(Projects.ProjectID) AS ProjectID, " + _
               "MAX(Projects.ProjectName) AS ProjectName ," + _
               "MAX(Consultants.FirstName) AS FirstName, " + _
               "MAX(Consultants.LastName)AS LastName, " + _
               "SUM(BillableHours.Hours) AS Hours, " + _
               "SUM(Activities.Rate*BillableHours.Hours) AS Revenue FROM Activities, " + _
               "BillableHours, Consultants, Projects, Clients WHERE " + _
               "Clients.ClientID = Projects.ClientID AND " + _
               "Projects.ProjectID = BillableHours.ProjectID AND " + _
               "BillableHours.ActivityID = Activities.ActivityID AND " + _
               "BillableHours.ProjectID = " + lnProjId.ToString()

        sSqlQuestion += " AND Consultants.ConsultantID=BillableHours.ConsultantID"

        sSqlQuestion += " AND BillableHours.DateWorked >= @dtStartDate " + _
                        "AND BillableHours.DateWorked <= @dtEndDate "

        sSqlQuestion += "GROUP BY Consultants.ConsultantID;"


        Return Me.dtGeneral_Report(sSQLQuery:=sSqlQuestion, dtStartDate:=dtStart, dtEndDate:=dtEnd)

    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function returns all activities and consultants 
    '           related to one selected project in a datatable.
    '
    ' Arguments:    lnProjID    The selected project ID in the database.
    '               dtStart     The selected start date for the report.
    '               dtEnd       The selected end date for the report.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/13/08      Dennis Wallentin    Ch25    Initial version

    Friend Function dtActivities_Consultants(ByVal lnProjId As Long, ByVal dtStart As Date, _
                                             ByVal dtEnd As Date) As DataTable


        'SQL statement where both activities and consultants are selected.
        Dim sSqlQuestion As String = _
               "SELECT MAX(Clients.ClientID) AS ClientID, " + _
               "MAX(Clients.ClientName) AS ClientName, " + _
               "MAX(Projects.ProjectID) AS ProjectID, " + _
               "MAX(Projects.ProjectName) AS ProjectName, " + _
               "MAX(Activities.ActivityName) AS Activity, " + _
               "MAX(Consultants.FirstName)AS FirstName, " + _
               "MAX(Consultants.LastName) AS LastName, " + _
               "SUM(Billablehours.Hours) AS Hours, " + _
               "SUM(Activities.Rate*BillableHours.Hours) AS Revenue FROM Activities, " + _
               "BillableHours, Consultants, Clients,Projects WHERE " + _
               "Clients.ClientID = Projects.ClientID AND " + _
               "Projects.Projectid = BillableHours.ProjectID AND " + _
               "BillableHours.ActivityID = Activities.ActivityID AND " + _
               "BillableHours.ProjectID = " + lnProjId.ToString()


        sSqlQuestion += " AND Consultants.ConsultantID=BillableHours.ConsultantID"

        sSqlQuestion += " AND BillableHours.DateWorked >= @dtStartDate " + _
                        "AND BillableHours.DateWorked <= @dtEndDate "

        sSqlQuestion += "GROUP BY Activities.ActivityID, Consultants.ConsultantID;"


        Return Me.dtGeneral_Report(sSQLQuery:=sSqlQuestion, dtStartDate:=dtStart, dtEndDate:=dtEnd)

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function returns all hours and revenue for
    '           one selected project in a datatable.
    '
    ' Arguments:    lnProjID    The selected project ID in the database.
    '               dtStart     The selected start date for the report.
    '               dtEnd       The selected end date for the report.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/13/08      Dennis Wallentin    Ch25    Initial version

    Friend Function dtSummary(ByVal lnProjId As Long, ByVal dtStart As Date, _
                              ByVal dtEnd As Date) As DataTable

        'SQL statement where both activities and consultants are selected.
        Dim sSqlQuestion As String = _
               "SELECT MAX(Clients.ClientID) AS ClientID, " + _
               "MAX(Clients.ClientName) AS ClientName, " + _
               "MAX(Projects.ProjectID) AS ProjectID, " + _
               "MAX(Projects.ProjectName) AS ProjectName, " + _
               "SUM(BillableHours.Hours) AS Hours, " + _
               "SUM(Activities.Rate*BillableHours.hours) AS Revenue FROM Activities, " + _
               "BillableHours, Projects, Clients WHERE " + _
               "Clients.ClientID = Projects.ClientID AND " + _
               "Projects.ProjectID = BillableHours.ProjectID AND " + _
               "BillableHours.ActivityID= Activities.ActivityID AND " + _
               "BillableHours.ProjectID = " + lnProjId.ToString()


        sSqlQuestion += " AND BillableHours.DateWorked >= @dtStartDate AND BillableHours.DateWorked <= @dtEndDate"

        sSqlQuestion = sSqlQuestion & " GROUP BY Projects.ProjectID;"

        Return Me.dtGeneral_Report(sSQLQuery:=sSqlQuestion, dtStartDate:=dtStart, dtEndDate:=dtEnd)


    End Function

#End Region



End Class
