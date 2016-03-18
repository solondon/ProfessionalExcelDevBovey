'
' Description:  This class contains all the required procedures to export to
'               Excel.
'
' Authors:      Dennis Wallentin, www.excelkb.com
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Imported namespaces
'
'To read the Windows Registry subkey.
Imports Microsoft.Win32
'To use regular expressions.
Imports System.Text.RegularExpressions
'To catch COM exceptions and to release COM objects.
Imports System.Runtime.InteropServices
'To check that Excel templates exist or not.
Imports System.IO
'To get the path to the solution.
Imports System.Reflection
'Namespace alias for Excel.
Imports Excel = Microsoft.Office.Interop.Excel
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Friend Class CExportExcel

#Region "Export to Excel"

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: This function creates a copy of a template based on
    '           which Excel template that has been selected. 
    '           It populates the Data worksheet with selected data.           
    '
    ' Arguments:    dtTable     The in-memory datatable we get all data from.
    '               sClient     The selected client's name.
    '               sProject    The selected project's name.
    '               sStartDate  The selected start date for the report.
    '               sEndDate    The selected end date for the report.
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/15/08      Dennis Wallentin    Ch25    Initial version
    ' 09/16/08      Dennis Wallentin    Ch25    Updated variables 
    ' 09/16/08      Dennis Wallentin    Ch25    Moved function bFile_Exist
    '                                           to the class CCommonMethods
    ' 09/16/08      Dennis Wallentin    Ch25    Moved function Get_Path
    '                                           to the class CCommonMethods

    Friend Function bExport_Excel(ByVal dtTable As DataTable, _
                                  ByVal sClient As String, _
                                  ByVal sProject As String, _
                                  ByVal sStartDate As String, _
                                  ByVal sEndDate As String) As Boolean

        'Declare and instantiate a new instance of the class.
        Dim cCMethods As New CCommonMethods()
        Dim cConnect As New Connect()

        'Constants for the function's exception messages.
        Const sERROR_MESSAGE As String = "An unexpected error has occured "
        Const sERROR_MESSAGE_EXCEL As String = " in Excel."
        Const sERROR_MESSAGE_NET As String = " in the .NET application."

        'All variables for working with Excel COM objects.
        Dim xlCalcMode As Excel.XlCalculation = Nothing
        Dim xlPtCache As Excel.PivotCache = Nothing
        Dim xlWkbTarget As Excel.Workbook = Nothing
        Dim xlWksData As Excel.Worksheet = Nothing
        Dim xlWksPivot As Excel.Worksheet = Nothing
        Dim xlRngProjectData As Excel.Range = Nothing
        Dim xlRngFields As Excel.Range = Nothing
        Dim xlRngData As Excel.Range = Nothing
        Dim xlRngList As Excel.Range = Nothing
        Dim xlRngHours As Excel.Range = Nothing
        Dim xlRngRevenue As Excel.Range = Nothing
        Dim xlRngSortActivities As Excel.Range = Nothing
        Dim xlRngSortLastName As Excel.Range = Nothing
        Dim xlRngAutoFit As Excel.Range = Nothing

        'Variable for working with the datatable.
        Dim dtTableColumn As DataColumn = Nothing

        'When exporting to Excel the first 4 columns in the 
        'datatable are not of interest and since all index in .NET
        'are zero based we use the value 5 (4+1).
        Dim iNumberOfColumns As Integer = dtTable.Columns.Count - 5

        'Get the number of rows from the datatable.
        Dim iNumberOfRows As Integer = dtTable.Rows.Count - 1

        'Counters to iterate through the datatable's columns and rows.
        Dim iRowsCounter As Integer = Nothing
        Dim iColumnsCounter As Integer = Nothing

        'Array to hold the retrieved data from the datatable. Since
        'datatables includes a various of datatype we use here
        'the object datatype.
        Dim obDataArr(iNumberOfRows, iNumberOfColumns) As Object

        'An array that holds the project specific values.
        Dim sProjectDataArr() As String = {sClient, sProject, sStartDate + "-" + sEndDate}

        'We need to separate the PETRAS Report Summary.xlt from the other templates
        'as it does not contain a second worksheet with a Pivot table. Therefore we
        'use a boolean flag.
        Dim bSummaryReport As Boolean = False

        'Variable to return the function's outcome.
        Dim bExport As Boolean = Nothing

        Try

            'To prevent flickering when open and customize the report.
            swXLApp.ScreenUpdating = False

            'Create and open a copy of the selected template
            Dim sWorkBook As String = cCMethods.Get_Path + swTemplateListArr(swshSelectedTemplate).ToString()
            xlWkbTarget = swXLApp.Workbooks.Open(sWorkBook)

            'Save the present setting for Excel's calculation mode and temporarily turn it off.
            With swXLApp
                xlCalcMode = .Calculation
                .Calculation = Excel.XlCalculation.xlCalculationManual
            End With

            'We must explicit cast the object reference to the type Excel Worksheet.
            xlWksData = CType(xlWkbTarget.Worksheets(Index:=1), Excel.Worksheet)

            'If the Summary report template is in use we set the flag to true.
            If swshSelectedTemplate = 3 Then
                bSummaryReport = True
            End If

            'If not the selected template is the Summary report then
            'we need to work with additional Excel objects.
            If bSummaryReport = False Then

                'The second worksheet which contains the pivot table.
                xlWksPivot = CType(xlWkbTarget.Worksheets(Index:=2), Excel.Worksheet)

                'The second worksheet in the templates includes a Pivot table and
                'we need access to its pivot cache. 
                xlPtCache = xlWkbTarget.PivotCaches(Index:=1)

                'The range object requires also to be casted to an Excel range.
                xlRngAutoFit = CType(xlWksPivot.Columns("D:D"), Excel.Range)

            End If

            'Range to add project specific data.
            xlRngProjectData = xlWksData.Range("C3:C5")

            'Add the project specific data.
            xlRngProjectData.Value = swXLApp.WorksheetFunction.Transpose(sProjectDataArr)

            'Populate the array of data from the Datatable.
            For iRowsCounter = 0 To iNumberOfRows
                For iColumnsCounter = 0 To iNumberOfColumns
                    'The first 4 columns hold data which is irrelevant in this
                    'context which we need to consider here by adding 4 to the columns's
                    'counter.
                    obDataArr(iRowsCounter, iColumnsCounter) = dtTable.Rows(iRowsCounter)(4 + iColumnsCounter)
                Next
            Next


            With xlWksData

                'The fields's range, data's range, hours's range and revenue's range are all
                'depended on which template that has been selected.
                Select Case swshSelectedTemplate

                    Case xltTemplate.xlActivities

                        'Range to add the data to.
                        xlRngData = .Range(.Cells(RowIndex:=10, ColumnIndex:=4), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=iNumberOfColumns + 4))

                        'Range for the fields.
                        xlRngFields = .Range("D9:F9")

                        'Range which holds the project hours.
                        xlRngHours = .Range(.Cells(RowIndex:=10, ColumnIndex:=5), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=5))

                        'Range which holds the project revenues values.
                        xlRngRevenue = .Range(.Cells(RowIndex:=10, ColumnIndex:=6), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=6))

                        'Range which holds the cell to base the sorting.
                        xlRngSortActivities = .Range(Cell1:="B9")

                    Case xltTemplate.xltActivitiesConsultants

                        'Range to add the data to.
                        xlRngData = .Range(.Cells(RowIndex:=10, ColumnIndex:=2), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=iNumberOfColumns + 2))

                        'Range for the fields.
                        xlRngFields = .Range("B9:F9")

                        'Range which holds the project hours.
                        xlRngHours = .Range(.Cells(RowIndex:=10, ColumnIndex:=5), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=5))

                        'Range which holds the project revenues values.
                        xlRngRevenue = .Range(.Cells(RowIndex:=10, ColumnIndex:=6), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=6))

                        'Ranges which hold the cells to base the sorting on.
                        xlRngSortActivities = .Range(Cell1:="B9")
                        xlRngSortLastName = .Range(Cell1:="D9")

                    Case xltTemplate.xltConsultants

                        'Range to add the data to.
                        xlRngData = .Range(.Cells(RowIndex:=10, ColumnIndex:=3), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=iNumberOfColumns + 3))

                        'Range for the fields.
                        xlRngFields = .Range("C9:F9")

                        'Range which holds the project hours.
                        xlRngHours = .Range(.Cells(RowIndex:=10, ColumnIndex:=5), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=5))

                        'Range which holds the project revenues values.
                        xlRngRevenue = .Range(.Cells(RowIndex:=10, ColumnIndex:=6), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=6))

                        'Range which holds the cell to base the sorting.
                        xlRngSortLastName = .Range(Cell1:="D9")

                    Case xltTemplate.xltSummary

                        'Range to add the data to.
                        xlRngData = .Range(.Cells(RowIndex:=10, ColumnIndex:=2), _
                                   .Cells(RowIndex:=iNumberOfRows + 10, ColumnIndex:=iNumberOfColumns + 2))

                        'Range for the fields.
                        xlRngFields = .Range("B9:C9")

                End Select

            End With

            'Populate the data range with data.
            xlRngData.Value = obDataArr

            'Concatenate the two ranges into one range.
            xlRngList = swXLApp.Union(xlRngFields, xlRngData)

            'Sort the list based on which template is in use.
            Select Case swshSelectedTemplate

                Case xltTemplate.xlActivities
                    'Sort on activity name.
                    xlRngList.Sort(Key1:=xlRngSortActivities, _
                                   Order1:=Excel.XlSortOrder.xlAscending, _
                                   Header:=Excel.XlYesNoGuess.xlYes, _
                                   Orientation:=Excel.XlSortOrientation.xlSortRows)

                Case xltTemplate.xltActivitiesConsultants
                    'Sort first on activity name and then on lastname.
                    xlRngList.Sort(Key1:=xlRngSortActivities, _
                                   Order1:=Excel.XlSortOrder.xlAscending, _
                                   Key2:=xlRngSortLastName, _
                                   Order2:=Excel.XlSortOrder.xlAscending, _
                                   Header:=Excel.XlYesNoGuess.xlYes, _
                                   Orientation:=Excel.XlSortOrientation.xlSortColumns)

                Case xltTemplate.xltConsultants
                    'Sort by lastname.
                    xlRngList.Sort(Key1:=xlRngSortLastName, _
                                    Order1:=Excel.XlSortOrder.xlAscending, _
                                    Header:=Excel.XlYesNoGuess.xlYes, _
                                    Orientation:=Excel.XlSortOrientation.xlSortColumns)

            End Select

            'Apply a built-in listformat to the list.
            xlRngList.AutoFormat(Format:=Excel.XlRangeAutoFormat.xlRangeAutoFormatList3, _
                                 Number:=True, Font:=True, Alignment:=True, Border:=True, _
                                 Pattern:=True, Width:=True)

            With xlWksData
                'Autosize the range area we use.
                .UsedRange.Columns.AutoFit()
                'Give the worksheet a project specific name.
                .Name = cCMethods.sCreate_Name( _
                        sClientName:=sClient, sProjectName:=sProject) + " " + "Data"
            End With

            'Restore the calculation mode.
            swXLApp.Calculation = xlCalcMode

            'If not the selected template is the Summary report then
            'we need to work with additional Excel objects.
            If bSummaryReport = False Then

                'Update all the range names we use so they cover the actual
                'ranges in the Data worksheet.
                With xlWkbTarget
                    .Names.Item("rnList").RefersTo = xlRngList
                    .Names.Item("rnHours").RefersTo = xlRngHours
                    .Names.Item("rnRevenue").RefersTo = xlRngRevenue
                End With

                'Give the worksheet a project specific name.
                xlWksPivot.Name = cCMethods.sCreate_Name( _
                                  sClientName:=sClient, sProjectName:=sProject) + " " + "Pivot Table"

                'Update the Pivot Cache.
                xlPtCache.Refresh()

                'Size the column D in the Pivot Table sheet.
                xlRngAutoFit.AutoFit()

            End If

            'Things worked out as expected so we set the boolean value to true. 
            bExport = True

        Catch COMExc As COMException

            'All exceptions in COM Servers generate HRESULT messages. In most cases
            'this message is not human understandable and therefore we need to use  
            'a customized message here as well.
            MessageBox.Show(text:= _
                            sERROR_MESSAGE & _
                            sERROR_MESSAGE_EXCEL, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

            'Things didn't worked out as we expected so we set the boolean value
            'to false.
            bExport = False


        Catch Generalexc As Exception

            'Show customized message.
            MessageBox.Show(text:=sERROR_MESSAGE & sERROR_MESSAGE_NET, _
                            caption:=swsCaption, _
                            buttons:=MessageBoxButtons.OK, _
                            icon:=MessageBoxIcon.Stop)

            'Things didn't worked out as we expected so we set the boolean value
            'to false.
            bExport = False

        Finally

            'Restore the status.
            swXLApp.ScreenUpdating = True

           'Prepare the object for GC.
            dtTableColumn = Nothing

            'Release all resources consumed by the variable from the
            'memory.
            dtTable.Dispose()
            dtTable = Nothing

            cCMethods = Nothing
            cConnect = Nothing

            'Calling the Garbish Collector (GC)is a resource consuming process
            'but when working with COM objects it's a necessary process.
            'To make sure that all indirectly Excel COM objects will be released 
            'we call the GC twice.
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

            'Releae all Excel COM objects.
            Release_All_ExcelCOMObjects(xlRngAutoFit)
            Release_All_ExcelCOMObjects(xlRngSortLastName)
            Release_All_ExcelCOMObjects(xlRngSortActivities)
            Release_All_ExcelCOMObjects(xlRngRevenue)
            Release_All_ExcelCOMObjects(xlRngHours)
            Release_All_ExcelCOMObjects(xlRngList)
            Release_All_ExcelCOMObjects(xlRngData)
            Release_All_ExcelCOMObjects(xlRngFields)
            Release_All_ExcelCOMObjects(xlRngProjectData)
            Release_All_ExcelCOMObjects(xlWksPivot)
            Release_All_ExcelCOMObjects(xlWksData)
            Release_All_ExcelCOMObjects(xlWkbTarget)
            Release_All_ExcelCOMObjects(xlPtCache)
            Release_All_ExcelCOMObjects(xlCalcMode)

        End Try

        'Inform the calling procedure about the outcome.
        Return bExport


    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Comments: Release all used Excel objects from memory.
    '           
    '               
    ' Date          Developers          Chap    Action
    ' --------------------------------------------------------------
    ' 09/16/08      Dennis Wallentin    Ch25    Initial version

    Private Sub Release_All_ExcelCOMObjects(ByVal oxlObject As Object)

        Try
            'Release the object and set it to nothing.
            Marshal.FinalReleaseComObject(oxlObject)
            oxlObject = Nothing
        Catch ex As Exception
            oxlObject = Nothing
        End Try

    End Sub

#End Region


End Class
