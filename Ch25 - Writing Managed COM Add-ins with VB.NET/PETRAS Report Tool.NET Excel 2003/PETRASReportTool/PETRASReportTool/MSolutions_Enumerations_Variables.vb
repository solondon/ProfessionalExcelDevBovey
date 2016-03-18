'
' Description:  This module contains all solutions wide variables that 
'               are in use and also enumerations.
'
' Authors:      Dennis Wallentin, www.excelkb.com
'

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Imported namespaces
'Namespace alias for Excel.
Imports Excel = Microsoft.Office.Interop.Excel


Module MSolutions_Enumerations_Variables

    'The Excel Application object variable.
    Friend swXLApp As Excel.Application

    'Read the title of the application into a variable.
    Friend ReadOnly swsCaption As String = My.Application.Info.Title.ToString()

    'A variable that holds the selected Excel template id-number.
    Friend swshSelectedTemplate As Short

    'An array that holds all the Excel templates's names.
    Friend swTemplateListArr() As String = {"PETRAS Report Activities.xlt", _
                                           "PETRAS Report Activities Consultants.xlt", _
                                           "PETRAS Report Consultants.xlt", _
                                           "PETRAS Report Summary.xlt"}

    'Enumeration of Report Templates.
    Friend Enum xltTemplate As Short
        xlActivities = 0
        xltActivitiesConsultants = 1
        xltConsultants = 2
        xltSummary = 3
    End Enum

    'Enumeration of Excel versions.
    Friend Enum xlVersion As Short
        NoVersion = 0
        RightVersion = 1
        WrongVersion = 2
    End Enum

End Module
