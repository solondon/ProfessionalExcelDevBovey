Public Class CWorkbooks

    Friend Function bAdd_Workbook() As Boolean

        'Accessing the Excel Application object from a Class module.
        Globals.ThisAddIn.Application. _
        Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)

    End Function



End Class
