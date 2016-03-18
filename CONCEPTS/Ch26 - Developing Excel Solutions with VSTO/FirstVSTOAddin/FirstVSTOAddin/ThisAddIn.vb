Imports System.Windows.Forms

Public Class ThisAddIn

    'Variables for loading and unloading XLAs.
    'Private Const m_sXLWBUDGET As String = "Budget.xlsx"
    'Private Const m_sXLABUDGET As String = "C:\Budget\Budget.xlam"
    'Private Const m_sXLABUDGET_DISPLAYNAME As String = "Budget"

    'Variables for loading and unloading COM Add-ins.
    'Private Const m_sXLWBUDGET As String = "Budget.xlsx"
    'Name of the add-in as listed in the COM Add-ins dialog.
    'Private Const m_sXLABUDGET As String = "BudgetReport"
    'Private m_xlCOMAddins As Office.COMAddIns
    'Private m_xlCOMBudget As Office.COMAddIn

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon()
    End Function

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

        'MessageBox.Show("Hello World")

        'Accessing the Excel Application object.
        'Me.Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)

        'Access the Excel Application object in a Class module.
        'Dim CBooks As New CWorkbooks
        'CBooks.bAdd_Workbook()
        'CBooks = Nothing

    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

        MessageBox.Show("Goodbye World")

    End Sub


    Private Sub Application_WorkbookBeforeClose(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose

        'Unloading XLAs.
        'Dim xlAddin As Excel.AddIn = Nothing

        'Try

        'If Wb.Name.ToString = m_sXLWBUDGET Then

        'For Each xlAddin In _
        'Globals.ThisAddIn.Application.AddIns
        'If xlAddin.FullName = m_sXLABUDGET Then
        'xlAddin.Installed = False
        'Exit For
        'End If
        'Next

        'End If

        'Catch ex As Exception

        'MsgBox(ex.Message.ToString())

        'End Try


        'Unloading COM Add-ins.
        'm_xlCOMAddins = Globals.ThisAddIn.Application.COMAddIns
        'm_xlCOMBudget = m_xlCOMAddins.Item(m_sXLABUDGET)

        'Try

        'If Wb.Name.ToString = m_sXLWBUDGET Then

        'If m_xlCOMBudget.Connect Then _
        'm_xlCOMBudget.Connect = False

        'End If

        'Catch ex As Exception

        'MsgBox(ex.Message.ToString())

        'End Try


    End Sub


    Private Sub Application_WorkbookOpen(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook) Handles Application.WorkbookOpen

        'Loading XLAs.
        'Try


        'If Wb.Name.ToString() = m_sXLWBUDGET Then

        'With Globals.ThisAddIn.Application
        '.AddIns.Add(Filename:=m_sXLABUDGET)
        '.AddIns(m_sXLABUDGET_DISPLAYNAME).Installed = True
        'End With

        'End If


        'Catch ex As Exception

        'MsgBox(ex.Message.ToString())

        'End Try


        'Loading COM Add-ins.
        'm_xlCOMAddins = Globals.ThisAddIn.Application.COMAddIns
        'm_xlCOMBudget = m_xlCOMAddins.Item(m_sXLABUDGET)

        'Try


        'If Wb.Name.ToString() = m_sXLWBUDGET Then

        'If m_xlCOMBudget.Connect = False Then _
        'm_xlCOMBudget.Connect = True
        'End If


        'Catch ex As Exception

        'MsgBox(ex.Message.ToString())

        'End Try

    End Sub
End Class
