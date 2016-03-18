Attribute VB_Name = "MEntryPoints"
'
' Description:  This module contains all entry points into the application.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This is the startup procedure for the front loader.
'           It calls the functions that determine whether Word
'           and Outlook are available on this computer. If both
'           applications are available it starts our Excel app.
'           If either Word or Outlook is not available it
'           displays a warning message to the user and exits.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Sub Main()

    Dim bHasWord As Boolean
    Dim bHasOutlook As Boolean
    Dim xlApp As Excel.Application
    Dim wkbPetras As Excel.Workbook
    Dim frmWarning As FWarning
    
    ' Verify that we can automate both Word and Outlook on this computer.
    bHasWord = bWordAvailable()
    bHasOutlook = bOutlookAvailable()
    
    If bHasWord And bHasOutlook Then
        ' If we successfully automated both Word and Outlook,
        ' load our Excel app and turn it over to the user.
        Set xlApp = New Excel.Application
        xlApp.Visible = True
        xlApp.UserControl = True
        Set wkbPetras = xlApp.Workbooks.Open(App.Path & _
                                        "\PetrasAddin.xla")
        wkbPetras.RunAutoMacros xlAutoOpen
        Set wkbPetras = Nothing
        Set xlApp = Nothing
    Else
        ' If we failed to get a reference to either Word or
        ' Outlook, display a warning message to the user and
        ' exit without taking further action.
        Set frmWarning = New FWarning
        frmWarning.Show
        Set frmWarning = Nothing
    End If
    
End Sub


