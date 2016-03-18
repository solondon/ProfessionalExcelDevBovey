VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11340
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   15375
   _ExtentX        =   27120
   _ExtentY        =   20003
   _Version        =   393216
   Description     =   "Adds/removes calls to the PerfMonitor class."
   DisplayName     =   "PerfMon: VB6 IDE Addin"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Class to allow connection from the Office/VB6 IDE
'
'   This class handles the connection/disconnection and sets up and removes the menu items.
'   It initialises instances of the CMenuHandler class, which handles the menus being clicked
'
'   Version Date        Author          Comment
'   0.0.1   09-02-2004  Stephen Bullen  Initial Version
'
Option Explicit
Option Compare Binary

'A reference to the Office/VB6 IDE
Dim moVBE As VBIDE.VBE

'A collection of our menu bars, to be deleted when we tidy up
Dim moBars As Collection

'A collection of our CMenuHandler class, to handle the menu click events
Dim moEvents As Collection

'Called when the addin is loaded by the VBE
'Sets up the PerfMon menus
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    'save the vb instance
    Set moVBE = Application

    On Error Resume Next
    Set moBars = New Collection
    Set moEvents = New Collection

    AddPopupToBar moVBE.CommandBars("Add-Ins")
    AddPopupToBar moVBE.CommandBars("Code Window")

End Sub


'Called when the addin is unloaded by the VBE
'Removes the PerfMon menus
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    Dim oBar As CommandBarControl

    On Error Resume Next

    'Delete the command bar entry
    For Each oBar In moBars
        oBar.Delete
    Next

End Sub


'Adds the PerfMon popup to the given commandbar.
'Called twice, for the Tools > Addins menu and the Code Window right-click popup
Private Sub AddPopupToBar(oBar As CommandBar)

    Dim cbPopUp As CommandBarPopup

    On Error Resume Next

    'Does the PerfMon popup already exist?
    Set cbPopUp = oBar.Controls("PerfMon")

    If cbPopUp Is Nothing Then

        'No, so create it
        Set cbPopUp = oBar.Controls.Add(msoControlPopup, temporary:=True)
        cbPopUp.Caption = "PerfMon"

        'And add it to our collection for tidy up
        moBars.Add cbPopUp
    End If

    'Add our buttons to the popup menu
    AddControlToPopup cbPopUp.CommandBar, "Add PerfMon Calls", "AddCode"
    AddControlToPopup cbPopUp.CommandBar, "Remove PerfMon Calls", "RemoveCode"
    AddControlToPopup cbPopUp.CommandBar, "About PerfMon", "About"

End Sub


'Adds a control to a commandbar and sets up the event handlers
Private Sub AddControlToPopup(oPopup As CommandBar, sCaption As String, sParam As String)

    Dim cbControl As Office.CommandBarButton
    Dim oHandler As CMenuHandler

    'Add the control to the command bar
    Set cbControl = oPopup.Controls.Add(1)

    'Set its properties
    cbControl.Caption = sCaption
    cbControl.Parameter = sParam

    'Each control gets a unique tag
    cbControl.Tag = "PerfMon" & Rnd()

    'Create a new instance of our menu event handler ...
    Set oHandler = New CMenuHandler

    '... initialise it ...
    oHandler.Initialise moVBE, cbControl

    '... and add it to our collection to keep it alive
    moEvents.Add oHandler

End Sub

