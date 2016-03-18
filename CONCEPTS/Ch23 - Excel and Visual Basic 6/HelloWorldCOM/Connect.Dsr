VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6015
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   7395
   _ExtentX        =   13044
   _ExtentY        =   10610
   _Version        =   393216
   Description     =   "A simple ""Hello World"" COM addin written in VB6"
   DisplayName     =   "Hello World COM Add-in"
   AppName         =   "Microsoft Excel"
   AppVer          =   "Microsoft Excel 11.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel"
   RegExtra        =   "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Excel\Addins\HelloWorld.Connect"
   RegInfoCount    =   3
   RegType0        =   1
   RegKeyName0     =   "FriendlyName"
   RegSData0       =   "llo World COM Add-"
   RegType1        =   1
   RegKeyName1     =   "Description"
   RegSData1       =   "simple ""Hello World"" COM addin written in V"
   RegType2        =   2
   RegKeyName2     =   "LoadBehavior"
   RegDData2       =   3
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Run when the addin is loaded
Private Sub AddinInstance_OnConnection( _
            ByVal Application As Object, _
            ByVal ConnectMode As _
                  AddInDesignerObjects.ext_ConnectMode, _
            ByVal AddInInst As Object, _
            custom() As Variant)

    MsgBox "Hello World"
End Sub

'Run when the addin is unloaded
Private Sub AddinInstance_OnDisconnection( _
            ByVal RemoveMode As _
                  AddInDesignerObjects.ext_DisconnectMode, _
            custom() As Variant)

    MsgBox "Goodbye World"
End Sub


