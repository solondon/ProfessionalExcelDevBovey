VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResourceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

' ************************************************************
' Class Enumeration Declarations Follow
' ************************************************************
Public Enum resIcon
    resIconArrows = 101
    resIconArrowsMask = 102
End Enum


' ************************************************************
' Class Property Procedures Follow
' ************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This property procedure provides access to the icons
'           stored in the Icons.res resource file.
'
' Arguments:    Icon            One of the resIcon enumeration
'                               members that identify the specific
'                               icon you want to retrieve.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Property Get Icon(ByVal uName As resIcon) As StdPicture
    Set Icon = LoadResPicture(uName, vbResBitmap)
End Property
