VERSION 5.00
Begin VB.Form FControlArrays 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Array Demo"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FControlArrays.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optType 
      Caption         =   "Aditional &Expense"
      Height          =   255
      Index           =   5
      Left            =   300
      TabIndex        =   8
      Top             =   2220
      Width           =   2055
   End
   Begin VB.OptionButton optType 
      Caption         =   "&Additional Capital"
      Height          =   255
      Index           =   4
      Left            =   300
      TabIndex        =   7
      Top             =   1860
      Width           =   2055
   End
   Begin VB.OptionButton optType 
      Caption         =   "&Turnover"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   6
      Top             =   1500
      Width           =   2055
   End
   Begin VB.OptionButton optType 
      Caption         =   "Pri&ce"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   5
      Top             =   1140
      Width           =   2055
   End
   Begin VB.OptionButton optType 
      Caption         =   "&Demand"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   780
      Width           =   2055
   End
   Begin VB.ListBox lstValue 
      Height          =   1815
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   1515
   End
   Begin VB.OptionButton optType 
      Caption         =   "&Population"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   420
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1740
      TabIndex        =   0
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Caption         =   "&Select a Value:"
      Height          =   195
      Left            =   2820
      TabIndex        =   3
      Top             =   360
      Width           =   1395
   End
End
Attribute VB_Name = "FControlArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

' ************************************************************
' Form Property Procedures Follow
' ************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This property procedure translates the index number
'           of the option button selected by the user into a
'           human-readable description.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Property Get OptionSelected() As String
    Dim lIndex As Long
    For lIndex = optType.LBound To optType.UBound
        If optType(lIndex).Value Then
            OptionSelected = Replace(optType(lIndex).Caption, "&", "")
            Exit For
        End If
    Next lIndex
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This property procedure returns the value selected
'           by the user from the lstValue listbox.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Public Property Get ListSelection() As Double
    ListSelection = CDbl(lstValue.Text)
End Property


' ************************************************************
' Form Event Procedures Follow
' ************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This single event procedure is fired in response to
'           a click on any option button in the control array.
'
' Arguments:    Index           This argument is passed to the
'                               event procedure by the system.
'                               It holds the index number of the
'                               option button that was clicked.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Private Sub optType_Click(Index As Integer)
    Dim vItem As Variant
    Dim vaList As Variant
    lstValue.Clear
    Select Case Index
        Case 0  ' Population
            vaList = Array(500, 1000, 100000, 100000)
        Case 1  ' Demand
            vaList = Array(50, 100, 1000, 10000)
        Case 2  ' Price
            vaList = Array(9.99, 19.99, 29.99, 39.99)
        Case 3  ' Turnover
            vaList = Array(0.01, 0.015, 0.02, 0.025, 0.03)
        Case 4  ' Additional Capital
            vaList = Array(1000, 2000, 3000, 4000)
        Case 5  ' Additional Expense
            vaList = Array(500, 1000, 1500, 2000)
    End Select
    For Each vItem In vaList
        lstValue.AddItem vItem
    Next vItem
    lstValue.ListIndex = -1
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This event procedure verifies that the appropriate
'           selections have been made by the user before hiding
'           the form and returning control to the calling procedure.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Private Sub cmdOK_Click()
    Dim bOptionSelected As Boolean
    Dim lIndex As Long
    ' Do not allow the user to continue unless an option button
    ' has been selected and a list item has been selected
    For lIndex = optType.LBound To optType.UBound
        If optType(lIndex).Value Then
            bOptionSelected = True
            Exit For
        End If
    Next lIndex
    If Not bOptionSelected Or lstValue.ListIndex = -1 Then
        MsgBox "You must select an option and a list item."
    Else
        Me.Hide
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This event procedure reroutes all clicks to the form
'           X-close button through the cmdOK_Click event procedure.
'
' Arguments:    Cancel          Set this argument to True to cancel
'                               unloading the form, which otherwise
'                               is the default behavior when the
'                               X-close button is clicked.
'               UnloadMode      This argument is passed to the event
'                               procedure by the system. It tells us
'                               by what manner the form is being
'                               unloaded. The only one we want to
'                               trap is the X-close button, which
'                               corresponds to an UnloadMode value
'                               of vbFormControlMenu.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Private Sub Form_QueryUnload(Cancel As Integer, _
                                    UnloadMode As Integer)
    ' Route the x-close button through the
    ' cmdOK_Click event procedure.
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        cmdOK_Click
    End If
End Sub

