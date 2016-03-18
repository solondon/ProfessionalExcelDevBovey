VERSION 5.00
Begin VB.Form FRoman 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Command1"
      Height          =   435
      Left            =   420
      TabIndex        =   1
      Top             =   660
      Width           =   1635
   End
   Begin VB.TextBox txtConvert 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label lblResult 
      Caption         =   "Label1"
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   2355
   End
End
Attribute VB_Name = "FRoman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit

' **************************************************************
' Form Event Procedures Follow
' **************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Runs automatically whenever the form loads. This
'           event procedure is responsible for initialiing the
'           form and all of its controls.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Private Sub Form_Load()
    ' Form properties
    Me.BorderStyle = vbFixedDouble
    Me.Caption = "Convert to Roman Numerals"
    ' CommandButton properties
    cmdConvert.Caption = "Convert To Roman"
    ' TextBox properties
    txtConvert.Alignment = vbRightJustify
    txtConvert.MaxLength = 4
    txtConvert.Text = ""
    ' Label properties
    lblResult.Alignment = vbCenter
    lblResult.BackColor = &HE0E0E0
    lblResult.BorderStyle = vbFixedSingle
    lblResult.Caption = ""
    lblResult.Font.Name = "Courier"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This event procedure is used to limit the characters
'           that can be entered into the txtConvert textbox to
'           numerals and the backspace key.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Private Sub txtConvert_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            ' Backspace and numerals 0 through 9
            ' these are all OK. Take no action.
        Case Else
            ' No other characters are permited.
            KeyAscii = 0
    End Select
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This event procedure converts the whole number in
'           the txtConvert textbox into Roman numerals using the
'           Excel ROMAN worksheet function. The result is stored
'           in the lblResult label control.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 04/30/08      Rob Bovey       Ch23    Initial version
'
Private Sub cmdConvert_Click()

    Dim bError As Boolean
    Dim xlApp As Excel.Application
    Dim lConvert As Long
    Dim sErrMsg As String
    
    ' Coerce the text box value into a long.
    ' Val is required in case it is empty.
    lConvert = CLng(Val(txtConvert.Text))
    
    ' Don't do anything unless txtConvert contains
    ' a number greater than zero.
    If lConvert > 0 Then
    
        ' The maximum number that can be converted
        ' to Roman numeral is 3999.
        If lConvert <= 3999 Then
        
            Set xlApp = New Excel.Application
            lblResult.Caption = _
                xlApp.WorksheetFunction.Roman(lConvert)
            xlApp.Quit
            Set xlApp = Nothing
    
        Else
            sErrMsg = "The maximum number that can be converted"
            sErrMsg = sErrMsg & " to a Roman numeral is 3999."
            bError = True
        End If
        
    Else
        sErrMsg = "The minimum number that can be converted"
        sErrMsg = sErrMsg & " to a Roman numeral is 1."
        bError = True
    End If
    
    If bError Then
        MsgBox sErrMsg, vbCritical, "Error"
        txtConvert.SetFocus
        txtConvert.SelStart = 0
        txtConvert.SelLength = Len(txtConvert.Text)
    End If
    
End Sub


