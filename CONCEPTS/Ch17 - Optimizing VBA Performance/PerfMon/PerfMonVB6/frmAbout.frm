VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About PerfMon"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   120
      Picture         =   "frmAbout.frx":0152
      ScaleHeight     =   660
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
   Begin VB.Label lblExampleSub 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sub LongRoutine()"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.BMSLtd.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      MouseIcon       =   "frmAbout.frx":1E74
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblMore 
      Caption         =   "For more utilities and examples (primarily for Microsoft Excel), visit:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   5175
   End
   Begin VB.Label lblInstruct 
      Caption         =   $"frmAbout.frx":1FC6
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label lblCopyright 
      Caption         =   "© 2003-2009 by Business Model Services Ltd"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      Caption         =   "PerfMon v1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   An About form, giving an example routine for using the PerfMon dll
'
'   Version Date        Author          Comment
'   0.0.1   09-02-2004  Stephen Bullen  Initial Version

Option Explicit
Option Compare Binary

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()

    Me.lblExampleSub.Caption = "Sub LongRoutine()" & vbLf & vbLf & _
                               "    PerfMonStartMonitoring" & vbLf & _
                               "    PerfMonProcStart ""Project.Module.LongRoutine""" & vbLf & vbLf & _
                               "    'Do Stuff" & vbLf & vbLf & _
                               "    PerfMonProcEnd ""Project.Module.LongRoutine""" & vbLf & _
                               "    PerfMonStopMonitoring ""C:\LongRoutineTiming.txt""" & vbLf & vbLf & _
                               "End Sub"

End Sub

Private Sub lblURL_Click()

    On Error Resume Next

    ShellExecute 0&, vbNullString, "www.BMSLtd.co.uk", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub btnOK_Click()

    On Error Resume Next

    Unload Me

End Sub


