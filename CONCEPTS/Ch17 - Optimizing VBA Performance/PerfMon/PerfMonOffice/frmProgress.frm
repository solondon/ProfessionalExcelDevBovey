VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PerfMon Progress"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblProc 
      Caption         =   "Label2"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4440
   End
   Begin VB.Label lblAction 
      Caption         =   "Adding PerfMon calls to procedure:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4410
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   Simple form to show the progress of the routine adding/removing the PerfMon calls
'
'   Version Date        Author          Comment
'   0.0.1   09-02-2004  Stephen Bullen  Initial Version
'
Option Explicit
Option Compare Binary

Public Property Let Action(sNew As String)
    Me.lblAction.Caption = sNew
End Property

Public Property Let Procedure(sProc As String)
    Me.lblProc.Caption = sProc
End Property

