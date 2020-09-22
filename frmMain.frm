VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NT/2000 Event Log Demo"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSuccess 
      Caption         =   "Success Event"
      Height          =   390
      Left            =   75
      TabIndex        =   2
      Top             =   1155
      Width           =   2625
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Information Event"
      Height          =   390
      Left            =   75
      TabIndex        =   1
      Top             =   705
      Width           =   2625
   End
   Begin VB.CommandButton cmdCritical 
      Caption         =   "Critical Event"
      Height          =   390
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   2610
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Logger As NTEventLog

Private Sub cmdCritical_Click()
    Logger.WriteToLog &H1, EVENTLOG_ERROR_TYPE, catSystemEvent
End Sub

Private Sub cmdInfo_Click()
    Logger.WriteToLog &H2, EVENTLOG_INFORMATION_TYPE, catSystemEvent
End Sub


Private Sub cmdSuccess_Click()
    Logger.WriteToLog &H4, EVENTLOG_SUCCESS, catSystemEvent
End Sub


Private Sub Form_Load()
    Set Logger = New NTEventLog
    Logger.UseLog = evtApplication
    Logger.AppName = "TestApp"
    Logger.MessageDLL = "%SystemRoot%\system32\testapp.dll"
    Logger.Connect
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Logger = Nothing
End Sub


