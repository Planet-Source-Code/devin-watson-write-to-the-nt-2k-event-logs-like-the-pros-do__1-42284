VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NTEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NTEventLog.cls - Allows VB to work with Message DLLs, and
'                 also write directly to the Event Log, rather
'                 than use the VB facility.

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, lpUserSid As Any, ByVal wNumStrings As Long, ByVal dwDataSize As Long, ByVal lpStrings As Long, lpRawData As Any) As Long
Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long

Private MyRegistry As RegistryRoutines  'Used to make sure we have a message DLL.
Private hLog As Long                    'Handle to the event log
Private mApp As String                  'Application name

'Which event log to log to?
Public Enum EvtLog
    evtApplication
    evtSecurity
    evtSystem
End Enum

Private Log As EvtLog
Private mMessageDLL As String

Public Enum EvtLogType
    EVENTLOG_SUCCESS = &H0
    EVENTLOG_ERROR_TYPE = &H1
    EVENTLOG_WARNING_TYPE = &H2
    EVENTLOG_INFORMATION_TYPE = &H4
    EVENTLOG_AUDIT_SUCCESS = &H8
    EVENTLOG_AUDIT_FAILURE = &H10
End Enum

Public Enum EvtCategoryType
    catDevices = 1
    catDisk = 2
    catPrinters = 3
    catServices = 4
    catShell = 5
    catSystemEvent = 6
    catNetwork = 7
End Enum

Public Property Let AppName(strVal As String)
    mApp = Trim(strVal)
    'Insert your own application name here. Keep the subkey path
    'the same.
    If Log = evtApplication Then
        MyRegistry.Subkey = "CurrentControlSet\Services\EventLog\Application\" & mApp
    End If
    
    If Log = evtSecurity Then
        MyRegistry.Subkey = "CurrentControlSet\Services\EventLog\Security\" & mApp
    End If
    
    If Log = evtSystem Then
        MyRegistry.Subkey = "CurrentControlSet\Services\EventLog\System\" & mApp
    End If
End Property


Private Function CurComputerName() As String
    Dim CName As String
    CName = Space(512)
    GetComputerName CName, 512
    CurComputerName = Trim(CName)
End Function

Public Property Get AppName() As String
    AppName = mApp
End Property


Public Sub Connect()
    'Routine to check if the key does or does not
    'exist. I keep this inside of the sub here,
    'because otherwise, you may get corruption.
    'Doing this ensures that everything else
    'down the line is cool.
    If Not MyRegistry.KeyExists Then
        MyRegistry.CreateKey
        MyRegistry.SetRegistryValue "EventMessageFile", mMessageDLL, REG_EXPAND_SZ
        MyRegistry.SetRegistryValue "TypesSupported", 7, REG_DWORD
    End If
    
    Dim TheName As String
    TheName = CurComputerName
    TheName = MyRegistry.TrimNull(TheName)
    'Now, we go ahead and get our handle to the log file.
    hLog = RegisterEventSource(TheName, mApp)

End Sub

Public Property Get MessageDLL() As String
    MessageDLL = mMessageDLL
End Property

Public Property Let MessageDLL(strVal As String)
    mMessageDLL = Trim(strVal)
End Property

Public Property Get UseLog() As EvtLog
    UseLog = Log
End Property

Public Property Let UseLog(LogToUse As EvtLog)
    Log = LogToUse
End Property

Public Sub WriteToLog(MessageID As Long, LogType As EvtLogType, LogCategory As EvtCategoryType)
    ReportEvent hLog, LogType, LogCategory, MessageID, ByVal 0, 0, 0, ByVal 0, ByVal 0
End Sub

Private Sub Class_Initialize()
    Set MyRegistry = New RegistryRoutines
    mApp = ""
    Log = evtApplication
    MyRegistry.hkey = HKEY_LOCAL_MACHINE
    MyRegistry.KeyRoot = "System"
End Sub


Private Sub Class_Terminate()
    Set MyRegistry = Nothing
    DeregisterEventSource hLog
End Sub


