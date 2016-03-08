VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "salir"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSpy 
      Caption         =   "spy"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblWatching 
      Caption         =   "Label1"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Types with Suffix 2 are from Dan Appleman's API book but I think they are wrong
Private Type PRINTER_NOTIFY_OPTIONS2
    Version As Long
    Flags As Long
    Count As Long
    pTypes As Long
End Type

Private Type PRINTER_NOTIFY_OPTIONS_TYPE2
    Type As Integer
    Reserved0 As Integer
    Reserved1 As Long
    Reserved2 As Long
    Count As Long
    pFields As Long
End Type

Private Type PRINTER_NOTIFY_INFO2
    Version As Long
    Flags As Long
    Count As Long
End Type

Private Type PRINTER_NOTIFY_INFO_DATA2
    Type As Integer
    Field As Integer
    Reserved As Long
    Id As Long
    Buf As Long
End Type

' suffix 1 are from MSDN documentation
Private Type PRINTER_NOTIFY_INFO_DATA1
    Type As Integer
    Field As Integer
    Reserved As Long
    Id As Long
    cbBuf As Long  ' part of union also awData[0]
    pBuf As Long   ' part of union
End Type

Private Type PRINTER_NOTIFY_OPTIONS_TYPE1
    Type As Integer
    Reserved0 As Integer
    Reserved1 As Long
    Reserved2 As Long
    Count As Long
    pFields As Long ' to array of Integer
End Type

Private Type PRINTER_NOTIFY_OPTIONS1
    Version As Long
    Flags As Long
    Count As Long
    pTypes As Long  'to array of PRINTER_NOTIFY_OPTIONS_TYPE
End Type

Private Type PRINTER_NOTIFY_INFO1
    Version As Long
    Flags As Long
    Count As Long
    aData As Long ' to array of PRINTER_NOTIFY_INFO_DATA
End Type

' this is also form Dan's book...conflicting?
' but it's the winner
Private Type PRINTER_NOTIFY_INFO3
  Version As Long
  Flags As Long
  Count As Long
  aData(4) As PRINTER_NOTIFY_INFO_DATA1  ' Varies
  ' this works but only because I have specified adata(4) here when
  ' I know I am only asking for 3 pnid's
  ' I need to learn how to use a pointer to a VB array which is what adata
End Type

Private Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As Long
        DesiredAccess As Long
End Type


Private Const PRINTER_CHANGE_ADD_FORM = &H10000
Private Const PRINTER_CHANGE_ADD_JOB = &H100
Private Const PRINTER_CHANGE_ADD_PORT = &H100000
Private Const PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
Private Const PRINTER_CHANGE_ADD_PRINTER = &H1
Private Const PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
Private Const PRINTER_CHANGE_ALL = &H7777FFFF
Private Const PRINTER_CHANGE_CONFIGURE_PORT = &H200000
Private Const PRINTER_CHANGE_DELETE_FORM = &H40000
Private Const PRINTER_CHANGE_DELETE_JOB = &H400
Private Const PRINTER_CHANGE_DELETE_PORT = &H400000
Private Const PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
Private Const PRINTER_CHANGE_DELETE_PRINTER = &H4
Private Const PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
Private Const PRINTER_CHANGE_FAILED_CONNECTION_PRINTER = &H8
Private Const PRINTER_CHANGE_FORM = &H70000
Private Const PRINTER_CHANGE_JOB = &HFF00
Private Const PRINTER_CHANGE_PORT = &H700000
Private Const PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
Private Const PRINTER_CHANGE_PRINTER = &HFF
Private Const PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
Private Const PRINTER_CHANGE_SET_FORM = &H20000
Private Const PRINTER_CHANGE_SET_JOB = &H200
Private Const PRINTER_CHANGE_SET_PRINTER = &H2
Private Const PRINTER_CHANGE_SET_PRINTER_DRIVER = &H20000000
Private Const PRINTER_CHANGE_TIMEOUT = &H80000000
Private Const PRINTER_CHANGE_WRITE_JOB = &H800



Private Const WAIT_FAILED = -1&
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_ABANDONED = &H80&
Private Const WAIT_ABANDONED_0 = &H80&
Private Const WAIT_TIMEOUT = &H102&
Private Const WAIT_IO_COMPLETION = &HC0&
Private Const STILL_ACTIVE = &H103&
Private Const INFINITE = -1&
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Const PRINTER_NOTIFY_FIELD_SERVER_NAME = &H0
Private Const PRINTER_NOTIFY_FIELD_PRINTER_NAME = &H1
Private Const PRINTER_NOTIFY_FIELD_SHARE_NAME = &H2
Private Const PRINTER_NOTIFY_FIELD_PORT_NAME = &H3
Private Const PRINTER_NOTIFY_FIELD_DRIVER_NAME = &H4
Private Const PRINTER_NOTIFY_FIELD_COMMENT = &H5
Private Const PRINTER_NOTIFY_FIELD_LOCATION = &H6
Private Const PRINTER_NOTIFY_FIELD_DEVMODE = &H7
Private Const PRINTER_NOTIFY_FIELD_SEPFILE = &H8
Private Const PRINTER_NOTIFY_FIELD_PRINT_PROCESSOR = &H9
Private Const PRINTER_NOTIFY_FIELD_PARAMETERS = &HA
Private Const PRINTER_NOTIFY_FIELD_DATATYPE = &HB
Private Const PRINTER_NOTIFY_FIELD_SECURITY_DESCRIPTOR = &HC
Private Const PRINTER_NOTIFY_FIELD_ATTRIBUTES = &HD
Private Const PRINTER_NOTIFY_FIELD_PRIORITY = &HE
Private Const PRINTER_NOTIFY_FIELD_DEFAULT_PRIORITY = &HF
Private Const PRINTER_NOTIFY_FIELD_START_TIME = &H10
Private Const PRINTER_NOTIFY_FIELD_UNTIL_TIME = &H11
Private Const PRINTER_NOTIFY_FIELD_STATUS = &H12
Private Const PRINTER_NOTIFY_FIELD_STATUS_STRING = &H13
Private Const PRINTER_NOTIFY_FIELD_CJOBS = &H14
Private Const PRINTER_NOTIFY_FIELD_AVERAGE_PPM = &H15
Private Const PRINTER_NOTIFY_FIELD_TOTAL_PAGES = &H16
Private Const PRINTER_NOTIFY_FIELD_PAGES_PRINTED = &H17
Private Const PRINTER_NOTIFY_FIELD_TOTAL_BYTES = &H18
Private Const PRINTER_NOTIFY_FIELD_BYTES_PRINTED = &H19

Private Const JOB_NOTIFY_FIELD_PRINTER_NAME = &H0
Private Const JOB_NOTIFY_FIELD_MACHINE_NAME = &H1
Private Const JOB_NOTIFY_FIELD_PORT_NAME = &H2
Private Const JOB_NOTIFY_FIELD_USER_NAME = &H3
Private Const JOB_NOTIFY_FIELD_NOTIFY_NAME = &H4
Private Const JOB_NOTIFY_FIELD_DATATYPE = &H5
Private Const JOB_NOTIFY_FIELD_PRINT_PROCESSOR = &H6
Private Const JOB_NOTIFY_FIELD_PARAMETERS = &H7
Private Const JOB_NOTIFY_FIELD_DRIVER_NAME = &H8
Private Const JOB_NOTIFY_FIELD_DEVMODE = &H9
Private Const JOB_NOTIFY_FIELD_STATUS = &HA
Private Const JOB_NOTIFY_FIELD_STATUS_STRING = &HB
Private Const JOB_NOTIFY_FIELD_SECURITY_DESCRIPTOR = &HC
Private Const JOB_NOTIFY_FIELD_DOCUMENT = &HD
Private Const JOB_NOTIFY_FIELD_PRIORITY = &HE
Private Const JOB_NOTIFY_FIELD_POSITION = &HF
Private Const JOB_NOTIFY_FIELD_SUBMITTED = &H10
Private Const JOB_NOTIFY_FIELD_START_TIME = &H11
Private Const JOB_NOTIFY_FIELD_UNTIL_TIME = &H12
Private Const JOB_NOTIFY_FIELD_TIME = &H13
Private Const JOB_NOTIFY_FIELD_TOTAL_PAGES = &H14
Private Const JOB_NOTIFY_FIELD_PAGES_PRINTED = &H15
Private Const JOB_NOTIFY_FIELD_TOTAL_BYTES = &H16
Private Const JOB_NOTIFY_FIELD_BYTES_PRINTED = &H17
Private Const PRINTER_NOTIFY_TYPE = &H0
Private Const JOB_NOTIFY_TYPE = &H1
Private Const INVALID_HANDLE_VALUE = -1
Private Const PRINTER_NOTIFY_OPTIONS_REFRESH = &H1
Private Const PRINTER_NOTIFY_INFO_DISCARDED = &H1

Private Const JOB_STATUS_BLOCKED_DEVQ = &H200
Private Const JOB_STATUS_DELETED = &H100
Private Const JOB_STATUS_DELETING = &H4
Private Const JOB_STATUS_ERROR = &H2
Private Const JOB_STATUS_OFFLINE = &H20
Private Const JOB_STATUS_PAPEROUT = &H40
Private Const JOB_STATUS_PAUSED = &H1
Private Const JOB_STATUS_PRINTED = &H80
Private Const JOB_STATUS_PRINTING = &H10
Private Const JOB_STATUS_RESTART = &H800
Private Const JOB_STATUS_SPOOLING = &H8
Private Const JOB_STATUS_USER_INTERVENTION = &H10000



Private Declare Function FindFirstPrinterChangeNotification& Lib "winspool.drv" (ByVal hPrinter As Long, ByVal fdwFlags As Long, ByVal fdwOptions As Long, ByVal pPrinterNotifyOptions As Long)
Private Declare Function FindNextPrinterChangeNotification& Lib "winspool.drv" (ByVal hChange As Long, pdwChange As Long, ByVal pPrinterNotifyOptions As Long, ppPrinterNotifyInfo As Long)
Private Declare Function FindClosePrinterChangeNotification& Lib "winspool.drv" (ByVal hChange As Long)
Private Declare Function FreePrinterNotifyInfo Lib "winspool.drv" (ByVal addr As Long) As Long
Private Declare Function OpenPrinter& Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS)
Private Declare Function ClosePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
Private Declare Function WaitForSingleObject& Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Dim hNotification As Long
Dim hPrinter As Long
Dim quitFlag As Boolean
Dim watching As Boolean
Dim s1 As String, s2 As String



Private Sub cmdExit_Click()
  quitFlag = True
  If Not watching Then Unload Me
End Sub

Private Sub cmdSpy_Click()
  watching = True
  lblWatching.Caption = watching
  Dim res As Long
  Dim lFlags As Long
  Dim x As Long
  Dim c As Integer
  Dim mbres As Long
  mbres = vbOK
  
  Dim pDefault As PRINTER_DEFAULTS
  With pDefault
    .pDatatype = vbNullString
    .pDevMode = 0
    .DesiredAccess = PRINTER_ACCESS_USE
  End With
  
  If hPrinter <> 0 Then Call ClosePrinter(hPrinter)
  
  res = OpenPrinter(Printer.DeviceName, hPrinter, pDefault)
  
  If res = 0 Then
    Debug.Print "Unable to open printer " & Printer.DeviceName & " " & GetLastError
    Exit Sub
  Else
    Debug.Print "Open printer " & Printer.DeviceName
  End If
  
  Dim pni As PRINTER_NOTIFY_INFO3
  Dim pnid() As PRINTER_NOTIFY_INFO_DATA1
  Dim pnot As PRINTER_NOTIFY_OPTIONS_TYPE1
  Dim pno As PRINTER_NOTIFY_OPTIONS1
  Dim myData(4) As PRINTER_NOTIFY_INFO_DATA1
  Dim myOptions(0) As PRINTER_NOTIFY_OPTIONS_TYPE1
  Dim myInfo(4) As Integer
  Dim ppni As Long, ppno As Long
  
  myInfo(0) = JOB_NOTIFY_FIELD_MACHINE_NAME
  myInfo(1) = JOB_NOTIFY_FIELD_USER_NAME
  myInfo(2) = JOB_NOTIFY_FIELD_PRINTER_NAME
  myInfo(3) = JOB_NOTIFY_FIELD_TOTAL_PAGES
  myInfo(4) = JOB_NOTIFY_FIELD_STATUS
  
  With myOptions(0)
    .Type = JOB_NOTIFY_TYPE
    .Count = 5
    .pFields = VarPtr(myInfo(0))
  End With
  
  With myData(0)
    .Type = JOB_NOTIFY_TYPE
    .Field = JOB_NOTIFY_FIELD_MACHINE_NAME
  End With
  
  With myData(1)
    .Type = JOB_NOTIFY_TYPE
    .Field = JOB_NOTIFY_FIELD_USER_NAME
  End With
  
  With myData(2)
    .Type = JOB_NOTIFY_TYPE
    .Field = JOB_NOTIFY_FIELD_PRINTER_NAME
  End With
  
  With myData(3)
    .Type = JOB_NOTIFY_TYPE
    .Field = JOB_NOTIFY_FIELD_PAGES_PRINTED
  End With
  
  With myData(4)
    .Type = JOB_NOTIFY_TYPE
    .Field = JOB_NOTIFY_FIELD_STATUS_STRING
  End With
  
  
  With pno
    .Count = 1
    .Version = 2
    .pTypes = VarPtr(myOptions(0))
    .Flags = PRINTER_NOTIFY_OPTIONS_REFRESH
  End With
  
  
  ppno = VarPtr(pno)
  
  ' option 1
  ' OK
  'hNotification = FindFirstPrinterChangeNotification(hPrinter, PRINTER_CHANGE_JOB, 0, 0)
  
  'option2 was crap so it's gone -
  
  'option 3
  ' results in dll error 120 (This function is not supported on this system. )
  ' OK when not using network printer
  hNotification = FindFirstPrinterChangeNotification(hPrinter, 0, 0, ppno)
  
  'option 4
  ' results in dll error 120 (This function is not supported on this system. )
  ' OK when not using network printer
  'hNotification = FindFirstPrinterChangeNotification(hPrinter, PRINTER_CHANGE_JOB, 0, ppno)
  
  If hNotification = INVALID_HANDLE_VALUE Then
    Debug.Print "Unable to create notification object " & GetLastError
    watching = False
    lblWatching.Caption = watching
    Exit Sub
  End If


  Do
    ' currently, with option 1 above, I get a signal of -256 every 10 seconds!
    ' which is not documented.  I assume somehing on the network is doing this
    ' perhaps to keep the queue refreshed or something??
    
    ' my printer queue in the office is called "\\AKTHSE1\THSE11_4SI_MP_UX"
    ' when a real job is queued, I don't get a signal
      
    ' stop the press!  it works fine for a printer I set up locally (FILE:)
    ' so it must be to do with the network or the Netware queue.
    ' Continuing with tests on local queue instead.
    
    res = WaitForSingleObject(hNotification, 500)
    If res = WAIT_TIMEOUT Then
      ' do nothing. set quitflag if you want to
    Else
      res = FindNextPrinterChangeNotification(hNotification, lFlags, 0, ppni)
      Debug.Print res, lFlags, GetLastError
      
      If (lFlags And PRINTER_CHANGE_ADD_JOB) = PRINTER_CHANGE_ADD_JOB Then
          Debug.Print Now & " Job was added"
      End If
      
      If (lFlags And PRINTER_CHANGE_WRITE_JOB) = PRINTER_CHANGE_WRITE_JOB Then
          Debug.Print Now & " Job was written"
      End If
      
      If (lFlags And PRINTER_CHANGE_DELETE_JOB) = PRINTER_CHANGE_DELETE_JOB Then
          Debug.Print Now & " Job was deleted"
      End If
      
      If ppni <> 0 Then
        CopyMemory ByVal (VarPtr(pni)), ByVal ppni, Len(pni)
        Debug.Print "pni:", "Flags", "Count", "Version", "aData"
        Debug.Print "pni:", pni.Flags, pni.Count, pni.Version, VarPtr(pni.aData(0))
        If pni.Count > 0 Then
          'ReDim pnid(pni.Count - 1)
          ' tricky stuff. aData is not a pointer but the first
          ' byte of the array of pnid
          ' I might have a problem here because the pBuf seems
          ' to point to a buffer, but I can't get the string out of it
          'CopyMemory pnid(0), addr, Len(pnid(0)) * pni.Count
          
          ' attempt # 2 - maybe the aData points to a safearray?
          ' darn documentation is hard to follow
          pnid = pni.aData 'will this work?
          
          If (pni.Flags And PRINTER_NOTIFY_INFO_DISCARDED) = PRINTER_NOTIFY_INFO_DISCARDED Then
            Debug.Print "Some data was discarded "
            ' to be done: use PRINTER_NOTIFY_OPTIONS_REFRESH to get the discarded data
          End If
          
          Debug.Print "pnid : ", "Type", "Field", "Id", "cbBuf", "pBuf"
          For c = 0 To UBound(pnid)
            Debug.Print "pnid : ", pnid(c).Type, pnid(c).Field, pnid(c).Id, pnid(c).cbBuf, pnid(c).pBuf,
            PrintPrinterNotifyInfoData pnid(c)
          Next
        End If
        res = FreePrinterNotifyInfo(ppni)
      End If
    
    End If
    DoEvents
  Loop While Not quitFlag
  
  watching = False
  lblWatching.Caption = watching
    
  If quitFlag Then Unload Me
    
End Sub

Private Sub PrintPrinterNotifyInfoData(pnid As PRINTER_NOTIFY_INFO_DATA1)
  With pnid
    Select Case .Type
    Case PRINTER_NOTIFY_TYPE
    
    Case JOB_NOTIFY_TYPE
      Select Case .Field
        Case JOB_NOTIFY_FIELD_PRINTER_NAME
          Debug.Print "Printer Name: " & GetStringFromLPSTR(.pBuf, .cbBuf),
          
        Case JOB_NOTIFY_FIELD_MACHINE_NAME
          Debug.Print "Machine Name: " & GetStringFromLPSTR(.pBuf, .cbBuf),
          
        Case JOB_NOTIFY_FIELD_USER_NAME
          Debug.Print "User Name: " & GetStringFromLPSTR(.pBuf, .cbBuf),
          
        Case JOB_NOTIFY_FIELD_STATUS_STRING
          Debug.Print "Status: " & GetStringFromLPSTR(.pBuf, .cbBuf),
          
        Case JOB_NOTIFY_FIELD_PAGES_PRINTED
          Debug.Print "Pages Printed: " & CStr(.cbBuf),
          
        Case JOB_NOTIFY_FIELD_TOTAL_BYTES
          Debug.Print "Total Bytes: " & CStr(.cbBuf),
          
        Case JOB_NOTIFY_FIELD_TOTAL_PAGES
          Debug.Print "Total Pages: " & CStr(.cbBuf),
          
        Case JOB_NOTIFY_FIELD_STATUS
          Debug.Print "Status: " & CStr(.cbBuf),
          
        Case Else
          Debug.Print "Unparsed filed value = " & .Field,
      End Select
    End Select
  End With
  Debug.Print
End Sub

Private Function GetStringFromLPSTR(ByVal lpstr As Long, ByVal chars As Long) As String
  Dim myStringRes As String
  Dim bString() As Byte
  Dim c As Long
  Dim pos1 As Long
  c = 40
  ' lpstr is wide string that is null terminated
  ' I've messed around in here to try to get it working
  ' but it doesn't look like a string in here at all...
  
  ' finally working I think...
  If chars <= 0 Then Exit Function
  ReDim bString(chars - 1)
  
  CopyMemory bString(0), ByVal lpstr, chars
  myStringRes = StrConv(bString, vbUnicode)
  GetStringFromLPSTR = Left(myStringRes, InStr(1, myStringRes, Chr(0)))
  
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not quitFlag And watching Then
    quitFlag = True
    Cancel = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If hNotification <> 0 Then FindClosePrinterChangeNotification (hNotification)
  If hPrinter <> 0 Then ClosePrinter hPrinter
End Sub
