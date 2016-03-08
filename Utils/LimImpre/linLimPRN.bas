Attribute VB_Name = "linLimPRN"
Option Explicit

'Constantes
Private Const NULLPTR = 0&
' Constants for DEVMODE
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
' Constants for DocumentProperties
Private Const DM_MODIFY = 8
Private Const DM_COPY = 2
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_OUT_BUFFER = DM_COPY
' Constants for dmOrientation
Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2
' Constants for dm Print Quality
Private Const DMRES_DRAFT = (-1)
Private Const DMRES_HIGH = (-4)
Private Const DMRES_LOW = (-2)
Private Const DMRES_MEDIUM = (-3)
' Constants for dmTTOption
Private Const DMTT_BITMAP = 1
Private Const DMTT_DOWNLOAD = 2
Private Const DMTT_DOWNLOAD_OUTLINE = 4
Private Const DMTT_SUBDEV = 3
' Constants for dmColor
Private Const DMCOLOR_COLOR = 2
Private Const DMCOLOR_MONOCHROME = 1
  


Private Declare Function OpenPrinter Lib "winspool.drv" _
    Alias "OpenPrinterA" (ByVal pPrinterName As String, _
    phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" _
    (ByVal hPrinter As Long) As Long

Public Declare Function FindFirstPrinterChangeNotificationLong Lib "winspool.drv" Alias "FindFirstPrinterChangeNotification" _
  (ByVal hPrinter As Long, ByVal fdwFlags As Long, ByVal fdwOptions As Long, ByVal lpPrinterNotifyOptions As Long) As Long

Private Type DEVMODE
    dmDeviceName(1 To CCHDEVICENAME) As Byte
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName(1 To CCHFORMNAME) As Byte
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type


Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Const INFINITE = &HFFFF ' Infinite timeout

Declare Function FindNextPrinterChangeNotificationByLong Lib "winspool.drv" Alias "FindNextPrinterChangeNotification" _
    (ByVal hChange As Long, pdwChange As Long, pPrinterOptions As PRINTER_NOTIFY_OPTIONS, ppPrinterNotifyInfo As Long) As Long

Public Type PRINTER_NOTIFY_INFO_DATA
  Type As Integer
  Field As Integer
  Reserved As Long
  id As Long
  adwData(0 To 1) As Long
End Type

Public Type PRINTER_NOTIFY_INFO
  dwVersion As Long
  dwFlags As Long
  dwCount As Long
End Type

Private Type PRINTER_DEFAULTS
  pDatatype As String
  pDevMode As DEVMODE
  DesiredAccess As Long
End Type


Declare Function FreePrinterNotifyInfoByLong Lib "winspool.drv" Alias "FreePrinterNotifyInfo" (ByVal pInfo As Long) As Long

Declare Sub CopyMemoryPRINTER_NOTIFY_INFO Lib "kernel32" Alias "RtlMoveMemory" (Destination As PRINTER_NOTIFY_INFO, ByVal Source As Long, ByVal Length As Long)
Declare Sub CopyMemoryPRINTER_NOTIFY_INFO_DATA Lib "kernel32" Alias "RtlMoveMemory" (Destination As PRINTER_NOTIFY_INFO_DATA, ByVal Source As Long, ByVal Length As Long)
Dim aData() As PRINTER_NOTIFY_INFO_DATA

Declare Function FindClosePrinterChangeNotification Lib "winspool.drv" (ByVal hChange As Long) As Long

'See the Getting the status of the selected printer from Visual Basic article for more details about this.
'Ask for the notifications you are interested in
'There are a vast number of events that can happen to a printer or to a print job. You can request notification whenever one or more of them happens by creating a PRINTER_NOTIFY_OPTIONS variable which you pass to the initial request to set up a notification object using the FindFirstPrinterChangeNotification API call.
'The types of printer notification are:

Public Enum PrinterChangeNotifications
    PRINTER_CHANGE_ADD_PRINTER = &H1
    PRINTER_CHANGE_SET_PRINTER = &H2
    PRINTER_CHANGE_DELETE_PRINTER = &H4
    PRINTER_CHANGE_FAILED_CONNECTION_PRINTER = &H8
    PRINTER_CHANGE_PRINTER = &HFF
    PRINTER_CHANGE_ADD_JOB = &H100
    PRINTER_CHANGE_SET_JOB = &H200
    PRINTER_CHANGE_DELETE_JOB = &H400
    PRINTER_CHANGE_WRITE_JOB = &H800
    PRINTER_CHANGE_JOB = &HFF00
    PRINTER_CHANGE_ADD_FORM = &H10000
    PRINTER_CHANGE_SET_FORM = &H20000
    PRINTER_CHANGE_DELETE_FORM = &H40000
    PRINTER_CHANGE_FORM = &H70000
    PRINTER_CHANGE_ADD_PORT = &H100000
    PRINTER_CHANGE_CONFIGURE_PORT = &H200000
    PRINTER_CHANGE_DELETE_PORT = &H400000
    PRINTER_CHANGE_PORT = &H700000
    PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
    PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
    PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
    PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
    PRINTER_CHANGE_SET_PRINTER_DRIVER = &H20000000
    PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
    PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
    PRINTER_CHANGE_TIMEOUT = &H80000000
End Enum

'And the types of print job notification are:

Public Enum JobChangeNotificationFields
    JOB_NOTIFY_FIELD_PRINTER_NAME = &H0
    JOB_NOTIFY_FIELD_MACHINE_NAME = &H1
    JOB_NOTIFY_FIELD_PORT_NAME = &H2
    JOB_NOTIFY_FIELD_USER_NAME = &H3
    JOB_NOTIFY_FIELD_NOTIFY_NAME = &H4
    JOB_NOTIFY_FIELD_DATATYPE = &H5
    JOB_NOTIFY_FIELD_PRINT_PROCESSOR = &H6
    JOB_NOTIFY_FIELD_PARAMETERS = &H7
    JOB_NOTIFY_FIELD_DRIVER_NAME = &H8
    JOB_NOTIFY_FIELD_DEVMODE = &H9
    JOB_NOTIFY_FIELD_STATUS = &HA
    JOB_NOTIFY_FIELD_STATUS_STRING = &HB
    JOB_NOTIFY_FIELD_SECURITY_DESCRIPTOR = &HC
    JOB_NOTIFY_FIELD_DOCUMENT = &HD
    JOB_NOTIFY_FIELD_PRIORITY = &HE
    JOB_NOTIFY_FIELD_POSITION = &HF
    JOB_NOTIFY_FIELD_SUBMITTED = &H10
    JOB_NOTIFY_FIELD_START_TIME = &H11
    JOB_NOTIFY_FIELD_UNTIL_TIME = &H12
    JOB_NOTIFY_FIELD_TIME = &H13
    JOB_NOTIFY_FIELD_TOTAL_PAGES = &H14
    JOB_NOTIFY_FIELD_PAGES_PRINTED = &H15
    JOB_NOTIFY_FIELD_TOTAL_BYTES = &H16
    JOB_NOTIFY_FIELD_BYTES_PRINTED = &H17
End Enum

'However in order to prevent unnecessary notifications which would potentially slow down your system, the notification events only trigger for the events which you have set them to monitor. You do this by creating a PRINTER_NOTIFY_OPTIONS_TYPE record for the printer events you need and one for the job events you need and putting these in a PRINTER_NOTIFY_OPTIONS variable that we pass to the notification API calls. For example, if we wish to be notified when the printer name, share name and status change or when a job is printed we woud do so thus:

'\\ Declarations
Public Type PRINTER_NOTIFY_OPTIONS
    Version As Long '\\should be set to 2
    Flags As Long
    Count As Long
    lpPrintNotifyOptions As Long
End Type

Public Type PRINTER_NOTIFY_OPTIONS_TYPE
    Type As Integer
    Reserved_0 As Integer
    Reserved_1 As Long
    Reserved_2 As Long
    Count As Long
    pFields As Long
End Type

Private PrintOptions As PRINTER_NOTIFY_OPTIONS
Private PrinterNotifyOptions(0 To 1) As PRINTER_NOTIFY_OPTIONS_TYPE

Private mhPrinter As Long

Private mData As PRINTER_NOTIFY_INFO

'\\ Initialising the PrintOptions
Private Sub InitialiseNotifyOptions()

     With PrintOptions
          .Version = 2 '\\ This must be set to 2
          .Count = 2 '\\ There is job notification and printer notification
          '\\ The type of printer events we are interested in...
          With PrinterNotifyOptions(0)
            '.Type = PRINTER_NOTIFY_TYPE
            ReDim pFieldsPrinter(0 To 19) As Integer
            '\\ Add the list of printer events you are interested in being notified about
            '\\ to this list. Note that the fewer notifications you ask for the less of a
            '\\ burden your app place upon the system.
            pFieldsPrinter(0) = PRINTER_NOTIFY_FIELD_PRINTER_NAME
            pFieldsPrinter(1) = PRINTER_NOTIFY_FIELD_SHARE_NAME
            pFieldsPrinter(2) = PRINTER_NOTIFY_FIELD_STATUS
            .Count = (UBound(pFieldsPrinter) - LBound(pFieldsPrinter)) + 1 '\\ Add one as the array is zero based
            .pFields = VarPtr(pFieldsPrinter(0))
          End With
          '\\ The type of print job events we are interested in...
          With PrinterNotifyOptions(1)
            .Type = JOB_NOTIFY_TYPE
            '\\ Add the list of print job events you are interested in being notified about
            '\\ to this list. Note that the fewer notifications you ask for the less of a
            '\\ burden your app place upon the system.
            ReDim pFieldsJob(0 To 22) As Integer
            pFieldsJob(0) = JOB_NOTIFY_FIELD_PRINTER_NAME
            pFieldsJob(1) = JOB_NOTIFY_FIELD_MACHINE_NAME
            pFieldsJob(2) = JOB_NOTIFY_FIELD_PORT_NAME
            pFieldsJob(3) = JOB_NOTIFY_FIELD_USER_NAME
            pFieldsJob(4) = JOB_NOTIFY_FIELD_NOTIFY_NAME
            pFieldsJob(5) = JOB_NOTIFY_FIELD_DATATYPE
            pFieldsJob(6) = JOB_NOTIFY_FIELD_PRINT_PROCESSOR
            pFieldsJob(7) = JOB_NOTIFY_FIELD_PARAMETERS
            pFieldsJob(8) = JOB_NOTIFY_FIELD_DRIVER_NAME
            pFieldsJob(9) = JOB_NOTIFY_FIELD_DEVMODE
            pFieldsJob(10) = JOB_NOTIFY_FIELD_STATUS
            pFieldsJob(11) = JOB_NOTIFY_FIELD_STATUS_STRING
            pFieldsJob(12) = JOB_NOTIFY_FIELD_DOCUMENT
            pFieldsJob(13) = JOB_NOTIFY_FIELD_PRIORITY
            pFieldsJob(14) = JOB_NOTIFY_FIELD_POSITION
            pFieldsJob(15) = JOB_NOTIFY_FIELD_SUBMITTED
            pFieldsJob(16) = JOB_NOTIFY_FIELD_START_TIME
            pFieldsJob(17) = JOB_NOTIFY_FIELD_UNTIL_TIME
            pFieldsJob(18) = JOB_NOTIFY_FIELD_TIME
            pFieldsJob(19) = JOB_NOTIFY_FIELD_TOTAL_PAGES
            pFieldsJob(20) = JOB_NOTIFY_FIELD_PAGES_PRINTED
            pFieldsJob(21) = JOB_NOTIFY_FIELD_TOTAL_BYTES
            .Count = (UBound(pFieldsJob) - LBound(pFieldsJob)) + 1 '\\ Add one as the array is zero based
            .pFields = VarPtr(pFieldsJob(0))
          End With
          .lpPrintNotifyOptions = VarPtr(PrinterNotifyOptions(0))
    End With

End Sub

Private Sub Main()

    mEventHandle = FindFirstPrinterChangeNotificationLong(mhPrinter, 0, 0, VarPtr(PrintOptions))
    
    Call WaitForSingleObject(mEventHandle, INFINITE)
    
    Dim lpPrintInfoBuffer As Long
    Call FindNextPrinterChangeNotificationByLong(mEventHandle, pdwChange, PrintOptions, lpPrintInfoBuffer)
    
    Call FindClosePrinterChangeNotification(mEventHandle)
    
    
    
Call CopyMemoryPRINTER_NOTIFY_INFO(mData, lpPrintInfoBuffer, Len(mData))
'\\ mData contains a valid PRINTER_NOTIFY_INFO structure
If mData.dwCount > 0 Then
  ReDim aData(1 To mData.dwCount) As PRINTER_NOTIFY_INFO_DATA
  '\\ Copy the structure in full
  Call CopyMemoryPRINTER_NOTIFY_INFO_DATA(aData(1), lpPrintInfoBuffer + Len(mData), Len(aData(1)) * mData.dwCount)
  '\\ Operate on the changes

  '\\ and clear out the buffer
  Erase aData
  Call FreePrinterNotifyInfoByLong(lpPrintInfoBuffer)
    
    
    
    
End Sub



