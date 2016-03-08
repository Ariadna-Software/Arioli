Attribute VB_Name = "PublicVariables"
Option Explicit

'Normal Public Variables
Public BtApp As BarTender.Application
Public BtFormat As BarTender.Format
Public m_sNewFormat As String
Public m_sSelectedFormat As String
Public m_bBtVisibility As Boolean
 
'Error Flags Variables
Public m_bOpenFormatFailed As Boolean
Public m_iIgnoreAllErrorsFlag As Integer
Public m_bExitCurrentProcedure As Boolean
