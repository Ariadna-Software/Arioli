Attribute VB_Name = "Flags"
Option Explicit

Public Sub p_ErrorHandler()
   'if an error occurs then capture it and show the description of the error.
    If m_iIgnoreAllErrorsFlag = 1 Then
        Exit Sub
    Else
        Dialog.lblErrorDescription.Caption = Err.Description
        Dialog.Show vbModal, Batchelor
    End If
End Sub
