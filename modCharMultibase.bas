Attribute VB_Name = "modCharMultibase"



Public Function RevisaCaracterMultibase(CADENA As String) As String
Dim i As Integer
Dim J As Integer
Dim L As String
Dim C As String

    L = ""
    For i = 1 To Len(CADENA)
        C = Mid(CADENA, i, 1)
        J = Asc(C)
        If J > 125 Then
            Select Case J
            Case 128
                C = "�"
            Case 164  '� minuscula
                C = "�"
            Case 165
                'Es la �
                C = "�"
            Case 166
                C = "�"
            Case 167, 186
                C = "�"
            Case 209
            
            Case Else
                
            End Select
        End If
        L = L & C
    Next i
    RevisaCaracterMultibase = L

End Function
