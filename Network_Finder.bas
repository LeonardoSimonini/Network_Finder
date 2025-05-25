Attribute VB_Name = "Modulo1"


Public Function FIND_NETWORK(ip As String, rg As Range) As Variant
    Dim biggest_network_number As Integer: biggest_network_number = 0
    For Each cll In rg.Cells
        Dim numb_str As String
        Dim net_str As String
        Dim numb_int As Integer
        Dim arStrNetwork() As String
        Dim arStrIP() As String
        Dim arSubmask(0 To 3) As Byte
        Dim arNetwork(0 To 3) As Integer
        Dim arIP(0 To 3) As Integer
        Dim arTrunkIP(0 To 3) As Integer
        Dim pos As Integer
        Dim str_len As Integer
        Dim module As Integer
        Dim quart As Byte
        Dim trunkated As Integer
        Dim count As Integer: count = 0
        pos = InStr(cll.Value, "/")
        str_len = Len(cll.Value) - pos
        numb_str = Right(cll.Value, str_len)
        numb_int = CInt(numb_str)
        For i = 0 To 3
            If i = 0 Then
                If numb_int \ 8 > 0 Then
                    arSubmask(0) = 255
                ElseIf numb_int \ 8 = 0 Then
                    module = numb_int Mod 8
                    arSubmask(0) = GET_BYTE(module)
                Else
                    arSubmask(0) = 0
                End If
            ElseIf i = 1 Then
                If numb_int \ 8 > 1 Then
                    arSubmask(1) = 255
                ElseIf numb_int \ 8 = 1 Then
                    module = numb_int Mod 8
                    arSubmask(1) = GET_BYTE(module)
                Else
                    arSubmask(1) = 0
                End If
            ElseIf i = 2 Then
                If numb_int \ 8 > 2 Then
                    arSubmask(2) = 255
                ElseIf numb_int \ 8 = 2 Then
                    module = numb_int Mod 8
                    arSubmask(2) = GET_BYTE(module)
                Else
                    arSubmask(2) = 0
                End If
            ElseIf i = 3 Then
                If numb_int \ 8 > 3 Then
                    arSubmask(3) = 255
                ElseIf numb_int \ 8 = 3 Then
                    module = numb_int Mod 8
                    arSubmask(3) = GET_BYTE(module)
                Else
                    arSubmask(3) = 0
                End If
            End If
        Next i
        
        net_str = Left(cll.Value, pos - 1)
        arStrNetwork() = Split(net_str, ".")
        arStrIP() = Split(ip, ".")
        For j = 0 To 3
            arNetwork(j) = CInt(arStrNetwork(j))
            arIP(j) = CInt(arStrIP(j))
        Next j
        
        For k = 0 To 3
            If arSubmask(k) = 255 Then
                arTrunkIP(k) = arIP(k)
            ElseIf arSubmask(k) = 0 Then
                arTrunkIP(k) = 0
            Else
                quart = CByte(arIP(k))
                trunkated = quart And arSubmask(k)
                arTrunkIP(k) = CInt(trunkated)
            End If
        Next k
        
        For l = 0 To 3
            If arNetwork(l) = arTrunkIP(l) Then
                count = count + 1
                If count >= 4 Then
                    If biggest_network_number < numb_int Then
                        biggest_network_number = numb_int
                        FIND_NETWORK = cll.Value
                    End If
                End If
            Else
                Exit For
            End If
        Next l
    Next cll
    If biggest_network_number = 0 Then
        FIND_NETWORK = CVErr(xlErrNA)
    End If
End Function


Function GET_BYTE(intero As Integer) As Byte
    Select Case intero
        Case 0
            GET_BYTE = 0
        Case 1
            GET_BYTE = 128
        Case 2
            GET_BYTE = 192
        Case 3
            GET_BYTE = 224
        Case 4
            GET_BYTE = 240
        Case 5
            GET_BYTE = 248
        Case 6
            GET_BYTE = 252
        Case 7
            GET_BYTE = 254
        Case 8
            GET_BYTE = 255
    End Select
End Function
