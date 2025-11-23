Attribute VB_Name = "Module1"
Option Explicit

'===================================================
' 메인 정렬 함수
'===================================================
Sub SortIP_Mixed()

    Dim rng As Range
    Dim arr As Variant
    Dim list() As String
    Dim i As Long, j As Long, n As Long
    Dim tmp As String

    ' 정렬 범위: D3 ~ D열 마지막 데이터까지 자동 선택
    Set rng = Range("D3:D" & Cells(Rows.Count, "D").End(xlUp).Row)

    arr = Application.Transpose(rng.Value)

    ' 공백/빈 셀 제거하여 정리된 배열 생성
    For i = LBound(arr) To UBound(arr)
        If Len(Trim$(CStr(arr(i)))) > 0 Then
            n = n + 1
            ReDim Preserve list(1 To n)
            list(n) = Trim$(CStr(arr(i)))
        End If
    Next i

    If n <= 1 Then Exit Sub

    ' 버블 정렬
    For i = 1 To n - 1
        For j = i + 1 To n
            If CompareIP(list(i), list(j)) > 0 Then
                tmp = list(i)
                list(i) = list(j)
                list(j) = tmp
            End If
        Next j
    Next i

    ' 시트에 정렬 결과 쓰기
    For i = 1 To n
        rng.Cells(i, 1).Value = list(i)
    Next i

    If n < rng.Rows.Count Then
        rng.Offset(n, 0).Resize(rng.Rows.Count - n, 1).ClearContents
    End If

End Sub

'===================================================
' IP 비교 (IPv4 먼저, IPv6 나중)
'===================================================
Function CompareIP(ByVal ip1 As String, ByVal ip2 As String) As Long

    Dim isV4_1 As Boolean, isV4_2 As Boolean
    Dim b1() As Byte, b2() As Byte
    Dim p1() As String, p2() As String
    Dim k As Long

    ip1 = Trim$(ip1)
    ip2 = Trim$(ip2)

    ' IPv4 판별
    isV4_1 = (InStr(ip1, ".") > 0 And InStr(ip1, ":") = 0)
    isV4_2 = (InStr(ip2, ".") > 0 And InStr(ip2, ":") = 0)

    '---------------------------------------
    ' 1) IPv4 우선
    '---------------------------------------
    If isV4_1 And Not isV4_2 Then CompareIP = -1: Exit Function
    If Not isV4_1 And isV4_2 Then CompareIP = 1: Exit Function

    '---------------------------------------
    ' 2) 둘 다 IPv4 → 숫자 비교 (안전)
    '---------------------------------------
    If isV4_1 And isV4_2 Then
        p1 = Split(ip1, ".")
        p2 = Split(ip2, ".")

        If UBound(p1) <> 3 Or UBound(p2) <> 3 Then
            CompareIP = StrComp(ip1, ip2)
            Exit Function
        End If

        For k = 0 To 3
            If ToLongSafe(p1(k)) < ToLongSafe(p2(k)) Then
                CompareIP = -1: Exit Function
            ElseIf ToLongSafe(p1(k)) > ToLongSafe(p2(k)) Then
                CompareIP = 1: Exit Function
            End If
        Next k

        CompareIP = 0
        Exit Function
    End If

    '---------------------------------------
    ' 3) 둘 다 IPv6 → 16바이트 비교
    '---------------------------------------
    b1 = IPToBytes(ip1)
    b2 = IPToBytes(ip2)

    For k = 0 To 15
        If b1(k) < b2(k) Then CompareIP = -1: Exit Function
        If b1(k) > b2(k) Then CompareIP = 1: Exit Function
    Next k

    CompareIP = 0
End Function

'===================================================
' 안전한 숫자 변환 (런타임 오류 13 방지)
'===================================================
Private Function ToLongSafe(ByVal s As String) As Long
    s = Trim$(s)

    If s = "" Then
        ToLongSafe = 0
    ElseIf IsNumeric(s) Then
        ToLongSafe = CLng(s)
    Else
        ToLongSafe = 0
    End If
End Function

'===================================================
' IPv4 → 16바이트 변환
' IPv6 → 16바이트 변환 (압축(::) 포함)
'===================================================
Function IPToBytes(ByVal IP As String) As Byte()
    Dim parts() As String
    Dim bytes() As Byte
    Dim expanded As String, sections() As String
    Dim leftPart As String, rightPart As String
    Dim leftArr() As String, rightArr() As String
    Dim missing As Long, v As Long, k As Long
    Dim i As Long

    IP = Trim$(IP)

    ReDim bytes(15)

    '----------------------------
    ' IPv4 처리
    '----------------------------
    If InStr(IP, ".") > 0 And InStr(IP, ":") = 0 Then
        parts = Split(IP, ".")
        If UBound(parts) <> 3 Then GoTo Bad

        For i = 0 To 3
            If Not IsNumeric(parts(i)) Then GoTo Bad
            If CLng(parts(i)) < 0 Or CLng(parts(i)) > 255 Then GoTo Bad
            bytes(12 + i) = CByte(parts(i))
        Next i

        IPToBytes = bytes
        Exit Function
    End If

    '----------------------------
    ' IPv6 처리
    '----------------------------
    If InStr(IP, ":") > 0 Then

        ' :: 압축 처리
        If InStr(IP, "::") > 0 Then
            Dim tmp() As String
            tmp = Split(IP, "::")
            leftPart = tmp(0)
            If UBound(tmp) >= 1 Then rightPart = tmp(1)

            If leftPart <> "" Then leftArr = Split(leftPart, ":") Else ReDim leftArr(-1)
            If rightPart <> "" Then rightArr = Split(rightPart, ":") Else ReDim rightArr(-1)

            missing = 8 - (UBound(leftArr) + 1) - (UBound(rightArr) + 1)
            If missing < 0 Then GoTo Bad

            expanded = ""
            For i = LBound(leftArr) To UBound(leftArr)
                expanded = expanded & leftArr(i) & ":"
            Next i
            For i = 1 To missing
                expanded = expanded & "0:"
            Next i
            For i = LBound(rightArr) To UBound(rightArr)
                expanded = expanded & rightArr(i) & ":"
            Next i
            If Right$(expanded, 1) = ":" Then expanded = Left$(expanded, Len(expanded) - 1)
        Else
            expanded = IP
        End If

        sections = Split(expanded, ":")
        If UBound(sections) <> 7 Then GoTo Bad

        k = 0
        For i = 0 To 7
            If sections(i) = "" Then sections(i) = "0"
            v = CLng("&H" & sections(i))
            bytes(k) = (v \ 256) And &HFF
            bytes(k + 1) = v And &HFF
            k = k + 2
        Next i

        IPToBytes = bytes
        Exit Function
    End If

Bad:
    For i = 0 To 15: bytes(i) = &HFF: Next
    IPToBytes = bytes
End Function

