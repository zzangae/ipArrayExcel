Attribute VB_Name = "Module2"
Sub Clear_All_Delete()
    '▶ IP정렬 D열 삭제
    With Sheets("IP정렬")
        Dim lastRowD As Long
        lastRowD = .Cells(.Rows.Count, "D").End(xlUp).Row
        If lastRowD >= 3 Then .Range("D3:D" & lastRowD).ClearContents
    End With

    '▶ 데이타비교 C열 삭제
    With Sheets("데이타비교")
        Dim lastRowC As Long
        lastRowC = .Cells(.Rows.Count, "C").End(xlUp).Row
        If lastRowC >= 3 Then .Range("C3:C" & lastRowC).ClearContents
    End With

    '▶ 데이타비교 D열 삭제
    With Sheets("데이타비교")
        Dim lastRowD2 As Long
        lastRowD2 = .Cells(.Rows.Count, "D").End(xlUp).Row
        If lastRowD2 >= 3 Then .Range("D3:D" & lastRowD2).ClearContents
    End With
End Sub

