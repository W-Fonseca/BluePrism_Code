Dim st = GetInstance(handle)
Dim ws = GetWorksheet(handle, wb_in, ws_in)
Dim LoopCondition
 
Orientation = Orientation.ToUpper()
Select Case Orientation
Case "U"
    Dim currentRow = st.ActiveCell.Row
Do
currentRow = currentRow - 1
Loop While currentRow > 1 And (st.ActiveSheet.Rows(currentRow).Hidden Or st.Cells(currentRow, st.ActiveCell.Column).Value Is Nothing)
st.Cells(currentRow, st.ActiveCell.Column).Select
 
Case "D"
    Dim currentRow = st.ActiveCell.Row
Do
currentRow = currentRow + 1
If Is_Visible = True And Is_Empty = True Then
LoopCondition = currentRow > 1 And (st.ActiveSheet.Rows(currentRow).Hidden Or st.Cells(currentRow, st.ActiveCell.Column).Value Is Nothing)
ElseIf Is_Visible = True And Is_Empty <> True Then
LoopCondition = currentRow > 1 And (st.ActiveSheet.Rows(currentRow).Hidden)
ElseIf Is_Visible <> True And Is_Empty = True Then
LoopCondition = currentRow > 1 And st.Cells(currentRow, st.ActiveCell.Column).Value Is Nothing
Else
LoopCondition = currentRow > 1
End If
Loop While LoopCondition
st.Cells(currentRow, st.ActiveCell.Column).Select
Case "L"
Dim currentColumn = st.ActiveCell.Column
If currentColumn > 1 Then
Do
    currentColumn = currentColumn - 1
    If Is_Visible = True And Is_Empty = True Then
LoopCondition = currentColumn > 1 And (st.ActiveSheet.Columns(currentColumn).Hidden Or st.Cells(st.ActiveCell.Row, currentColumn).Value Is Nothing)
ElseIf Is_Visible = True And Is_Empty <> True Then
LoopCondition = currentColumn > 1 And (st.ActiveSheet.Columns(currentColumn).Hidden)
ElseIf Is_Visible <> True And Is_Empty = True Then
LoopCondition = currentColumn > 1 And st.Cells(st.ActiveCell.Row, currentColumn).Value Is Nothing
Else
LoopCondition = currentColumn > 1
End If
Loop While LoopCondition
end if
st.Cells(st.ActiveCell.Row, currentColumn).Select
Case "R"
Dim currentColumn = st.ActiveCell.Column
Do
    currentColumn = currentColumn + 1
    If Is_Visible = True And Is_Empty = True Then
LoopCondition = currentColumn > 1 And (st.ActiveSheet.Columns(currentColumn).Hidden Or st.Cells(st.ActiveCell.Row, currentColumn).Value Is Nothing)
ElseIf Is_Visible = True And Is_Empty <> True Then
LoopCondition = currentColumn > 1 And (st.ActiveSheet.Columns(currentColumn).Hidden)
ElseIf Is_Visible <> True And Is_Empty = True Then
LoopCondition = currentColumn > 1 And st.Cells(st.ActiveCell.Row, currentColumn).Value Is Nothing
Else
LoopCondition = currentColumn > 1
End If
Loop While LoopCondition
    st.Cells(st.ActiveCell.Row, currentColumn).Select
End Select
