Dim c as range

For Each c In Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row)
Next c
