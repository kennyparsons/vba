Sub SearchVBA()
    Call findWordInModules(InputBox(Chr(10) & Chr(10) & Chr(10) & "What are you looking for?"))
End Sub

Public Sub findWordInModules(ByVal sSearchTerm As String)
'object.Find(target, startline, startcol, endline, endcol [, wholeword] [, matchcase] [, patternsearch])
'https://msdn.microsoft.com/en-us/library/aa443952(v=vs.60).aspx
 
' VBComponent requires reference to Microsoft Visual Basic for Applications Extensibility
'             or keep it as is and use Late Binding instead
    Dim oComponent            As Object    'VBComponent
 
    For Each oComponent In Application.VBE.ActiveVBProject.VBComponents
        If oComponent.CodeModule.Find(sSearchTerm, 1, 1, -1, -1, False, False, False) = True Then
            Debug.Print "Module: " & oComponent.name  'Name of the current module in which the term was found (at least once)
            'Need to execute a recursive listing of where it is found in the module since it could be found more than once
            Call listLinesinModuleWhereFound(oComponent, sSearchTerm)
        End If
    Next oComponent
End Sub
 
Sub listLinesinModuleWhereFound(ByVal oComponent As Object, ByVal sSearchTerm As String)
    Dim lTotalNoLines         As Long   'total number of lines within the module being examined
    Dim lLineNo               As Long   'will return the line no where the term is found
 
    lLineNo = 1
    With oComponent
        lTotalNoLines = .CodeModule.CountOfLines
        Do While .CodeModule.Find(sSearchTerm, lLineNo, 1, -1, -1, False, False, False) = True
            Debug.Print vbTab & "Line No:" & lLineNo & Trim(.CodeModule.Lines(lLineNo, 1))  'Remove any padding spaces
            lLineNo = lLineNo + 1    'Restart the search at the next line looking for the next occurence
        Loop
    End With
End Sub
