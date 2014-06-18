ExcelSudokuSolver
=================

Sudoku solver using VBA in Excel

//Code

Sub Solve()
    grid = Range("F3:N11")
    Call recursivelyTry(1, 1, grid)
End Sub

'Recursively tries all valid values for a particular cell for a given grid state
Function recursivelyTry(ByVal r As Integer, ByVal c As Integer, grid As Variant)
    Dim i, j, posValues
    
    'If r = 10, puzzle has been completed
    If Not r = 10 Then
        'Skip cell if it already contains a value. Only the case for given values, because recursion resets cell values on backtrack
        If grid(r, c) = "" Then
            'Only iterate over valid entries for this cell to make algorithm more efficient
            posValues = validEntries(r, c)
            
            'If there are valid values for this cell given the current state of the puzzle
            If Not UBound(posValues) = 0 Then
                For i = LBound(posValues) + 1 To UBound(posValues)
                    grid(r, c) = posValues(i)
                    Call recursivelyTry(r, c, grid)
                    grid(r, c) = "" 'Reset cell value
                Next i
            End If
        Else
            'Note: does not reset cell value here because only given values will hit this condition
            If c = 9 Then
                Call recursivelyTry(r + 1, 1, grid)
            Else
                Call recursivelyTry(r, c + 1, grid)
            End If
        End If
    Else
        'Print out results
        For i = 1 To 9
            For j = 1 To 9
                Cells(2 + i, 5 + j) = grid(i, j)
            Next j
        Next i
        
        End
    End If
End Function

'Finds all valid values for a cell for a given grid state
Function validEntries(ByVal r As Integer, ByVal c As Integer) As Variant
    Dim result(), temp(), count, x, y
    'Temporarily stores which values do not conflict with relative cells (i.e. those in row/col/box)
    temp() = Array(True, True, True, True, True, True, True, True, True, True)
    count = 9 'Tracks the number of valid entries for the cell
    
    'Search for invalid values that appear in the same row/col
    For x = 1 To 9
        If Not grid(x, c) = "" And temp(Val(grid(x, c))) Then
            temp(grid(x, c)) = False
            count = count - 1
        End If
        
        If Not grid(r, x) = "" And temp(Val(grid(r, x))) Then
            temp(grid(r, x)) = False
            count = count - 1
        End If
    Next x
    
    'Search for invalid values that appear in the same box
    For x = (((r - 1) \ 3) * 3 + 1) To (((r - 1) \ 3 + 1) * 3)
        For y = (((c - 1) \ 3) * 3 + 1) To (((c - 1) \ 3 + 1) * 3)
            If Not grid(x, y) = "" And temp(Val(grid(x, y))) Then
                temp(grid(x, y)) = False
                count = count - 1
            End If
        Next y
    Next x
    
    'Build actualy list of valid values
    ReDim result(count)
    count = 1
    For x = 1 To 9
        If temp(x) = True Then
            result(count) = x
            count = count + 1
        End If
    Next x
                
    validEntries = result
End Function

