Attribute VB_Name = "Module1"
Public rVal As Integer
Public gVal As Integer
Public bVal As Integer

Public nextRVal As Integer
Public nextGVal As Integer
Public nextBVal As Integer

Public Const MIN_ROW = 5
Public Const MAX_ROW = 26
Public Const MIN_COL = 5
Public Const MAX_COL = 16

Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const VK_LEFT As Long = 37
Const VK_DOWN As Long = 40
Const VK_RIGHT As Long = 39
Const VK_UP As Long = 38
Const VK_HOME As Long = 36

Public currentRow1 As Integer
Public currentCol1 As Integer
Public currentRow2 As Integer
Public currentCol2 As Integer
Public currentRow3 As Integer
Public currentCol3 As Integer
Public currentRow4 As Integer
Public currentCol4 As Integer

Public nextRow1 As Integer
Public nextCol1 As Integer
Public nextRow2 As Integer
Public nextCol2 As Integer
Public nextRow3 As Integer
Public nextCol3 As Integer
Public nextRow4 As Integer
Public nextCol4 As Integer

Public currentShape As Integer
Public nextShape As Integer
Public orientation As Integer
Public pivotRow As Integer
Public pivotCol As Integer

Public startTime As Single

Public shapeJustPlaced As Boolean

Option Explicit

Sub PickColors()
    Select Case currentShape
        Case 1 'light blue line
            rVal = 0
            gVal = 255
            bVal = 255
        Case 2 'dark blue L
            rVal = 0
            gVal = 0
            bVal = 255
        Case 3 'orange L
            rVal = 255
            gVal = 153
            bVal = 0
        Case 4 'yellow square
            rVal = 255
            gVal = 255
            bVal = 0
        Case 5 'green trapezoid
            rVal = 60
            gVal = 250
            bVal = 78
        Case 6 'purple T
            rVal = 112
            gVal = 48
            bVal = 160
        Case 7 'red trapezoid
            rVal = 255
            gVal = 0
            bVal = 0
        Case Else
            rVal = 0
            gVal = 0
            bVal = 0
    End Select
End Sub



Sub PickShapeCoords()
    currentRow2 = pivotRow
    currentCol2 = pivotCol
    Select Case currentShape
        Case 1 'Line done
            Select Case orientation
                Case 1
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol - 1
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol + 1
                    
                    currentRow4 = pivotRow
                    currentCol4 = pivotCol + 2
                    
                Case 2
                    currentRow1 = pivotRow - 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow + 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow + 2
                    currentCol4 = pivotCol
                    
                Case 3
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol + 1
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol - 1
                    
                    currentRow4 = pivotRow
                    currentCol4 = pivotCol - 2
                    
                Case 4
                    currentRow1 = pivotRow + 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow - 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow - 2
                    currentCol4 = pivotCol
            End Select
        Case 2 'dark blue L done
            Select Case orientation
                Case 1
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol + 1
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol - 1
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol - 1
                    
                Case 2
                    currentRow1 = pivotRow + 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow - 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol + 1
                    
                Case 3
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol - 1
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol + 1
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol + 1
                    
                Case 4
                    currentRow1 = pivotRow - 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow + 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol - 1
            End Select
        Case 3 'orange L done
            Select Case orientation
                Case 1
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol - 1
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol + 1
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol + 1
                    
                Case 2
                    currentRow1 = pivotRow - 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow + 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol + 1
                    
                Case 3
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol + 1
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol - 1
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol - 1
                    
                Case 4
                    currentRow1 = pivotRow + 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow - 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol - 1
            End Select
        Case 4 'Square done
            currentRow1 = pivotRow
            currentCol1 = pivotCol - 1
            
            currentRow3 = pivotRow + 1
            currentCol3 = pivotCol
            
            currentRow4 = pivotRow + 1
            currentCol4 = pivotCol - 1
            
        Case 5 'green Trapezoid done
            Select Case orientation
                Case 1
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol - 1
                    
                    currentRow3 = pivotRow - 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol + 1
                    
                Case 2
                    currentRow1 = pivotRow - 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol + 1
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol + 1
                    
                Case 3
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol + 1
                    
                    currentRow3 = pivotRow + 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol - 1
                    
                Case 4
                    currentRow1 = pivotRow + 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol - 1
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol - 1
            End Select
        Case 6 'T done
            Select Case orientation
                Case 1
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol - 1
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol + 1
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol
                    
                Case 2
                    currentRow1 = pivotRow - 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow + 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow
                    currentCol4 = pivotCol + 1
                    
                Case 3
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol + 1
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol - 1
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol
                    
                Case 4
                    currentRow1 = pivotRow + 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow - 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow
                    currentCol4 = pivotCol - 1
            End Select
        Case 7 'red Trapezoid done
            Select Case orientation
                Case 1
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol + 1
                    
                    currentRow3 = pivotRow - 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol - 1
                    
                Case 2
                    currentRow1 = pivotRow + 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol + 1
                    
                    currentRow4 = pivotRow - 1
                    currentCol4 = pivotCol + 1
                    
                Case 3
                    currentRow1 = pivotRow
                    currentCol1 = pivotCol - 1
                    
                    currentRow3 = pivotRow + 1
                    currentCol3 = pivotCol
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol + 1
                    
                Case 4
                    currentRow1 = pivotRow - 1
                    currentCol1 = pivotCol
                    
                    currentRow3 = pivotRow
                    currentCol3 = pivotCol - 1
                    
                    currentRow4 = pivotRow + 1
                    currentCol4 = pivotCol - 1
            End Select
    End Select
End Sub

Sub PickShape()
    
    'currentShape = 1
    currentShape = nextShape
    nextShape = Application.WorksheetFunction.RandBetween(1, 7)
    DrawNextShape
    orientation = 1
    PickColors
    
End Sub

Sub PickNextShapeCoords()
    
    Select Case nextShape
        Case 1 'line
            nextRow1 = MIN_ROW + 8
            nextCol1 = MAX_COL + 4
            nextRow2 = MIN_ROW + 9
            nextCol2 = MAX_COL + 4
            nextRow3 = MIN_ROW + 10
            nextCol3 = MAX_COL + 4
            nextRow4 = MIN_ROW + 11
            nextCol4 = MAX_COL + 4
        Case 2 'dark blue L
            nextRow1 = MIN_ROW + 9
            nextCol1 = MAX_COL + 5
            nextRow2 = MIN_ROW + 10
            nextCol2 = MAX_COL + 5
            nextRow3 = MIN_ROW + 11
            nextCol3 = MAX_COL + 5
            nextRow4 = MIN_ROW + 11
            nextCol4 = MAX_COL + 4
        Case 3 'orange L
            nextRow1 = MIN_ROW + 9
            nextCol1 = MAX_COL + 4
            nextRow2 = MIN_ROW + 10
            nextCol2 = MAX_COL + 4
            nextRow3 = MIN_ROW + 11
            nextCol3 = MAX_COL + 4
            nextRow4 = MIN_ROW + 11
            nextCol4 = MAX_COL + 5
        Case 4 'square
            nextRow1 = MIN_ROW + 9
            nextCol1 = MAX_COL + 4
            nextRow2 = MIN_ROW + 9
            nextCol2 = MAX_COL + 5
            nextRow3 = MIN_ROW + 10
            nextCol3 = MAX_COL + 4
            nextRow4 = MIN_ROW + 10
            nextCol4 = MAX_COL + 5
        Case 5 'green trapezoid
            nextRow1 = MIN_ROW + 9
            nextCol1 = MAX_COL + 4
            nextRow2 = MIN_ROW + 10
            nextCol2 = MAX_COL + 4
            nextRow3 = MIN_ROW + 10
            nextCol3 = MAX_COL + 5
            nextRow4 = MIN_ROW + 11
            nextCol4 = MAX_COL + 5
        Case 6 'purple T
            nextRow1 = MIN_ROW + 9
            nextCol1 = MAX_COL + 4
            nextRow2 = MIN_ROW + 10
            nextCol2 = MAX_COL + 4
            nextRow3 = MIN_ROW + 11
            nextCol3 = MAX_COL + 4
            nextRow4 = MIN_ROW + 10
            nextCol4 = MAX_COL + 5
        Case 7 'red trapezoid
            nextRow1 = MIN_ROW + 9
            nextCol1 = MAX_COL + 5
            nextRow2 = MIN_ROW + 10
            nextCol2 = MAX_COL + 5
            nextRow3 = MIN_ROW + 10
            nextCol3 = MAX_COL + 4
            nextRow4 = MIN_ROW + 11
            nextCol4 = MAX_COL + 4
        End Select
        
End Sub

Sub PickNextColors()
    Select Case nextShape
        Case 1 'light blue line
            nextRVal = 0
            nextGVal = 255
            nextBVal = 255
        Case 2 'dark blue L
            nextRVal = 0
            nextGVal = 0
            nextBVal = 255
        Case 3 'orange L
            nextRVal = 255
            nextGVal = 153
            nextBVal = 0
        Case 4 'yellow square
            nextRVal = 255
            nextGVal = 255
            nextBVal = 0
        Case 5 'green trapezoid
            nextRVal = 60
            nextGVal = 250
            nextBVal = 78
        Case 6 'purple T
            nextRVal = 112
            nextGVal = 48
            nextBVal = 160
        Case 7 'red trapezoid
            nextRVal = 255
            nextGVal = 0
            nextBVal = 0
        Case Else
            nextRVal = 0
            nextGVal = 0
            nextBVal = 0
    End Select
End Sub

Sub DrawNextShape()

    Cells(nextRow1, nextCol1).Interior.Color = xlNone
    Cells(nextRow2, nextCol2).Interior.Color = xlNone
    Cells(nextRow3, nextCol3).Interior.Color = xlNone
    Cells(nextRow4, nextCol4).Interior.Color = xlNone
    
    PickNextShapeCoords
    PickNextColors
    
    Cells(nextRow1, nextCol1).Interior.Color = RGB(nextRVal, nextGVal, nextBVal)
    Cells(nextRow2, nextCol2).Interior.Color = RGB(nextRVal, nextGVal, nextBVal)
    Cells(nextRow3, nextCol3).Interior.Color = RGB(nextRVal, nextGVal, nextBVal)
    Cells(nextRow4, nextCol4).Interior.Color = RGB(nextRVal, nextGVal, nextBVal)
    
End Sub

Function CheckBounds() As Boolean

    Dim rowBeingChecked As Variant
    Dim rowsBeingChecked(1 To 4)
    
    rowsBeingChecked(1) = currentRow1
    rowsBeingChecked(2) = currentRow2
    rowsBeingChecked(3) = currentRow3
    rowsBeingChecked(4) = currentRow4
    
    For Each rowBeingChecked In rowsBeingChecked
        If rowBeingChecked >= MAX_ROW Or rowBeingChecked <= MIN_ROW Then
            CheckBounds = False
            Exit Function
        End If
    Next rowBeingChecked
    
    Dim colBeingChecked As Variant
    Dim colsBeingChecked(1 To 4)
    
    colsBeingChecked(1) = currentCol1
    colsBeingChecked(2) = currentCol2
    colsBeingChecked(3) = currentCol3
    colsBeingChecked(4) = currentCol4
    
    For Each colBeingChecked In colsBeingChecked
        If colBeingChecked >= MAX_COL Or colBeingChecked <= MIN_ROW Then
            CheckBounds = False
            Exit Function
        End If
    Next colBeingChecked
    
    CheckBounds = True
    
End Function
Sub EraseShape()

    If currentRow1 = MIN_ROW Or currentRow1 = MAX_ROW Or currentCol1 = MIN_COL Or currentCol1 = MAX_COL Then
        Cells(currentRow1, currentCol1).Interior.Color = RGB(77, 77, 77)
    Else
        Cells(currentRow1, currentCol1).Interior.Color = xlNone
    End If
    
    If currentRow2 = MIN_ROW Or currentRow2 = MAX_ROW Or currentCol2 = MIN_COL Or currentCol2 = MAX_COL Then
        Cells(currentRow2, currentCol2).Interior.Color = RGB(77, 77, 77)
    Else
        Cells(currentRow2, currentCol2).Interior.Color = xlNone
    End If
    
    If currentRow3 = MIN_ROW Or currentRow3 = MAX_ROW Or currentCol3 = MIN_COL Or currentCol3 = MAX_COL Then
        Cells(currentRow3, currentCol3).Interior.Color = RGB(77, 77, 77)
    Else
        Cells(currentRow3, currentCol3).Interior.Color = xlNone
    End If
    
    If currentRow4 = MIN_ROW Or currentRow4 = MAX_ROW Or currentCol4 = MIN_COL Or currentCol4 = MAX_COL Then
        Cells(currentRow4, currentCol4).Interior.Color = RGB(77, 77, 77)
    Else
        Cells(currentRow4, currentCol4).Interior.Color = xlNone
    End If
    
End Sub

Sub DrawShape()
    
    Cells(currentRow1, currentCol1).Interior.Color = RGB(rVal, gVal, bVal)
    Cells(currentRow2, currentCol2).Interior.Color = RGB(rVal, gVal, bVal)
    Cells(currentRow3, currentCol3).Interior.Color = RGB(rVal, gVal, bVal)
    Cells(currentRow4, currentCol4).Interior.Color = RGB(rVal, gVal, bVal)
    
End Sub

Function CheckForBlocks() As Boolean

    If Cells(currentRow1, currentCol1).Interior.ColorIndex <> xlNone Then
        CheckForBlocks = False
    ElseIf Cells(currentRow2, currentCol2).Interior.ColorIndex <> xlNone Then
        CheckForBlocks = False
    ElseIf Cells(currentRow3, currentCol3).Interior.ColorIndex <> xlNone Then
        CheckForBlocks = False
    ElseIf Cells(currentRow4, currentCol4).Interior.ColorIndex <> xlNone Then
        CheckForBlocks = False
    Else
        CheckForBlocks = True
    End If
    
End Function

Function CheckRow(currentRow As Integer) As Boolean
    Dim currentRange As Range
    Set currentRange = Range(Cells(currentRow, MIN_COL + 1), Cells(currentRow, MAX_COL - 1))
    Dim c As Range
    For Each c In currentRange
        If c.Interior.ColorIndex = xlNone Then
            CheckRow = False
            Exit Function
        End If
    Next c
    
    CheckRow = True

End Function

Sub ClearRow(currentRow As Integer)

    If (CheckRow(currentRow)) Then
        Dim currentRange As Range
        Set currentRange = Range(Cells(currentRow, MIN_COL + 1), Cells(currentRow, MAX_COL - 1))
        currentRange.Interior.Color = xlNone
        
        Dim currentCol As Integer
        For currentCol = MIN_COL + 1 To MAX_COL - 1
            ShiftRows (currentCol)
        Next currentCol
    
    End If
    
End Sub
Sub ShiftRows(currentCol As Integer)
    
    Dim currentNumColoredSquares As Integer: currentNumColoredSquares = 0
    Dim coloredSquaresCol As New Collection
    Dim rowIndex As Integer
    For rowIndex = MIN_ROW + 1 To MAX_ROW - 1
        If Cells(rowIndex, currentCol).Interior.ColorIndex <> xlNone Then
            currentNumColoredSquares = currentNumColoredSquares + 1
            coloredSquaresCol.Add rowIndex
            
        End If
        
    Next rowIndex
    
    If currentNumColoredSquares > 0 Then
    
        Dim coloredSquares() As Integer
        ReDim coloredSquares(1 To coloredSquaresCol.Count)
        
        Dim colIndex As Integer
        For colIndex = 1 To coloredSquaresCol.Count
            coloredSquares(colIndex) = coloredSquaresCol(colIndex)
        Next colIndex
    
        Dim checkRowIndex As Integer: checkRowIndex = MAX_ROW - 1
        Dim currentColor As Long
        
        For rowIndex = UBound(coloredSquares) To LBound(coloredSquares) Step -1
            
            currentColor = Cells(coloredSquares(rowIndex), currentCol).Interior.Color
            
            While (Cells(checkRowIndex, currentCol).Interior.ColorIndex <> xlNone)
                
                checkRowIndex = checkRowIndex - 1
                
            Wend
            
            Cells(coloredSquares(rowIndex), currentCol).Interior.Color = xlNone
            
            Cells(checkRowIndex, currentCol).Interior.Color = currentColor
            
            checkRowIndex = MAX_ROW - 1
            
        Next rowIndex
    
    End If
    
    For rowIndex = MIN_ROW + 1 To MAX_ROW - 1
    
        ClearRow (rowIndex)
        
    Next rowIndex
    
End Sub

Sub TestClear()
    ClearRow 8
End Sub

Sub SpawnShape()

    PickShape
    
    pivotCol = ((MAX_COL - MIN_COL) / 2) + MIN_COL
    
    If currentShape = 1 Or currentShape = 4 Then
        pivotRow = MIN_ROW + 1
    Else
        pivotRow = MIN_ROW + 2
    End If
        
    PickShapeCoords
    
    If (CheckForBlocks) Then
        
        DrawShape
        
    Else
        
        MsgBox "GAME OVER"
    
        End
    
    End If
    
End Sub

Sub SpawnFirstShape()

    currentShape = Application.WorksheetFunction.RandBetween(1, 7)
    nextShape = Application.WorksheetFunction.RandBetween(1, 7)
    PickNextShapeCoords
    PickNextColors
    DrawNextShape
    orientation = 1
    PickColors
    
    pivotCol = ((MAX_COL - MIN_COL) / 2) + MIN_COL
    
    If currentShape = 1 Or currentShape = 4 Then
        pivotRow = MIN_ROW + 1
    Else
        pivotRow = MIN_ROW + 2
    End If
        
    PickShapeCoords
    
    If (CheckForBlocks) Then
        
        DrawShape
        
    Else
        
        MsgBox "GAME OVER"
    
        End
    
    End If
    
End Sub

Sub ShiftShape(direction As Integer)

    'direction = -1 -----> Shift left
    'direction =  1 -----> Shift right
    
    pivotCol = pivotCol + direction
    
    If CheckMove Then
        DrawShape
    Else
        pivotCol = pivotCol - direction
        PickShapeCoords
        DrawShape
    End If
    
End Sub

Sub RotateShape(direction As Integer)
    
    'direction = -1 -----> Counter-clockwise
    'direction =  1 -----> Clockwise
    Dim oldOrientation As Integer
    oldOrientation = orientation
    
    If orientation + direction < 1 Then
        orientation = 4
    ElseIf orientation + direction > 4 Then
        orientation = 1
    Else
        orientation = orientation + direction
    End If
    
    If CheckMove Then
        DrawShape
    Else
        orientation = oldOrientation
        PickShapeCoords
        DrawShape
    End If
    
End Sub

Sub DrawBorders()

    Range(Cells(MIN_ROW, MIN_COL), Cells(MAX_ROW, MAX_COL + 7)).ColumnWidth = 3
    Range(Cells(MIN_ROW, MIN_COL), Cells(MAX_ROW, MAX_COL + 7)).Interior.Color = xlNone
    
    'Cells(MIN_ROW, MIN_COL).Interior.Color = RGB(77, 77, 77)
    'Cells(MIN_ROW, MAX_COL).Interior.Color = RGB(77, 77, 77)
    Range(Cells(MIN_ROW, MIN_COL), Cells(MIN_ROW, MAX_COL)).Interior.Color = RGB(77, 77, 77)
    Range(Cells(MAX_ROW, MIN_COL), Cells(MAX_ROW, MAX_COL)).Interior.Color = RGB(77, 77, 77)
    Range(Cells(MIN_ROW, MIN_COL), Cells(MAX_ROW, MIN_COL)).Interior.Color = RGB(77, 77, 77)
    Range(Cells(MIN_ROW, MAX_COL), Cells(MAX_ROW, MAX_COL)).Interior.Color = RGB(77, 77, 77)
    
    Range(Cells(MIN_ROW + 6, MAX_COL + 2), Cells(MIN_ROW + 13, MAX_COL + 2)).Interior.Color = RGB(77, 77, 77)
    Range(Cells(MIN_ROW + 6, MAX_COL + 7), Cells(MIN_ROW + 13, MAX_COL + 7)).Interior.Color = RGB(77, 77, 77)
    Range(Cells(MIN_ROW + 6, MAX_COL + 2), Cells(MIN_ROW + 6, MAX_COL + 7)).Interior.Color = RGB(77, 77, 77)
    Range(Cells(MIN_ROW + 13, MAX_COL + 2), Cells(MIN_ROW + 13, MAX_COL + 7)).Interior.Color = RGB(77, 77, 77)
    
End Sub

Sub LowerShape()

    pivotRow = pivotRow + 1
    
    If CheckMove Then
    
        DrawShape
        
    Else
        pivotRow = pivotRow - 1
        PickShapeCoords
        DrawShape
        
        shapeJustPlaced = True
        ClearRow (currentRow1)
        ClearRow (currentRow2)
        ClearRow (currentRow3)
        ClearRow (currentRow4)
        
    End If
End Sub

Function CheckEnd() As Boolean

    pivotRow = pivotRow + 1
    
    If CheckMove Then
    
        CheckEnd = True
        
    Else
        CheckEnd = False
        
    End If
    
    pivotRow = pivotRow - 1
    
    PickShapeCoords
    
End Function


Function CheckMove() As Boolean

    EraseShape
    
    PickShapeCoords
    
    If CheckBounds And CheckForBlocks Then
        CheckMove = True
        
    Else
        CheckMove = False
        
    End If
    
End Function
Sub SetupGame()

    DrawBorders
    SpawnFirstShape
    
    shapeJustPlaced = False
    
    startTime = Timer
End Sub

Sub AddTime(timeNum As Double)

    If startTime < 1 Then
        
        startTime = startTime + timeNum
        
    End If

End Sub

Sub StartGame()
    
    SetupGame
    
    While (1)
        
        If shapeJustPlaced Then
            startTime = Timer
            SpawnShape
            shapeJustPlaced = False
        End If
        
        If GetAsyncKeyState(VK_LEFT) <> 0 Then
            ShiftShape (-1)
            Sleep 500
            AddTime (0.4)
        End If
        
        If GetAsyncKeyState(VK_RIGHT) <> 0 Then
            ShiftShape (1)
            Sleep 500
            AddTime (0.4)
        End If
        
        If GetAsyncKeyState(VK_UP) <> 0 Then
            RotateShape (1)
            Sleep 500
            AddTime (0.4)
        End If
        
        If GetAsyncKeyState(VK_DOWN) <> 0 Then
            RotateShape (-1)
            Sleep 500
            AddTime (0.4)
        End If
        
        If GetAsyncKeyState(VK_HOME) <> 0 Then
        
            While (CheckEnd)
            
                LowerShape
                
            Wend
            
            LowerShape
            
            Sleep 500
            
            startTime = startTime + 1000
            
        End If
        
        
        If (Timer - startTime >= 0.8) Then
            
            LowerShape
            
            startTime = Timer
            
        End If
        
        DoEvents
        
    
    Wend
    
End Sub
