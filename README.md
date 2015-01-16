hearTogether
============

This is a android program to share the audio-data between two android equipment.
Option Explicit

Private 横軸最大値 As Integer, サブ区間数 As Integer, データ組数 As Integer, グラフ数 As Integer, シリーズ数 As Integer, checkedNum As Integer
Sub 一括処理()
    Dim sh As Worksheet
    
    Set sh = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    sh.Name = "BIN平均データ"
    
    With ActiveSheet.Buttons.Add(Range("B2").Left, _
                                 Range("B2").top, _
                                 Range("B2:C3").Width, _
                                 Range("B2:C3").Height)
        .OnAction = "initialSheetName"
        .Characters.Text = "スタート"
        
    End With
    
    With ActiveSheet.Buttons.Add(Range("B5").Left, _
                                 Range("B5").top, _
                                 Range("B5:C6").Width, _
                                 Range("B5:C6").Height)
        .OnAction = "calcBINAverage_2"
        .Characters.Text = "BIN平均"
        
    End With
    
    With ActiveSheet.Buttons.Add(Range("B8").Left, _
                                 Range("B8").top, _
                                 Range("B8:C9").Width, _
                                 Range("B8:C9").Height)
        .OnAction = "drawGraph"
        .Characters.Text = "グラフ作成"
        
    End With
    
    
End Sub

Private Sub initialSheetName()
    Dim m As Integer, n As Integer
    
    Dim sheetList() As String
    ReDim sheetList(Worksheets.Count - 1, 2)
    
    横軸最大値 = InputBox("横軸最大値を入力してください")
    サブ区間数 = InputBox("サブ区間数を入力してください")
    
    ActiveSheet.Cells(12, 2).Value = "横軸最大値"
    ActiveSheet.Cells(12, 3).Value = "サブ区間数"
    ActiveSheet.Cells(13, 2).Value = 横軸最大値
    ActiveSheet.Cells(13, 3).Value = サブ区間数
    
    ActiveSheet.Cells(15, 2).Value = "計算用シート選択"
    'MsgBox Worksheets.Count
    For m = 1 To Worksheets.Count - 1
        'sheetList(m, 1) = Worksheets(m).Index
        sheetList(m, 1) = m
        'MsgBox Str(m) + "," + Str(Worksheets(m).Index)
        sheetList(m, 2) = Worksheets(m).Name
        'MsgBox Worksheets(m).Name
    Next m
    For m = 1 To Worksheets.Count - 1
        With ActiveSheet.CheckBoxes.Add(Cells(15 + m, 2).Left, _
                                        Cells(15 + m, 2).top, _
                                        Cells(15 + m, 2).Width, _
                                        Cells(15 + m, 2).Height)
            '.Caption = Worksheets(m).Index
            .Caption = sheetList(m, 1)
            .Value = xlOff
        End With
        ActiveSheet.Cells(15 + m, 3).Value = sheetList(m, 2)
    Next m
    
    Erase sheetList
End Sub

Public Sub getCheckedNum()
    Dim checkedNum As Integer
    
    With ActiveSheet
    checkedNum = 1
    For n = 1 To checkNum
        If .CheckBoxes(n).Value = xlOn Then
            sheetIndex(checkedNum) = .CheckBoxes(n).Caption
            checkedNum = checkedNum + 1
        End If
    Next n
    getCheckedNum = checkedNum - 1
    End With
End Sub

Private Sub calcBINAverage()
    データ組数 = 7
    Dim sheetIndex() As Integer
    Dim l As Integer, n As Integer, checkNum As Integer
    Dim rowNum() As Integer
    
    Dim subLenth As Double
    Dim averageArray() As Double, datanumInSub() As Double
    
    
    Dim cell_2() As Double, cell_11() As Double, cell_12() As Double
    
    
    Dim i As Integer, m As Integer, m11 As Integer, m12 As Integer
    
    If サブ区間数 <> 0 Then
        subLenth = 横軸最大値 / サブ区間数
    End If
    
    checkNum = Worksheets.Count - 1
    ReDim sheetIndex(checkNum)
    
    With ActiveSheet
    checkedNum = 1
    For n = 1 To checkNum
        If .CheckBoxes(n).Value = xlOn Then
            sheetIndex(checkedNum) = .CheckBoxes(n).Caption
            checkedNum = checkedNum + 1
        End If
    Next n
    checkedNum = checkedNum - 1
    End With
    
    ReDim rowNum(checkedNum)
    'n = 1
    For n = 1 To checkedNum
        'MsgBox sheetIndex(n)
        rowNum(n) = Worksheets(sheetIndex(n)).UsedRange.Rows.Count
        'MsgBox Worksheets(sheetIndex(n)).Name
        'MsgBox rowNum(n)
    Next n
    Application.ScreenUpdating = False
    With ActiveSheet
    
    For n = 1 To checkedNum
        .Cells(1 + (n - 1) * (サブ区間数 + 4), 6).Value = Worksheets(sheetIndex(n)).Name
        For m = 1 To データ組数
            .Cells(2 + (n - 1) * (サブ区間数 + 4), 6 + (m - 1) * 4).Value = m & "号機"
            .Cells(3 + (n - 1) * (サブ区間数 + 4), 7 + (m - 1) * 4).Value = "FFT"
            .Cells(3 + (n - 1) * (サブ区間数 + 4), 8 + (m - 1) * 4).Value = "EnvFFt"
        Next m
    Next n
    
    End With
    
    For n = 1 To checkedNum
    For l = 1 To データ組数
        ReDim cell_2(rowNum(n), 0), cell_11(rowNum(n), 0), cell_12(rowNum(n), 0)
        ReDim averageArray(サブ区間数, 2), datanumInSub(サブ区間数, 2)
        With Worksheets(sheetIndex(n))
    
        For i = 2 To rowNum(n)
            If .Cells(i, 2).Value < 0 Then
                cell_2(i - 2, 0) = 0
            Else
                cell_2(i - 2, 0) = .Cells(i, 2).Value
            End If
            cell_11(i - 2, 0) = .Cells(i, 9 + (l - 1) * 3).Value
            cell_12(i - 2, 0) = .Cells(i, 10 + (l - 1) * 3).Value
        Next i
        
        For i = 0 To rowNum(n) - 2
            m = inWhichSubarea_2(cell_2(i, 0), subLenth)
            averageArray(m, 0) = averageArray(m, 0) + cell_2(i, 0)
            datanumInSub(m, 0) = datanumInSub(m, 0) + 1
            
            averageArray(m, 1) = averageArray(m, 1) + cell_11(i, 0)
            datanumInSub(m, 1) = datanumInSub(m, 1) + 1
            
            averageArray(m, 2) = averageArray(m, 2) + cell_12(i, 0)
            datanumInSub(m, 2) = datanumInSub(m, 2) + 1
        Next i
        
        For i = 1 To サブ区間数
            If datanumInSub(i, 0) = 0 Then
                averageArray(i, 0) = 0
            Else
                averageArray(i, 0) = averageArray(i, 0) / datanumInSub(i, 0)
            End If
            
            If datanumInSub(i, 1) = 0 Then
                averageArray(i, 1) = 0
            Else
                averageArray(i, 1) = averageArray(i, 1) / datanumInSub(i, 1)
            End If
            
            If datanumInSub(i, 2) = 0 Then
                averageArray(i, 2) = 0
            Else
                averageArray(i, 2) = averageArray(i, 2) / datanumInSub(i, 2)
            End If
            
            'MsgBox Str(averageArray(i, 0)) + ", " + Str(averageArray(i, 1)) + ", " + Str(averageArray(i, 2))
        Next i
    
        Erase cell_2
        Erase cell_11
        Erase cell_12
        
        End With
        
        With ActiveSheet
        
        'Application.ScreenUpdating = False
        
        Range(.Cells(4 + (n - 1) * (サブ区間数 + 4), 6 + (l - 1) * 4), _
              .Cells(24 + (n - 1) * (サブ区間数 + 4), 8 + (l - 1) * 4)) = averageArray

        'Application.ScreenUpdating = True
        
        Erase averageArray
        Erase datanumInSub
        
        End With
    Next l
    Next n
    Application.ScreenUpdating = True
End Sub

Private Sub calcBINAverage_2()
    
    データ組数 = 7
    Dim sheetIndex() As Integer
    Dim l As Integer, n As Integer, checkNum As Integer
    Dim rowNum() As Integer
    
    Dim subLenth As Double
    Dim averageArray() As Double, datanumInSub() As Double
    Dim temp As Double
    
    Dim cell_2() As Double, cell_11() As Double, cell_12() As Double
    Dim newSequ() As Integer
    Dim dataNumInSubarea As Integer, remainder As Integer
    
    Dim i As Integer, j As Integer, m As Integer, m11 As Integer, m12 As Integer
    Dim top As Integer, bot As Integer, swapd As Double, swapi As Integer
    
    'If サブ区間数 <> 0 Then
    '    subLenth = 横軸最大値 / サブ区間数
    'End If
    
    checkNum = Worksheets.Count - 1
    ReDim sheetIndex(checkNum)
    
    With ActiveSheet
    checkedNum = 1
    For n = 1 To checkNum
        If .CheckBoxes(n).Value = xlOn Then
            sheetIndex(checkedNum) = .CheckBoxes(n).Caption
            checkedNum = checkedNum + 1
        End If
    Next n
    checkedNum = checkedNum - 1
    End With
    
    ReDim rowNum(checkedNum)
    'n = 1
    For n = 1 To checkedNum
        'MsgBox sheetIndex(n)
        rowNum(n) = Worksheets(sheetIndex(n)).UsedRange.Rows.Count - 1
        'MsgBox Worksheets(sheetIndex(n)).Name
        'MsgBox rowNum(n)
    Next n
    Application.ScreenUpdating = False
    With ActiveSheet
    
    For n = 1 To checkedNum
        .Cells(1 + (n - 1) * (サブ区間数 + 4), 6).Value = Worksheets(sheetIndex(n)).Name
        For m = 1 To データ組数
            .Cells(2 + (n - 1) * (サブ区間数 + 4), 6 + (m - 1) * 4).Value = m & "号機"
            .Cells(3 + (n - 1) * (サブ区間数 + 4), 7 + (m - 1) * 4).Value = "FFT"
            .Cells(3 + (n - 1) * (サブ区間数 + 4), 8 + (m - 1) * 4).Value = "EnvFFt"
        Next m
    Next n
    
    End With
    
    For n = 1 To checkedNum
    For l = 1 To データ組数
        'ReDim dataNumInSubarea(checkedNum)
        'ReDim remainder(checkedNum)
        ReDim newSequ(rowNum(n))
        dataNumInSubarea = rowNum(n) \ 21
        remainder = rowNum(n) Mod 21
        'MsgBox dataNumInSubarea(n) & "-" & remainder(n)
        
        ReDim cell_2(rowNum(n), 0), cell_11(rowNum(n), 0), cell_12(rowNum(n), 0)
        ReDim averageArray(サブ区間数 + 1, 2), datanumInSub(サブ区間数 + 1, 2)
        With Worksheets(sheetIndex(n))
    
        For i = 1 To rowNum(n)
            If .Cells(i + 1, 2).Value < 0 Then
                cell_2(i - 1, 0) = 0
            Else
                cell_2(i - 1, 0) = .Cells(i + 1, 2).Value
            End If
            cell_11(i - 1, 0) = .Cells(i + 1, 9 + (l - 1) * 3).Value
            cell_12(i - 1, 0) = .Cells(i + 1, 10 + (l - 1) * 3).Value
        Next i
        
        For i = 1 To rowNum(n)
            newSequ(i) = i
        Next i
        top = 0
        bot = rowNum(n) - 1
        Do While 1
            Dim last_swap_index As Integer
            last_swap_index = top
            For i = top To bot - 1
                If cell_2(i, 0) > cell_2(i + 1, 0) Then
                    swapd = cell_2(i + 1, 0)
                    cell_2(i + 1, 0) = cell_2(i, 0)
                    cell_2(i, 0) = swapd
                    swapi = newSequ(i + 1)
                    newSequ(i + 1) = newSequ(i)
                    newSequ(i) = swapi
                    last_swap_index = i
                End If
            Next i
            bot = last_swap_index
            If top = bot Then
                Exit Do
            End If
            
            last_swap_index = bot
            For j = bot To top + 1 Step -1
                If cell_2(j, 0) < cell_2(j - 1, 0) Then
                    swapd = cell_2(j - 1, 0)
                    cell_2(j - 1, 0) = cell_2(j, 0)
                    cell_2(j, 0) = swapd
                    swapi = newSequ(j - 1)
                    newSequ(j - 1) = newSequ(j)
                    newSequ(j) = swapi
                    last_swap_index = j
                End If
            Next j
            top = last_swap_index
            If top = bot Then
                Exit Do
            End If
        Loop
        
        m = 0
        
        For i = 0 To rowNum(n) - 1
            'm = inWhichSubarea_2(cell_2(i, 0), subLenth)
            If i Mod dataNumInSubarea <> 0 Then
                averageArray(m, 0) = averageArray(m, 0) + cell_2(i, 0)
                datanumInSub(m, 0) = datanumInSub(m, 0) + 1
                
                averageArray(m, 1) = averageArray(m, 1) + cell_11(newSequ(i), 0)
                datanumInSub(m, 1) = datanumInSub(m, 1) + 1
            
                averageArray(m, 2) = averageArray(m, 2) + cell_12(newSequ(i), 0)
                datanumInSub(m, 2) = datanumInSub(m, 2) + 1
            Else
                averageArray(m, 0) = averageArray(m, 0) + cell_2(i, 0)
                datanumInSub(m, 0) = datanumInSub(m, 0) + 1
                
                averageArray(m, 1) = averageArray(m, 1) + cell_11(newSequ(i), 0)
                datanumInSub(m, 1) = datanumInSub(m, 1) + 1
            
                averageArray(m, 2) = averageArray(m, 2) + cell_12(newSequ(i), 0)
                datanumInSub(m, 2) = datanumInSub(m, 2) + 1
                m = m + 1
                'MsgBox "i:" & i & " m:" & m
            End If
        Next i
        
        For i = 1 To サブ区間数
            If datanumInSub(i, 0) = 0 Then
                averageArray(i, 0) = 0
            Else
                averageArray(i, 0) = averageArray(i, 0) / datanumInSub(i, 0)
            End If
            
            If datanumInSub(i, 1) = 0 Then
                averageArray(i, 1) = 0
            Else
                averageArray(i, 1) = averageArray(i, 1) / datanumInSub(i, 1)
            End If
            
            If datanumInSub(i, 2) = 0 Then
                averageArray(i, 2) = 0
            Else
                averageArray(i, 2) = averageArray(i, 2) / datanumInSub(i, 2)
            End If
            
            'MsgBox Str(averageArray(i, 0)) + ", " + Str(averageArray(i, 1)) + ", " + Str(averageArray(i, 2))
        Next i
    
        Erase cell_2
        Erase cell_11
        Erase cell_12
        
        End With
        
        With ActiveSheet
        
        'Application.ScreenUpdating = False
        
        Range(.Cells(4 + (n - 1) * (サブ区間数 + 4), 6 + (l - 1) * 4), _
              .Cells(24 + (n - 1) * (サブ区間数 + 4), 8 + (l - 1) * 4)) = averageArray

        'Application.ScreenUpdating = True
        
        Erase averageArray
        Erase datanumInSub
        
        End With
    Next l
    Next n
    Application.ScreenUpdating = True

End Sub

Private Sub drawGraph()
    グラフ数 = 7
    シリーズ数 = 2
    'ReDim title(checkedNum, グラフ数), series(シリーズ数), graphData(グラフ数, サブ区間数, シリーズ数 + 1)
    Dim graphData() As Double
    ReDim graphData(グラフ数, サブ区間数, シリーズ数 + 1)
    
    'series(0) = "FFT"
    'series(1) = "EnvFFt"
    
    Dim i As Integer, j As Integer, m As Integer, n As Integer
    
    Dim y_offset As Integer, x_offset As Integer
    
    Dim co_y As Integer, co_x As Integer, c_y As Integer, c_x As Integer
    Dim co_height As Integer, co_width As Integer, c_height As Integer, c_width As Integer
    
    'Dim co_name As String, c_name As String
    Dim temp As Range, target As Range
    
    y_offset = 4
    x_offset = 25
    
    'co_y = 12
    'co_x = 6
    c_y = 50
    c_x = 200
    'co_height = 120
    'co_width = 500
    c_height = 200
    c_width = 300
    
    If ActiveSheet.ChartObjects.Count > 1 Then
    For i = ActiveSheet.ChartObjects.Count To 1 Step -1
        ActiveSheet.ChartObjects(i).Delete
    Next i
    End If
    
    Application.ScreenUpdating = False
    For i = 1 To checkedNum
    
        'Dim chartObj As ChartObject
        'Set chartObj = ActiveSheet.ChartObjects.Add(co_x, co_y * i, co_width, co_height)
        'chartObj.Name = Cells(i, 6).Value
        'chartObj.Chart.ChartType = xlXYScatter
    
        For j = 1 To グラフ数
            
            ActiveSheet.Shapes.AddChart(xlXYScatterSmoothNoMarkers, c_x + (j - 1) * (c_width + 10), c_y + (i - 1) * (c_height + 10), c_width, c_height).Select
            ActiveChart.HasTitle = True
            ActiveChart.ChartTitle.Text = Cells(1 + (i - 1) * x_offset, 6).Value & "-" & Cells(2 + (i - 1) * x_offset, 6 + (j - 1) * y_offset).Value
            
            'If ActiveChart.SeriesCollection.Count > 1 Then
            'For m = ActiveChart.SeriesCollection.Count To 1 Step -1
            '    ActiveChart.SeriesCollection(i).Delete
            'Next m
            'End If
            
            'For n = 1 To サブ区間数
            '    graphData(j, n, 0) = Cells(4 + (i - 1) * x_offset + (n - 1), 6 + (j - 1) * y_offset).Value
            '    graphData(j, n, 1) = Cells(4 + (i - 1) * x_offset + (n - 1), 7 + (j - 1) * y_offset).Value
            '    graphData(j, n, 2) = Cells(4 + (i - 1) * x_offset + (n - 1), 8 + (j - 1) * y_offset).Value
            'Next n
            With ActiveChart
            
            With .Axes(xlCategory)
                .MaximumScaleIsAuto = "False"
                .MaximumScale = 2100
            End With
            
            With .Axes(xlValue)
                .MaximumScaleIsAuto = "False"
                If j <= 4 Then
                    .MaximumScale = 0.2
                Else
                    .MaximumScale = 1.2
                End If
            End With
            
            With .SeriesCollection.NewSeries
            
            .Name = "FFT"
            Set temp = Nothing
            Set target = Nothing
            Range(Cells(4 + (i - 1) * x_offset, 6 + (j - 1) * y_offset), Cells(24 + (i - 1) * x_offset, 6 + (j - 1) * y_offset)).Select
            For Each temp In Selection
                If temp.Value <> 0 Then
                    If target Is Nothing Then
                        Set target = temp
                    Else
                        Set target = Union(target, temp)
                    End If
                End If
            Next temp
            If Not target Is Nothing Then
                target.Select
            End If
            .XValues = target
            Set temp = Nothing
            Set target = Nothing
            Range(Cells(4 + (i - 1) * x_offset, 7 + (j - 1) * y_offset), Cells(24 + (i - 1) * x_offset, 7 + (j - 1) * y_offset)).Select
            For Each temp In Selection
                If temp.Value <> 0 Then
                    If target Is Nothing Then
                        Set target = temp
                    Else
                        Set target = Union(target, temp)
                    End If
                End If
            Next temp
            If Not target Is Nothing Then
                target.Select
            End If
            .Values = target
            Set temp = Nothing
            Set target = Nothing
            .MarkerSize = 3
            
            End With
            
            With .SeriesCollection.NewSeries
            
            .Name = "EnvFFt"
            Set temp = Nothing
            Set target = Nothing
            Range(Cells(4 + (i - 1) * x_offset, 6 + (j - 1) * y_offset), Cells(24 + (i - 1) * x_offset, 6 + (j - 1) * y_offset)).Select
            For Each temp In Selection
                If temp.Value <> 0 Then
                    If target Is Nothing Then
                        Set target = temp
                    Else
                        Set target = Union(target, temp)
                    End If
                End If
            Next temp
            If Not target Is Nothing Then
                target.Select
            End If
            .XValues = target
            Set temp = Nothing
            Set target = Nothing
            Range(Cells(4 + (i - 1) * x_offset, 8 + (j - 1) * y_offset), Cells(24 + (i - 1) * x_offset, 8 + (j - 1) * y_offset)).Select
            For Each temp In Selection
                If temp.Value <> 0 Then
                    If target Is Nothing Then
                        Set target = temp
                    Else
                        Set target = Union(target, temp)
                    End If
                End If
            Next temp
            If Not target Is Nothing Then
                target.Select
            End If
            .Values = target
            Set temp = Nothing
            Set target = Nothing
            
            .MarkerSize = 3
            
            End With
            
            End With
            
            ActiveSheet.Cells(1, 1).Select
        
        Next j
    Next i
    Application.ScreenUpdating = True
    MsgBox "処理完了"
End Sub
