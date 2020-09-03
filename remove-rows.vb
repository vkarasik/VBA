Dim rng As Range, value, i As Long                                                                                                                     

'2. Присваиваем имя "rng" диапазону, чтобы удобнее писать код.
Set rng = Range("A1:A1000")     
 
'3. Цикл по строкам диапазона снизу вверх.
For i = rng.Rows.Count To 1 Step -1

    '1) Копирование данных из ячейки в переменную, чтобы ускорить макрос,
    'чтобы обращаться к переменной, а не к объекту.
    value = rng.Cells(i, 1).value
   
    '2) Проверка, что находится в переменной "value".
    If (value = 0) Or (value = "CDL") Then
        '3) Удаление строки.
        rng.Rows(i).EntireRow.Delete
    End If

Next i