Sub Monthly_Report_04_21()

'Создание нового листа с конечной таблицей за месяц

Sheets.Add.Name = "Отчет по трудозатратам за 04.21"

'Добавление столбцов с заголовками в конечной таблицы

Worksheets("Отчет по трудозатратам за 04.21").Activate

Dim Col1 As Range: Set Col1 = Application.Range("A:A")
Col1.Insert Shift:=xlShiftToRight

[A1] = "ФИО"

Dim Col2 As Range: Set Col2 = Application.Range("B:B")
Col2.Insert Shift:=xlShiftToRight

[B1] = "Месячная норма"

Dim Col3 As Range: Set Col3 = Application.Range("C:C")
Col3.Insert Shift:=xlShiftToRight

[C1] = "Остаток за прошлый месяц"

Dim Col4 As Range: Set Col4 = Application.Range("D:D")
Col4.Insert Shift:=xlShiftToRight

[D1] = "Фактически отработано за текущий месяц"

Dim Col5 As Range: Set Col5 = Application.Range("E:E")
Col5.Insert Shift:=xlShiftToRight

[E1] = "Отгулы"

Dim Col6 As Range: Set Col6 = Application.Range("F:F")
Col6.Insert Shift:=xlShiftToRight

[F1] = "Итог за текущий месяц"

'Кастомизация таблицы

Range("A1:F1").ColumnWidth = 40

Range("A1:F1").Interior.Color = RGB(60, 80, 111)

Range("A1:F1").Font.ColorIndex = 2

Range("A1:F1").VerticalAlignment = xlCenter

'Ввод месячного часового норматива за Апрель

[B2:B500] = "175"

'Удаление ненужных столбцов и строк в таблице по трудозатратам (без отгулов)

Worksheets("Summary Report").Activate

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

For i = 10000 To 1 Step -1
    If Cells(i, 1) = "" Then
        Cells(i, 1).EntireRow.Delete
    End If
Next i

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

Dim R As Range
    Set R = [A1].CurrentRegion
    R.Rows(R.Rows.Count).Delete

Columns(2).Delete
Columns(4).Delete
Columns(2).Delete

'Удаление ненужных столбцов и строк в таблице по трудозатратам (с отгулами)

Worksheets("Summary Report (2)").Activate

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

For i = 1000 To 1 Step -1
    If Cells(i, 1) = "" Then
        Cells(i, 1).EntireRow.Delete
    End If
Next i

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

Dim R1 As Range
    Set R1 = [A1].CurrentRegion
    R1.Rows(R1.Rows.Count).Delete

Columns(2).Delete
Columns(4).Delete
Columns(2).Delete

'Копирование таблицы со всеми сотрудниками для дальнейшей обработки

Worksheets("All").Activate

Range("A1:A500").Select
Selection.Copy
Sheets("Отчет по трудозатратам за 04.21").Select
Range("W1:W500").Select
Selection.Insert Shift:=xlDown

'Удаление форматирования во всем файле

ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

'Копирование таблицы с сотрудниками, которые заполнили Clockify за прошлый месяц для дальнейшей обработки

Worksheets("Summary Report").Activate

Range("A:B").Select
Selection.Copy
Sheets("Отчет по трудозатратам за 04.21").Select
Range("T:U").Select
Selection.Insert Shift:=xlDown

'Копирование таблицы с трудозатратами за прошлый месяц

Worksheets("Отчет по трудозатратам за 03.21").Activate

Range("A:F").Select
Selection.Copy
Sheets("Отчет по трудозатратам за 04.21").Select
Range("J:O").Select
Selection.Insert Shift:=xlDown

'Повторное удаление форматирования во всем файле (необходимо для из-за формул в прошлой таблице)

ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

'Копирование всех сотрудников

Range("Z2:Z500").Select
Selection.Copy
Sheets("Отчет по трудозатратам за 04.21").Select
Range("A2:A500").Select
Selection.Insert Shift:=xlDown

'Цикл сверки сотрцдников, которых нет в таблице из Clockify

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

For i = 5000 To 1 Step -1
    If Cells(i, 1) <> Cells(i, 10) Then
        Cells(i, 10).Select
        Selection.Copy
        Cells(i, 1).Select
        Selection.Insert Shift:= xlDown
    End If
Next i

For i = 5000 To 1 Step -1
    If Cells(i, 1) = "" Then
        Cells(i, 1).EntireRow.Delete
    End If
Next i

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

'Удаление дубликатов

Range("A1").RemoveDuplicates 1

'Ввод формулы для подсчета часов

Range("F2").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-4]-RC[-1]+RC[-3]"
Range("F2").Select
Selection.AutoFill Destination:=Range("F2:F500"), Type:=xlFillDefault
Range("F2:F500").Select

'Удаление ненужных столбцов

Columns(10).Delete
Columns(11).Delete
Columns(12).Delete
Columns(13).Delete
Columns(14).Delete
Columns(15).Delete
Columns(16).Delete
Columns(17).Delete
Columns(18).Delete
Columns(19).Delete
Columns(20).Delete
Columns(21).Delete
Columns(22).Delete
Columns(23).Delete
Columns(24).Delete
Columns(25).Delete
Columns(26).Delete
Columns(27).Delete
Columns(28).Delete
Columns(29).Delete
Columns(30).Delete
Columns(31).Delete
Columns(10).Delete
Columns(11).Delete
Columns(12).Delete
Columns(13).Delete
Columns(14).Delete
Columns(15).Delete
Columns(16).Delete
Columns(17).Delete
Columns(18).Delete
Columns(19).Delete
Columns(20).Delete
Columns(21).Delete
Columns(22).Delete
Columns(23).Delete
Columns(24).Delete
Columns(25).Delete
Columns(26).Delete
Columns(27).Delete
Columns(28).Delete
Columns(29).Delete
Columns(30).Delete
Columns(31).Delete
Columns(10).Delete

'Форматирования всей таблици

Range("A:G").Font.Name = "Arial"

Range("A:G").Font.Bold = False

End Sub
