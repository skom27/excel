Attribute VB_Name = "Module1"
Option Explicit

' Zmienne globalne do przechowywania świąt i informacji śledzenia
Private holidays As Collection
Private monthDays As Integer
Private currentYear As Integer
Private currentMonth As Integer
Private employeeCount As Integer
Private Const SHEET_PASSWORD As String = "654321" ' Hasło do arkusza "Harmonogram Pracy"

Sub InitializeWorkSchedule()
    ' Main initialization procedure triggered by button
    Dim wsHarmonogram As Worksheet
    Dim wsComboBox As Worksheet
    Dim wsZmiany As Worksheet
    Dim wsPracownicy As Worksheet
    Dim dateA6 As Variant
    Dim selectedMonth As Integer
    Dim selectedYear As Integer
    Dim response As VbMsgBoxResult
    Dim dzial As String
    Dim newFileName As String

    ' Sprawdź czy arkusze istnieją
    On Error Resume Next
    Set wsHarmonogram = ThisWorkbook.Sheets("Harmonogram Pracy")
    Set wsComboBox = ThisWorkbook.Sheets("ComboBox")
    Set wsZmiany = ThisWorkbook.Sheets("Zmiany")
    Set wsPracownicy = ThisWorkbook.Sheets("Pracownicy")
    Dim wsKalendarz As Worksheet
    Set wsKalendarz = ThisWorkbook.Sheets("Kalendarz")
    On Error GoTo 0

    If wsHarmonogram Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'Harmonogram Pracy'!", vbExclamation
        Exit Sub
    End If

    If wsComboBox Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'ComboBox'!", vbExclamation
        Exit Sub
    End If

    If wsZmiany Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'Zmiany'!", vbExclamation
        Exit Sub
    End If

    If wsPracownicy Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'Pracownicy'!", vbExclamation
        Exit Sub
    End If

    ' Pobierz miesiąc i rok z arkusza ComboBox
    selectedMonth = Val(wsComboBox.Range("B2").Value)
    selectedYear = Val(wsComboBox.Range("D2").Value)

    ' Pobierz datę z komórki A6 arkusza Kalendarz (jeśli istnieje)
    If Not wsKalendarz Is Nothing Then
        dateA6 = wsKalendarz.Range("A6").Value
    Else
        dateA6 = Empty ' Jeśli arkusz Kalendarz nie istnieje, ustaw pustą wartość
    End If

    ' Sprawdź czy komórka A6 jest pusta
    If IsEmpty(dateA6) Or Not IsDate(dateA6) Then
        ' Przypadek 1: Komórka jest pusta - wykonaj skrypt
        MsgBox "Tworzenie nowego harmonogramu na " & MonthName(selectedMonth) & " " & selectedYear & ".", vbInformation, "Nowy harmonogram"
        
        Call ClearKalendarzSheet
        Call PrepareWorkScheduleTemplate
        Call PopulateMonthDates

        ' Nadaj nazwę plikowi
        dzial = wsPracownicy.Range("A1").Value
        newFileName = dzial & "-" & selectedYear & "-" & selectedMonth & ".xlsm"
        Call RenameWorkbook(newFileName)

        ' Tylko jeśli nie wyszliśmy wcześniej z procedury (gdy użytkownik wybrał "Nie")
        ' Wyczyść komórki w arkuszu "Harmonogram Pracy" od wiersza 4 w dół
        Call ClearHarmonogramPracyRows(wsHarmonogram)
        
        ' MsgBox "Harmonogram został przygotowany!", vbInformation
    Else
        ' Komórka A6 zawiera datę - sprawdź miesiąc i rok
        Dim dateMonth As Integer
        Dim dateYear As Integer

        dateMonth = Month(dateA6)
        dateYear = Year(dateA6)

        If dateMonth = selectedMonth And dateYear = selectedYear Then
            ' Przypadek 2: Data A6 jest w tym samym miesiącu i roku
            response = MsgBox("Wybrałeś ten sam miesiąc (" & MonthName(selectedMonth) & " " & selectedYear & "), który jest już utworzony. " & _
                             "Czy chcesz utworzyć harmonogram na nowo?" & vbCrLf & vbCrLf & _
                             "Tak - usuń bieżący harmonogram i wygeneruj nowy" & vbCrLf & _
                             "Nie - zachowaj bieżący harmonogram", _
                             vbYesNo + vbQuestion, "Potwierdzenie")

            If response = vbYes Then
                ' Wybrano Tak - usuń bieżący harmonogram i wygeneruj nowy
                Call ClearKalendarzSheet
                Call PrepareWorkScheduleTemplate
                Call PopulateMonthDates

                ' Nadaj nazwę plikowi
                dzial = wsPracownicy.Range("A1").Value
                newFileName = dzial & "-" & selectedYear & "-" & selectedMonth & ".xlsm"
                Call RenameWorkbook(newFileName)

                ' MsgBox "Harmonogram został przygotowany!", vbInformation
            Else
                ' Wybrano Nie - wyjdź z procedury
                Exit Sub
            End If
        Else
            ' Przypadek 3: Data A6 jest w innym miesiącu lub roku
            response = MsgBox("Aktualnie masz utworzony harmonogram na " & MonthName(dateMonth) & " " & dateYear & "." & vbCrLf & _
                             "Wybrałeś nowy miesiąc: " & MonthName(selectedMonth) & " " & selectedYear & "." & vbCrLf & vbCrLf & _
                             "Czy chcesz utworzyć nowy harmonogram?" & vbCrLf & vbCrLf & _
                             "Tak - usuń bieżący harmonogram i wygeneruj nowy" & vbCrLf & _
                             "Nie - zachowaj bieżący harmonogram", _
                             vbYesNo + vbQuestion, "Potwierdzenie")

            If response = vbYes Then
                ' Wybrano Tak - usuń bieżący harmonogram i wygeneruj nowy
                Call ClearKalendarzSheet
                Call PrepareWorkScheduleTemplate
                Call PopulateMonthDates

                ' Nadaj nazwę plikowi
                dzial = wsPracownicy.Range("A1").Value
                newFileName = dzial & "-" & selectedYear & "-" & selectedMonth & ".xlsm"
                Call RenameWorkbook(newFileName)

                ' MsgBox "Harmonogram został przygotowany!", vbInformation
            Else
                ' Wybrano Nie - wyjdź z procedury
                Exit Sub
            End If
        End If
    End If

    ' Wyczyść komórki w arkuszu "Harmonogram Pracy" od wiersza 4 w dół
    Call ClearHarmonogramPracyRows(wsHarmonogram)

    ' Kopiuj dane z arkusza Kalendarz do arkusza Harmonogram Pracy
    Call CopyKalendarzToHarmonogramPracy
End Sub

Sub CreateInitializationButton()
    Dim ws As Worksheet
    Dim btn As Button
    
    ' Make sure the Harmonogram Pracy worksheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Harmonogram Pracy")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'Harmonogram Pracy'!", vbExclamation
        Exit Sub
    End If
    
    ' Delete existing button if it exists
    On Error Resume Next
    ' ws.Buttons.Delete
    On Error GoTo 0
    
End Sub

Sub PrepareWorkScheduleTemplate()
    ' Main subroutine to set up the entire template
    Dim wsTarget As Worksheet
    Dim wsPracownicy As Worksheet
    Dim wsComboBox As Worksheet
    Dim i As Integer
    Dim lastRowPracownicy As Long
    
    Application.ScreenUpdating = False
    
    ' Check if necessary worksheets exist
    On Error Resume Next
    Set wsPracownicy = ThisWorkbook.Sheets("Pracownicy")
    Set wsComboBox = ThisWorkbook.Sheets("ComboBox")
    Set wsTarget = ThisWorkbook.Sheets("Kalendarz")
    On Error GoTo 0
    
    If wsPracownicy Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'Pracownicy'!", vbExclamation
        Exit Sub
    End If
    
    If wsComboBox Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'ComboBox'!", vbExclamation
        Exit Sub
    End If
    
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Worksheets.Add
        wsTarget.Name = "Kalendarz"
    End If
    
    ' Get input values from Pracownicy sheet
    Dim dzial As String
    dzial = wsPracownicy.Range("A2").Value
    
    ' Get the number of employees from the Pracownicy sheet
    lastRowPracownicy = wsPracownicy.Cells(Rows.Count, "B").End(xlUp).row
    employeeCount = lastRowPracownicy ' Poprawione: uwzględniamy B1, więc nie odejmujemy 1
    
    ' Get month and year from ComboBox sheet
    currentMonth = Val(wsComboBox.Range("B2").Value)
    currentYear = Val(wsComboBox.Range("D2").Value)
    
    ' Validate inputs
    If Not ValidateInputs() Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    With wsTarget
        ' Ustaw szerokości kolumn
        .Columns("A").ColumnWidth = 14
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 16
    
        ' Create header row dynamically based on employee count
        .Range("A1").Value = "Data"
        .Range("B1").Value = "Dzień tyg."
        .Range("C1").Value = "Rodz. dnia"
        
        ' Format headers - center align
        .Range("A1:C1").HorizontalAlignment = xlCenter
        
        ' Dynamically create headers for each employee
        For i = 1 To employeeCount
            ' Pobierz nazwę pracownika z arkusza "Pracownicy" z kolumny B
            .Cells(1, 3 + (i * 2) - 1).Value = wsPracownicy.Cells(i, "B").Value
            .Cells(1, 3 + (i * 2)).Value = "Godz."

            ' Format employee headers - center align
            .Cells(1, 3 + (i * 2) - 1).HorizontalAlignment = xlCenter
            .Cells(1, 3 + (i * 2)).HorizontalAlignment = xlCenter
        Next i
        
        ' Nadgodziny Row
        .Range("A2:C2").MergeCells = True
        .Range("A2").Value = "Nadgodziny z poprzed. m-ca"
        .Range("A2").HorizontalAlignment = xlRight
        
        ' Overtime Hours Input Cells - dynamically create
        For i = 1 To employeeCount
            .Cells(2, 3 + (i * 2)).Value = 0
            .Cells(2, 3 + (i * 2)).HorizontalAlignment = xlCenter
            .Cells(2, 3 + (i * 2)).NumberFormat = "0"
            
            ' Dodaj walidację danych dla nadgodzin - tylko wartości liczbowe od -64 do +64
            With .Cells(2, 3 + (i * 2)).Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:="-64", Formula2:="64"
                .IgnoreBlank = True
                .InputTitle = "Wprowadź nadgodziny"
                .ErrorTitle = "Nieprawidłowa wartość"
                .InputMessage = "Wprowadź liczbę całkowitą od -64 do +64."
                .ErrorMessage = "Wartość musi być liczbą całkowitą z zakresu od -64 do +64."
            End With
        Next i
        
        ' Usuń ramki z całego arkusza
        .Cells.Borders.LineStyle = xlNone
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub PopulateMonthDates()
    ' Validate Month and Year Input
    If Not ValidateInputs() Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Initialize holidays collection
    Set holidays = New Collection
    
    ' Find all holidays for the given month and year
    Call FindHolidays
    
    Dim wsTarget As Worksheet
    Dim wsZmiany As Worksheet ' Dodano arkusz "Zmiany"
    Dim currentDate As Date
    Dim row As Integer
    Dim dayType As String
    Dim lastRow As Long
    Dim holidayCount As Integer
    Dim dayShift As Integer ' Zmiana dla danego dnia
    Dim i As Integer
    
    ' Get target worksheet
    Set wsTarget = ThisWorkbook.Sheets("Kalendarz")
    Set wsZmiany = ThisWorkbook.Sheets("Zmiany") ' Ustaw arkusz "Zmiany"
    
    currentDate = DateSerial(currentYear, currentMonth, 1)
    row = 3
    holidayCount = 0
    dayShift = 1 ' Początkowa zmiana dla pierwszego dnia
    
    ' Przypisz stałe zmiany dla pracowników (R1, R2, R3)
    Dim employeeShifts() As String
    ReDim employeeShifts(1 To employeeCount)
    
    For i = 1 To employeeCount
        employeeShifts(i) = "R" & ((i - 1) Mod 3 + 1) ' R1, R2, R3, R1, R2, ...
    Next i
    
    ' Populate dates until end of month
    Do While Month(currentDate) = currentMonth
        With wsTarget
            ' Format date consistently as DD.MM.YYYY
            .Cells(row, 1).Value = Format(currentDate, "DD.MM.YYYY")
            .Cells(row, 1).NumberFormat = "@"
            .Cells(row, 1).HorizontalAlignment = xlCenter
            
            .Cells(row, 2).Value = WeekdayName(Weekday(currentDate), False, 1)
            .Cells(row, 2).HorizontalAlignment = xlCenter
            
            ' Determine day type
            dayType = DetermineWorkDayType(currentDate)
            .Cells(row, 3).Value = dayType
            .Cells(row, 3).HorizontalAlignment = xlCenter
            
            ' Wypełnianie kolumn pracowników
            For i = 1 To employeeCount
                Dim zmiana As String
                Dim godziny As Variant
                
                Select Case dayType
                    Case "roboczy"
                        ' Każdy pracownik ma stałą zmianę (R1, R2 lub R3)
                        zmiana = employeeShifts(i)
                    Case "sobota"
                        zmiana = "SO" & IIf(i Mod 2 = 1, 1, 2) ' SO1, SO2, SO1, SO2
                    Case "niedziela"
                        zmiana = "NIED"
                    Case Else
                        zmiana = "DW"
                End Select
                
                .Cells(row, 3 + (i * 2) - 1).Value = zmiana
                .Cells(row, 3 + (i * 2) - 1).HorizontalAlignment = xlCenter
                
                ' Pobierz liczbę godzin z arkusza "Zmiany"
                godziny = GetGodziny(wsZmiany, zmiana)
                
                ' Wpisz liczbę godzin do komórki "Godz."
                .Cells(row, 3 + (i * 2)).Value = godziny
                .Cells(row, 3 + (i * 2)).HorizontalAlignment = xlCenter
            Next i
            
            ' Color coding
            Select Case dayType
                Case "roboczy"
                    .Cells(row, 3).Font.Color = vbBlack
                Case "sobota"
                    .Cells(row, 3).Font.Color = RGB(0, 0, 128) ' Dark Blue
                Case "niedziela"
                    .Cells(row, 3).Font.Color = RGB(0, 0, 128) ' Dark Blue
                Case Else
                    If InStr(dayType, "święto") > 0 Then
                        .Cells(row, 3).Font.Color = vbRed
                    End If
            End Select
        End With
        
        ' Przejdź do następnego dnia
        currentDate = DateAdd("d", 1, currentDate)
        row = row + 1
        
        ' Zwiększ numer zmiany dla następnego dnia (tylko dla dni roboczych)
        If dayType = "roboczy" Then
            dayShift = dayShift + 1
            If dayShift > 3 Then dayShift = 1
        End If
    Loop
    
    ' Calculate total work days and work hours
    Call CalculateTotalWorkDays(holidayCount)
    
    ' Apply consistent formatting to the entire table
    lastRow = wsTarget.Cells(Rows.Count, 1).End(xlUp).row
    
    ' Sprawdź czy wszystkie wymagane zmiany są obsadzone dla każdego dnia
    Dim wsHarmonogram As Worksheet
    On Error Resume Next
    Set wsHarmonogram = ThisWorkbook.Sheets("Harmonogram Pracy")
    On Error GoTo 0
    
    If Not wsHarmonogram Is Nothing Then
        ' Oblicz liczbę dni w miesiącu
        Dim daysInMonth As Long
        daysInMonth = Day(DateSerial(currentYear, currentMonth + 1, 0))
        
        Call CheckRequiredShifts(wsHarmonogram, daysInMonth, employeeCount)
    End If
    
    Application.ScreenUpdating = True
End Sub

Function GetGodziny(wsZmiany As Worksheet, zmiana As String) As Variant
    ' Funkcja pomocnicza do pobierania liczby godzin z arkusza "Zmiany"
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = wsZmiany.Cells(Rows.Count, 1).End(xlUp).row
    
    For i = 1 To lastRow
        If wsZmiany.Cells(i, 1).Value = zmiana Then
            GetGodziny = wsZmiany.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    ' Jeśli nie znaleziono zmiany, zwróć 0
    GetGodziny = 0
End Function

Function DetermineWorkDayType(workDate As Date) As String
    Dim dayOfWeek As Integer
    
    dayOfWeek = Weekday(workDate)
    
    ' Check if it's a holiday
    If isHoliday(workDate) Then
        Select Case dayOfWeek
            Case 1: DetermineWorkDayType = "niedziela | święto"
            Case 7: DetermineWorkDayType = "sobota | święto"
            Case 2 To 6: DetermineWorkDayType = "święto"
        End Select
        Exit Function
    End If
    
    ' Regular day type
    Select Case dayOfWeek
        Case 1: DetermineWorkDayType = "niedziela"
        Case 2 To 6: DetermineWorkDayType = "roboczy"
        Case 7: DetermineWorkDayType = "sobota"
    End Select
End Function

Sub CalculateTotalWorkDays(holidayCount As Integer)
    Dim wsTarget As Worksheet
    Dim wsComboBox As Worksheet
    Dim lastRow As Long
    Dim workDaysCount As Integer
    Dim saturdayHolidayCount As Integer
    Dim workHours As Integer
    Dim i As Long, j As Long
    Dim dayType As String
    Dim employeeHours() As Double

    Set wsTarget = ThisWorkbook.Sheets("Kalendarz")

    ' Sprawdź czy arkusz ComboBox istnieje
    On Error Resume Next
    Set wsComboBox = ThisWorkbook.Sheets("ComboBox")
    On Error GoTo 0

    lastRow = wsTarget.Cells(Rows.Count, 3).End(xlUp).row
    workDaysCount = 0

    ' Inicjalizuj tablicę do przechowywania godzin dla każdego pracownika
    ReDim employeeHours(1 To employeeCount)

    ' Najpierw dodaj nadgodziny z wiersza 2 dla każdego pracownika
    For j = 1 To employeeCount
        employeeHours(j) = Val(wsTarget.Cells(2, 3 + (j * 2)).Value)
    Next j

    ' Zlicz dni robocze i sumuj godziny dla każdego pracownika
    For i = 3 To lastRow
        dayType = wsTarget.Cells(i, 3).Value

        ' Pobierz datę z komórki
        Dim currentDate As Date
        If IsDate(wsTarget.Cells(i, 1).Value) Then
            currentDate = CDate(wsTarget.Cells(i, 1).Value)
            
            ' Zlicz dni robocze (od poniedziałku do piątku)
            If Weekday(currentDate) >= 2 And Weekday(currentDate) <= 6 Then
                ' Sprawdź czy to nie jest święto
                If InStr(dayType, "święto") = 0 Then
                    workDaysCount = workDaysCount + 1
                End If
            End If
            
            ' Zlicz święta przypadające w soboty
            If Weekday(currentDate) = 7 And InStr(dayType, "święto") > 0 Then
                saturdayHolidayCount = saturdayHolidayCount + 1
            End If
        End If
        ' Sumuj godziny dla każdego pracownika
        For j = 1 To employeeCount
            employeeHours(j) = employeeHours(j) + Val(wsTarget.Cells(i, 3 + (j * 2)).Value)
        Next j
    Next i
    
    ' Odejmij jeden dzień za każde święto przypadające w sobotę
    If saturdayHolidayCount > 0 Then
        workDaysCount = workDaysCount - 1
    End If
    
    ' Oblicz ilość godzin roboczych
    workHours = workDaysCount * 8

    ' Write total work days and hours
    With wsTarget
        .Cells(lastRow + 1, 1).Value = "Dni robocze"
        .Cells(lastRow + 2, 1).Value = "Godz. robocze"

        .Cells(lastRow + 1, 3).Value = workDaysCount
        .Cells(lastRow + 2, 3).Value = workHours

        ' Format the summary rows
        .Cells(lastRow + 1, 1).Font.Bold = True
        .Cells(lastRow + 2, 1).Font.Bold = True
        .Cells(lastRow + 1, 3).Font.Bold = True
        .Cells(lastRow + 2, 3).Font.Bold = True
        .Cells(lastRow + 1, 3).HorizontalAlignment = xlCenter
        .Cells(lastRow + 2, 3).HorizontalAlignment = xlCenter

        ' Dodaj sumy godzin dla każdego pracownika w wierszu "Godz. robocze"
        For j = 1 To employeeCount
            .Cells(lastRow + 2, 3 + (j * 2)).Value = employeeHours(j)
            .Cells(lastRow + 2, 3 + (j * 2)).Font.Bold = True
            .Cells(lastRow + 2, 3 + (j * 2)).HorizontalAlignment = xlCenter
        Next j
    End With

    ' Wstaw ilość godzin pracy do ComboBox F2
    If Not wsComboBox Is Nothing Then
        wsComboBox.Range("F2").Value = workHours
    End If
End Sub

Function ValidateInputs() As Boolean
    Dim wsPracownicy As Worksheet
    Dim lastRowPracownicy As Long
    
    On Error Resume Next
    Set wsPracownicy = ThisWorkbook.Sheets("Pracownicy")
    On Error GoTo 0
    
    If wsPracownicy Is Nothing Then
        ValidateInputs = False
        Exit Function
    End If
    
    ' Get the number of employees from the Pracownicy sheet
    lastRowPracownicy = wsPracownicy.Cells(Rows.Count, "B").End(xlUp).row
    employeeCount = lastRowPracownicy
    
    ' Validate month and year
    If currentMonth < 1 Or currentYear < 2025 Or currentYear > 2032 Then
        ValidateInputs = False
        Exit Function
    End If
    
    ValidateInputs = True
End Function

Sub FindHolidays()
    Dim wsHolidays As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim holidayDate As Date
    
    ' Initialize the holidays collection
    Set holidays = New Collection
    
    ' Check if the "swieta" worksheet exists
    On Error Resume Next
    Set wsHolidays = ThisWorkbook.Sheets("swieta")
    On Error GoTo 0
    
    If wsHolidays Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'swieta'!", vbExclamation
        Exit Sub
    End If
    
    lastRow = wsHolidays.Cells(Rows.Count, 1).End(xlUp).row
    
    ' Loop through all dates in column A
    For i = 1 To lastRow
        If IsDate(wsHolidays.Cells(i, 1).Value) Then
            holidayDate = CDate(wsHolidays.Cells(i, 1).Value)
            
            ' Only add holiday if it's in the current month and year
            If Month(holidayDate) = currentMonth And Year(holidayDate) = currentYear Then
                holidays.Add holidayDate
            End If
        End If
    Next i
End Sub

Function isHoliday(checkDate As Date) As Boolean
    Dim holidayDate As Variant
    
    If holidays Is Nothing Then
        Set holidays = New Collection
    End If
    
    For Each holidayDate In holidays
        If checkDate = holidayDate Then
            isHoliday = True
            Exit Function
        End If
    Next holidayDate
    
    isHoliday = False
End Function

' Initialization procedure to be called when the workbook opens
Private Sub Workbook_Open()
    Call CreateInitializationButton
    
    ' Włącz obsługę zdarzeń
    Application.EnableEvents = True
End Sub

' Funkcja do czyszczenia arkusza "Kalendarz"
Sub ClearKalendarzSheet()
    Dim wsTarget As Worksheet

    ' Sprawdź czy arkusz "Kalendarz" istnieje
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("Kalendarz")
    On Error GoTo 0

    If wsTarget Is Nothing Then
        Exit Sub
    End If

    ' Wyczyść zawartość arkusza
    Application.ScreenUpdating = False
    wsTarget.Cells.Clear
    Application.ScreenUpdating = True
End Sub

' Funkcja do czyszczenia wierszy od 4 w dół w arkuszu "Harmonogram Pracy"
Sub ClearHarmonogramPracyRows(wsHarmonogram As Worksheet)
    ' Sprawdź czy arkusz istnieje
    If wsHarmonogram Is Nothing Then
        Exit Sub
    End If

    ' Wyłącz ochronę arkusza, jeśli jest włączona
    On Error Resume Next
    wsHarmonogram.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0

    ' Wyczyść komórki od wiersza 4 w dół
    Application.ScreenUpdating = False
    wsHarmonogram.Rows("4:" & wsHarmonogram.Rows.Count).Clear
    Application.ScreenUpdating = True
End Sub

Sub CopyKalendarzToHarmonogramPracy()
    Dim wsKalendarz As Worksheet
    Dim wsHarmonogram As Worksheet
    Dim wsPracownicy As Worksheet
    Dim wsComboBox As Worksheet
    Dim wsZmiany As Worksheet
    Dim wsSwieta As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim daysInMonth As Integer
    Dim selectedMonth As Integer
    Dim selectedYear As Integer
    Dim previousMonth As Integer
    Dim previousYear As Integer

    ' Sprawdź czy arkusze istnieją
    On Error Resume Next
    Set wsKalendarz = ThisWorkbook.Sheets("Kalendarz")
    Set wsHarmonogram = ThisWorkbook.Sheets("Harmonogram Pracy")
    Set wsPracownicy = ThisWorkbook.Sheets("Pracownicy")
    Set wsComboBox = ThisWorkbook.Sheets("ComboBox")
    Set wsZmiany = ThisWorkbook.Sheets("Zmiany")
    Set wsSwieta = ThisWorkbook.Sheets("Święta")
    On Error GoTo 0

    If wsKalendarz Is Nothing Or wsHarmonogram Is Nothing Or wsPracownicy Is Nothing Or wsComboBox Is Nothing Then
        MsgBox "Nie znaleziono wymaganych arkuszy!", vbExclamation
        Exit Sub
    End If

    ' Pobierz miesiąc i rok z arkusza ComboBox
    selectedMonth = Val(wsComboBox.Range("B2").Value)
    selectedYear = Val(wsComboBox.Range("D2").Value)

    ' Inicjalizuj kolekcję świąt
    Set holidays = New Collection

    ' Ustaw zmienne globalne dla FindHolidays
    currentMonth = selectedMonth
    currentYear = selectedYear

    ' Znajdź święta dla wybranego miesiąca i roku
    Call FindHolidays

    ' Oblicz poprzedni miesiąc i rok (dla potencjalnego przyszłego użycia)
    If selectedMonth = 1 Then
        previousMonth = 12
        previousYear = selectedYear - 1
    Else
        previousMonth = selectedMonth - 1
        previousYear = selectedYear
    End If

    ' Oblicz liczbę dni w miesiącu
    daysInMonth = Day(DateSerial(selectedYear, selectedMonth + 1, 0))

    ' Pobierz ostatni wiersz z danymi w arkuszu Kalendarz
    lastRow = wsKalendarz.Cells(Rows.Count, 1).End(xlUp).row

    ' Znajdź wiersz z "Godz. robocze" w arkuszu Kalendarz
    Dim godzRoboczeRow As Long
    godzRoboczeRow = 0

    For i = lastRow To lastRow + 5 ' Szukaj w kilku wierszach poniżej ostatniego wiersza z datami
        If wsKalendarz.Cells(i, 1).Value = "Godz. robocze" Then
            godzRoboczeRow = i
            Exit For
        End If
    Next i

    If godzRoboczeRow = 0 Then
        MsgBox "Nie znaleziono wiersza 'Godz. robocze' w arkuszu Kalendarz!", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Wyłącz ochronę arkusza, jeśli jest włączona
    On Error Resume Next
    wsHarmonogram.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0

    ' Ustaw nagłówki w arkuszu Harmonogram Pracy
    With wsHarmonogram
        ' Wiersz 4: Nagłówki
        .Cells(4, 1).Value = "Pracownik"
        .Cells(4, 2).Value = "Nadgodz"
        .Cells(4, 2).Orientation = -90 ' Orientacja -90 stopni (pionowo od dołu do góry)

        ' Ustaw numery dni w miesiącu oraz skrócone nazwy dni tygodnia jako nagłówki kolumn
        For i = 1 To daysInMonth
            Dim dayDate As Date
            Dim dayOfWeek As Integer
            Dim dayName As String
            
            ' Utwórz datę dla danego dnia
            dayDate = DateSerial(selectedYear, selectedMonth, i)
            dayOfWeek = Weekday(dayDate)
            
            ' Pobierz skróconą nazwę dnia tygodnia (pierwsze 2 znaki)
            dayName = Left(WeekdayName(dayOfWeek, True, 1), 2)
            
            ' Ustaw wartość komórki jako skrócona nazwa dnia i numer dnia
            .Cells(4, i + 2).Value = dayName & vbCrLf & i
            
            ' Włącz zawijanie tekstu
            .Cells(4, i + 2).WrapText = True
        Next i
        
        ' Zwiększ wysokość wiersza nagłówkowego dla lepszej widoczności
        .Rows(4).RowHeight = 30

        ' Dodaj nagłówek dla godzin roboczych z wartością z ComboBox F2 i napisem "Godz. robocze" w tej samej komórce
        ' Najpierw czyścimy komórkę i formatowanie
        .Cells(4, daysInMonth + 3).ClearContents
        .Cells(4, daysInMonth + 3).ClearFormats

        ' Zwiększamy wysokość wiersza, aby tekst był widoczny
        .Rows(4).RowHeight = 45 ' Zwiększamy wysokość wiersza dla lepszej widoczności

        ' Ustawiamy wyrównanie do dołu komórki
        .Cells(4, daysInMonth + 3).VerticalAlignment = xlBottom
        .Cells(4, daysInMonth + 3).HorizontalAlignment = xlCenter

        ' Dodajemy wartość z ComboBox F2 i "Godz. robocze" w dwóch liniach
        .Cells(4, daysInMonth + 3).Value = wsComboBox.Range("F2").Value & vbCrLf & "Godz. robocze"

        ' Formatujemy pierwszą część (wartość liczbową) - większa czcionka, czerwona, pogrubiona
        .Cells(4, daysInMonth + 3).Characters(1, Len(CStr(wsComboBox.Range("F2").Value))).Font.Size = 18
        .Cells(4, daysInMonth + 3).Characters(1, Len(CStr(wsComboBox.Range("F2").Value))).Font.Color = vbRed
        .Cells(4, daysInMonth + 3).Characters(1, Len(CStr(wsComboBox.Range("F2").Value))).Font.Bold = True

        ' Formatujemy drugą część (napis "Godz. robocze") - normalna czcionka, czarna
        .Cells(4, daysInMonth + 3).Characters(Len(CStr(wsComboBox.Range("F2").Value)) + 3).Font.Size = 10
        .Cells(4, daysInMonth + 3).Characters(Len(CStr(wsComboBox.Range("F2").Value)) + 3).Font.Color = vbBlack
        .Cells(4, daysInMonth + 3).Characters(Len(CStr(wsComboBox.Range("F2").Value)) + 3).Font.Bold = False

        ' Ustawiamy tekst poziomo (normalnie)
        .Cells(4, daysInMonth + 3).Orientation = 0 ' Orientacja 0 stopni (poziomo)

        ' Ustawiamy normalny rozmiar czcionki dla całej komórki
        .Cells(4, daysInMonth + 3).Font.Size = 10

        ' Formatujemy pierwszą część (wartość liczbową) - większa czcionka, czerwona, pogrubiona
        .Cells(4, daysInMonth + 3).Characters(1, Len(CStr(wsComboBox.Range("F2").Value))).Font.Size = 18
        .Cells(4, daysInMonth + 3).Characters(1, Len(CStr(wsComboBox.Range("F2").Value))).Font.Color = vbRed
        .Cells(4, daysInMonth + 3).Characters(1, Len(CStr(wsComboBox.Range("F2").Value))).Font.Bold = True

        ' Formatujemy drugą część (napis "Godz. robocze") - normalna czcionka, czarna
        .Cells(4, daysInMonth + 3).Characters(Len(CStr(wsComboBox.Range("F2").Value)) + 3).Font.Size = 10
        .Cells(4, daysInMonth + 3).Characters(Len(CStr(wsComboBox.Range("F2").Value)) + 3).Font.Color = vbBlack
        .Cells(4, daysInMonth + 3).Characters(Len(CStr(wsComboBox.Range("F2").Value)) + 3).Font.Bold = False
        .Cells(4, daysInMonth + 3).Characters(Len(CStr(wsComboBox.Range("F2").Value)) + 3).Font.Bold = True

        ' Ustawiamy wyrównanie do środka
        .Cells(4, daysInMonth + 3).HorizontalAlignment = xlCenter
        .Cells(4, daysInMonth + 3).VerticalAlignment = xlCenter

        ' Formatuj pozostałe nagłówki
        .Range(.Cells(4, 1), .Cells(4, daysInMonth + 2)).Font.Bold = True
        .Range(.Cells(4, 1), .Cells(4, daysInMonth + 2)).HorizontalAlignment = xlCenter

        ' Wiersze od 5 w dół: Dane pracowników
        For i = 1 To employeeCount
            ' Kolumna 1: Dane pracownika z arkusza Pracownicy
            .Cells(i + 4, 1).Value = wsPracownicy.Cells(i, "B").Value

            ' Kolumna 2: Ilość nadgodzin danego pracownika z poprzedniego miesiąca
            .Cells(i + 4, 2).Value = wsKalendarz.Cells(2, 3 + (i * 2)).Value

            ' Kolumny od 3: Zmiana danego pracownika w danym dniu
            For j = 3 To lastRow
                ' Pobierz datę z arkusza Kalendarz
                Dim currentDate As Date
                Dim dayOfMonth As Integer

                If IsDate(wsKalendarz.Cells(j, 1).Value) Then
                    currentDate = CDate(wsKalendarz.Cells(j, 1).Value)
                    dayOfMonth = Day(currentDate)

                    ' Sprawdź czy dzień jest w bieżącym miesiącu
                    If Month(currentDate) = selectedMonth And Year(currentDate) = selectedYear Then
                        ' Kopiuj zmianę pracownika
                        .Cells(i + 4, dayOfMonth + 2).Value = wsKalendarz.Cells(j, 3 + (i * 2) - 1).Value
                    End If
                End If
            Next j

            ' Ostatnia kolumna: Godziny robocze danego pracownika - pobierz z wiersza "Godz. robocze" w arkuszu Kalendarz
            Dim employeeHours As Double
            employeeHours = Val(wsKalendarz.Cells(godzRoboczeRow, 3 + (i * 2)).Value)
            .Cells(i + 4, daysInMonth + 3).Value = employeeHours
            
            ' Pobierz wartość docelową z ComboBox F2
            Dim targetHours As Double
            targetHours = Val(wsComboBox.Range("F2").Value)
            
            ' Formatuj komórkę w zależności od tego, czy suma godzin jest równa wartości docelowej
            With .Cells(i + 4, daysInMonth + 3)
                .Font.Bold = False
                If employeeHours = targetHours Then
                    ' Suma godzin jest równa wartości docelowej - formatuj na zielono i pogrubioną czcionką
                    .Font.Color = RGB(0, 128, 0) ' Zielony
                    .Font.Bold = True
                Else
                    ' Suma godzin jest różna od wartości docelowej - formatuj na czerwono
                    .Font.Color = RGB(255, 0, 0) ' Czerwony
                End If
            End With
        Next i

        ' Formatuj dane - kolumny od 2 do ostatniej wyśrodkowane
        .Range(.Cells(5, 2), .Cells(employeeCount + 4, daysInMonth + 3)).HorizontalAlignment = xlCenter
        
        ' Kolumna A (nazwiska pracowników) wyrównana do lewej
        .Range(.Cells(5, 1), .Cells(employeeCount + 4, 1)).HorizontalAlignment = xlLeft

        ' Ustaw szerokość kolumn
        .Columns(2).ColumnWidth = 5 ' Kolumna z "Nadgodz. z..."

        ' Ustaw szerokość kolumn z numerami dni
        For i = 3 To daysInMonth + 2
            .Columns(i).ColumnWidth = 5
        Next i

        ' Ustaw szerokość kolumny "Godz. robocze"
        .Columns(daysInMonth + 3).ColumnWidth = 13
        
        ' Automatycznie dopasuj szerokość kolumny A do najdłuższej nazwy pracownika
        Dim maxLength As Integer
        Dim employeeName As String
        
        maxLength = Len("Pracownik") ' Minimalna szerokość dla nagłówka
        
        ' Sprawdź długość każdej nazwy pracownika
        For i = 1 To employeeCount
            employeeName = wsPracownicy.Cells(i, "B").Value
            If Len(employeeName) > maxLength Then
                maxLength = Len(employeeName)
            End If
        Next i
        
        ' Ustaw szerokość kolumny A z niewielkim zapasem
        .Columns(1).ColumnWidth = maxLength + 2

        ' Zastosuj ochronę arkusza z możliwością edycji tylko wybranych komórek
        ' Najpierw odblokuj wszystkie komórki, które mają być edytowalne
        .Cells.Locked = True ' Najpierw zablokuj wszystkie komórki

        ' Odblokuj komórki z nadgodzinami (kolumna 2) dla wierszy pracowników
        For i = 1 To employeeCount
            .Cells(i + 4, 2).Locked = False
            ' Ustaw format liczbowy
            .Cells(i + 4, 2).NumberFormat = "0"
            
            ' Dodaj walidację danych dla nadgodzin - tylko wartości liczbowe od -64 do +64
            With .Cells(i + 4, 2).Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:="-64", Formula2:="64"
                .IgnoreBlank = True
                .InputTitle = "Wprowadź nadgodziny"
                .ErrorTitle = "Nieprawidłowa wartość"
                .InputMessage = "Wprowadź liczbę całkowitą od -64 do +64."
                .ErrorMessage = "Wartość musi być liczbą całkowitą z zakresu od -64 do +64."
            End With
        Next i

        ' Odblokuj komórki z dniami (kolumny 3 do daysInMonth+2) dla wierszy pracowników
        ' z wyjątkiem świąt
        Dim cellDate As Date
        Dim cellDayOfWeek As Integer

        For i = 1 To daysInMonth
            ' Utwórz datę dla danego dnia
            cellDate = DateSerial(selectedYear, selectedMonth, i)
            cellDayOfWeek = Weekday(cellDate)

            ' Jeśli to święto, zablokuj komórki i ustaw DW
            If isHoliday(cellDate) Then
                For j = 1 To employeeCount
                    .Cells(j + 4, i + 2).Locked = True
                    .Cells(j + 4, i + 2).Value = "DW"
                Next j
            Else
                ' Dla dni, które nie są świętami, odblokuj komórki i dodaj walidację
                For j = 1 To employeeCount
                    .Cells(j + 4, i + 2).Locked = False
                    ' Ustaw format tekstowy
                    .Cells(j + 4, i + 2).NumberFormat = "@"
                    
                    ' Dodaj walidację danych w zależności od typu dnia
                    Dim validValues As String
                    
                    ' Sprawdź typ dnia
                    If cellDayOfWeek >= 2 And cellDayOfWeek <= 6 Then
                        ' Dni robocze (poniedziałek-piątek)
                        validValues = "R1,R2,R3,UW,DW"
                    ElseIf cellDayOfWeek = 7 Then
                        ' Sobota
                        validValues = "SO1,SO2,UW,DW"
                    ElseIf cellDayOfWeek = 1 Then
                        ' Niedziela
                        validValues = "NIED,UW,DW"
                    End If
                    
                    ' Dodaj walidację danych
                    With .Cells(j + 4, i + 2).Validation
                        .Delete
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                             Operator:=xlBetween, Formula1:=validValues
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .InputTitle = "Wybierz zmianę"
                        .ErrorTitle = "Nieprawidłowa zmiana"
                        .InputMessage = "Wybierz jedną z dozwolonych zmian dla tego dnia."
                        .ErrorMessage = "Wprowadzona wartość nie jest dozwoloną zmianą dla tego typu dnia."
                    End With
                Next j
            End If
        Next i

        ' Najpierw wyczyść wszystkie formatowania tła
        .UsedRange.Interior.ColorIndex = xlNone

        ' Formatowanie kolorami dni specjalnych (święta, soboty, niedziele)
        Dim formatDate As Date
        Dim formatDayOfWeek As Integer

        ' Przejdź przez wszystkie dni miesiąca
        For i = 1 To daysInMonth
            ' Utwórz datę dla danego dnia
            formatDate = DateSerial(selectedYear, selectedMonth, i)
            formatDayOfWeek = Weekday(formatDate)

            ' Sprawdź czy to święto
            If isHoliday(formatDate) Then
                ' Zaznacz tylko komórki od wiersza 4 do ostatniego pracownika na lekko czerwono
                For j = 4 To employeeCount + 4
                    .Cells(j, i + 2).Interior.Color = RGB(255, 200, 200) ' Lekki czerwony
                Next j
            ' Sprawdź czy to sobota lub niedziela
            ElseIf formatDayOfWeek = 1 Or formatDayOfWeek = 7 Then
                ' Zaznacz tylko komórki od wiersza 4 do ostatniego pracownika na lekko żółto
                For j = 4 To employeeCount + 4
                    .Cells(j, i + 2).Interior.Color = RGB(255, 255, 200) ' Lekki żółty
                Next j
            End If
        Next i

        ' Formatowanie tekstu "DW" na czerwono i "UW" na granatowo
        Dim cell As Range

        ' Przejdź przez wszystkie komórki z danymi pracowników
        For i = 1 To employeeCount
            For j = 3 To daysInMonth + 2
                Set cell = .Cells(i + 4, j)
                ' Usuń wytłuszczenie ze wszystkich komórek
                cell.Font.Bold = False
                
                ' Ustaw kolor w zależności od wartości
                If cell.Value = "DW" Then
                    cell.Font.Color = RGB(255, 0, 0) ' Czerwony
                ElseIf cell.Value = "UW" Then
                    cell.Font.Color = RGB(0, 176, 240) ' Niebieski (cyan)
                Else
                    cell.Font.Color = RGB(0, 0, 0) ' Czarny (domyślny)
                End If
            Next j
        Next i

        ' Włącz ochronę arkusza z możliwością zaznaczania zablokowanych komórek
        ' Najpierw wyłącz ochronę, jeśli jest włączona
        On Error Resume Next
        .Unprotect Password:=SHEET_PASSWORD
        On Error GoTo 0

        ' Włącz ochronę z odpowiednimi opcjami
        .Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True, _
                AllowFormattingCells:=False, _
                AllowFormattingColumns:=False, _
                AllowFormattingRows:=False, _
                AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, _
                AllowInsertingHyperlinks:=False, _
                AllowDeletingColumns:=False, _
                AllowDeletingRows:=False, _
                AllowSorting:=False, _
                AllowFiltering:=False, _
                AllowUsingPivotTables:=False
    End With
    
    ' Sprawdź czy wszystkie wymagane zmiany są obsadzone dla każdego dnia
    Call CheckRequiredShifts(wsHarmonogram, CLng(daysInMonth), employeeCount)

    Application.ScreenUpdating = True
End Sub

' Funkcja sprawdzająca, czy wszystkie wymagane zmiany są obsadzone dla każdego dnia
Sub CheckRequiredShifts(wsHarmonogram As Worksheet, daysInMonth As Long, employeeCount As Integer)
    Dim i As Long, j As Long
    Dim dayOfWeek As Integer
    Dim cellDate As Date
    Dim selectedMonth As Long, selectedYear As Long
    Dim wsComboBox As Worksheet
    Dim shiftCounts As Object
    Dim isValid As Boolean
    Dim requiredShifts As String
    
    ' Pobierz arkusz ComboBox
    On Error Resume Next
    Set wsComboBox = ThisWorkbook.Sheets("ComboBox")
    On Error GoTo 0
    
    If wsComboBox Is Nothing Then Exit Sub
    
    ' Pobierz miesiąc i rok
    selectedMonth = Val(wsComboBox.Range("B2").Value)
    selectedYear = Val(wsComboBox.Range("D2").Value)
    
    ' Wyczyść cały wiersz z komunikatami "Źle" przed rozpoczęciem sprawdzania
    wsHarmonogram.Rows(employeeCount + 5).ClearContents
    
    ' Dla każdego dnia miesiąca
    For i = 1 To daysInMonth
        ' Utwórz słownik do zliczania zmian
        Set shiftCounts = CreateObject("Scripting.Dictionary")
        
        ' Utwórz datę dla danego dnia
        cellDate = DateSerial(selectedYear, selectedMonth, i)
        dayOfWeek = Weekday(cellDate)
        
        ' Sprawdź czy to święto
        If Not isHoliday(cellDate) Then
            ' Zlicz wszystkie zmiany dla danego dnia
            For j = 1 To employeeCount
                Dim shift As String
                shift = wsHarmonogram.Cells(j + 4, i + 2).Value
                
                ' Jeśli zmiana nie jest pusta, dodaj ją do słownika
                If Len(shift) > 0 Then
                    If shiftCounts.Exists(shift) Then
                        shiftCounts(shift) = shiftCounts(shift) + 1
                    Else
                        shiftCounts.Add shift, 1
                    End If
                End If
            Next j
            
            ' Sprawdź czy wszystkie wymagane zmiany są obsadzone
            isValid = True
            
            ' Określ wymagane zmiany w zależności od dnia tygodnia
            Select Case dayOfWeek
                Case 2 To 6 ' Poniedziałek do piątku (dni robocze)
                    ' Sprawdź czy jest przynajmniej jedna zmiana R1, R2 i R3
                    If Not shiftCounts.Exists("R1") Or Not shiftCounts.Exists("R2") Or Not shiftCounts.Exists("R3") Then
                        isValid = False
                        requiredShifts = "R1, R2, R3"
                    End If
                    
                Case 7 ' Sobota
                    ' Sprawdź czy jest przynajmniej jedna zmiana SO1 i SO2
                    If Not shiftCounts.Exists("SO1") Or Not shiftCounts.Exists("SO2") Then
                        isValid = False
                        requiredShifts = "SO1, SO2"
                    End If
                    
                Case 1 ' Niedziela
                    ' Sprawdź czy jest przynajmniej jedna zmiana NIED
                    If Not shiftCounts.Exists("NIED") Then
                        isValid = False
                        requiredShifts = "NIED"
                    End If
            End Select
            
            ' Jeśli nie wszystkie wymagane zmiany są obsadzone, dodaj komunikat "Źle" pod ostatnim pracownikiem
            If Not isValid Then
                ' Dodaj komunikat "Źle" pod ostatnim pracownikiem
                With wsHarmonogram.Cells(employeeCount + 5, i + 2)
                    .Value = "Źle"
                    .Font.Color = RGB(255, 0, 0) ' Czerwony
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter ' Wycentruj tekst
                End With
                
                ' Wyczyść komórkę w kolumnie z podsumowaniem godzin
                wsHarmonogram.Cells(employeeCount + 5, daysInMonth + 3).ClearContents
            Else
                ' Wyczyść komórkę, jeśli wszystko jest OK
                wsHarmonogram.Cells(employeeCount + 5, i + 2).ClearContents
            End If
        End If
    Next i
End Sub

' Procedura obsługi zdarzenia zmiany wartości w arkuszu
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Sprawdź, czy zmiana dotyczy arkusza "Harmonogram Pracy"
    If Sh.Name = "Harmonogram Pracy" Then
        ' Obsłuż zmianę w arkuszu "Harmonogram Pracy"
        Call HandleHarmonogramPracyChange(Target)
    End If
End Sub

' Procedura obsługi zmiany w arkuszu "Harmonogram Pracy"
Private Sub HandleHarmonogramPracyChange(ByVal Target As Range)
    ' Zabezpieczenie przed rekurencją
    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True
    
    ' Wyłącz obsługę zdarzeń, aby uniknąć rekurencji
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Sprawdź, czy zmiana dotyczy tylko jednej komórki
    If Target.Cells.Count = 1 Then
        ' Pobierz numer wiersza i kolumny zmienionej komórki
        Dim row As Long, col As Long
        row = Target.row
        col = Target.Column
        
        ' Sprawdź, czy zmiana dotyczy komórki z dniem (kolumny od 3 do liczba_dni+2)
        ' i wiersza pracownika (wiersze od 5 do liczba_pracowników+4)
        If row >= 5 And col >= 3 Then
            ' Pobierz arkusze
            Dim wsHarmonogram As Worksheet
            Dim wsKalendarz As Worksheet
            Dim wsZmiany As Worksheet
            Dim wsComboBox As Worksheet
            
            Set wsHarmonogram = ThisWorkbook.Sheets("Harmonogram Pracy")
            Set wsKalendarz = ThisWorkbook.Sheets("Kalendarz")
            Set wsZmiany = ThisWorkbook.Sheets("Zmiany")
            Set wsComboBox = ThisWorkbook.Sheets("ComboBox")
            
            ' Pobierz miesiąc i rok
            Dim selectedMonth As Long, selectedYear As Long
            selectedMonth = CLng(Val(wsComboBox.Range("B2").Value))
            selectedYear = CLng(Val(wsComboBox.Range("D2").Value))
            
            ' Oblicz liczbę dni w miesiącu
            Dim daysInMonth As Long
            daysInMonth = Day(DateSerial(selectedYear, selectedMonth + 1, 0))
            
            ' Sprawdź, czy kolumna jest w zakresie dni miesiąca
            If col <= daysInMonth + 2 Then
                ' Pobierz numer pracownika (indeks)
                Dim employeeIndex As Long
                employeeIndex = row - 4
                
                ' Pobierz dzień miesiąca
                Dim dayOfMonth As Long
                dayOfMonth = col - 2
                
                ' Pobierz nową wartość zmiany i zamień na wielkie litery
                Dim newShift As String
                newShift = UCase(Target.Value)
                
                ' Jeśli wartość się zmieniła, zaktualizuj komórkę
                If newShift <> Target.Value Then
                    Target.Value = newShift
                End If
                
                ' Formatuj komórkę w zależności od wartości
                If newShift = "DW" Then
                    Target.Font.Color = RGB(255, 0, 0) ' Czerwony
                ElseIf newShift = "UW" Then
                    Target.Font.Color = RGB(0, 176, 240) ' Niebieski (cyan)
                Else
                    Target.Font.Color = RGB(0, 0, 0) ' Czarny (domyślny)
                End If
                
                ' Pobierz liczbę godzin dla tej zmiany
                Dim hours As Variant
                hours = GetGodziny(wsZmiany, newShift)
                
                ' Znajdź odpowiedni wiersz w arkuszu Kalendarz
                Dim kalendarzRow As Long
                kalendarzRow = 0
                
                ' Utwórz datę dla danego dnia
                Dim currentDate As Date
                On Error Resume Next
                currentDate = DateSerial(selectedYear, selectedMonth, dayOfMonth)
                If Err.Number <> 0 Then
                    MsgBox "Błąd przy tworzeniu daty: " & Err.Description, vbExclamation
                    On Error GoTo ErrorHandler
                    GoTo ExitSub
                End If
                On Error GoTo ErrorHandler
                
                ' Znajdź wiersz w arkuszu Kalendarz odpowiadający tej dacie
                Dim i As Long
                For i = 3 To wsKalendarz.Cells(Rows.Count, 1).End(xlUp).row
                    If IsDate(wsKalendarz.Cells(i, 1).Value) Then
                        If CDate(wsKalendarz.Cells(i, 1).Value) = currentDate Then
                            kalendarzRow = i
                            Exit For
                        End If
                    End If
                Next i
                
                ' Jeśli znaleziono odpowiedni wiersz, zaktualizuj wartości w arkuszu Kalendarz
                If kalendarzRow > 0 Then
                    ' Zaktualizuj zmianę w arkuszu Kalendarz
                    wsKalendarz.Cells(kalendarzRow, 3 + (employeeIndex * 2) - 1).Value = newShift
                    
                    ' Zaktualizuj liczbę godzin w arkuszu Kalendarz
                    wsKalendarz.Cells(kalendarzRow, 3 + (employeeIndex * 2)).Value = hours
                    
                    ' Przelicz sumę godzin dla pracownika
                    Call UpdateEmployeeHours(employeeIndex, daysInMonth)
                    
                    ' Sprawdź czy wszystkie wymagane zmiany są obsadzone dla każdego dnia
                    Dim empCount As Integer
                    empCount = wsHarmonogram.Cells(Rows.Count, 1).End(xlUp).row - 4
                    Call CheckRequiredShifts(wsHarmonogram, CLng(daysInMonth), empCount)
                End If
            End If
        ElseIf row >= 5 And col = 2 Then
            ' Zmiana w kolumnie nadgodzin
            ' Pobierz arkusze
            Dim wsHarmonogram2 As Worksheet
            Dim wsKalendarz2 As Worksheet
            Dim wsComboBox2 As Worksheet
            
            Set wsHarmonogram2 = ThisWorkbook.Sheets("Harmonogram Pracy")
            Set wsKalendarz2 = ThisWorkbook.Sheets("Kalendarz")
            Set wsComboBox2 = ThisWorkbook.Sheets("ComboBox")
            
            ' Pobierz miesiąc i rok
            Dim selectedMonth2 As Long, selectedYear2 As Long
            selectedMonth2 = CLng(Val(wsComboBox2.Range("B2").Value))
            selectedYear2 = CLng(Val(wsComboBox2.Range("D2").Value))
            
            ' Oblicz liczbę dni w miesiącu
            Dim daysInMonth2 As Long
            daysInMonth2 = Day(DateSerial(selectedYear2, selectedMonth2 + 1, 0))
            
            ' Pobierz numer pracownika (indeks)
            Dim employeeIndex2 As Long
            employeeIndex2 = row - 4
            
            ' Pobierz nową wartość nadgodzin
            Dim newOvertime As Double
            newOvertime = Val(Target.Value)
            
            ' Zaktualizuj wartość nadgodzin w arkuszu Kalendarz
            wsKalendarz2.Cells(2, 3 + (employeeIndex2 * 2)).Value = newOvertime
            
            ' Przelicz sumę godzin dla pracownika
            Call UpdateEmployeeHours(employeeIndex2, daysInMonth2)
        End If
    End If
    
ExitSub:
    ' Włącz obsługę zdarzeń
    Application.EnableEvents = True
    isRunning = False
    Exit Sub
    
ErrorHandler:
    MsgBox "Wystąpił błąd w HandleHarmonogramPracyChange: " & Err.Description, vbExclamation
    Resume ExitSub
End Sub

' Funkcja do aktualizacji sumy godzin dla pracownika
Private Sub UpdateEmployeeHours(employeeIndex As Long, daysInMonth As Long)
    On Error GoTo ErrorHandler
    
    Dim wsKalendarz As Worksheet
    Dim wsHarmonogram As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalHours As Double
    
    ' Pobierz arkusze
    Set wsKalendarz = ThisWorkbook.Sheets("Kalendarz")
    Set wsHarmonogram = ThisWorkbook.Sheets("Harmonogram Pracy")
    
    ' Pobierz ostatni wiersz z danymi w arkuszu Kalendarz
    lastRow = wsKalendarz.Cells(Rows.Count, 1).End(xlUp).row
    
    ' Znajdź wiersz z "Godz. robocze" w arkuszu Kalendarz
    Dim godzRoboczeRow As Long
    godzRoboczeRow = 0
    
    For i = lastRow To lastRow + 5 ' Szukaj w kilku wierszach poniżej ostatniego wiersza z datami
        If wsKalendarz.Cells(i, 1).Value = "Godz. robocze" Then
            godzRoboczeRow = i
            Exit For
        End If
    Next i
    
    If godzRoboczeRow = 0 Then
        Exit Sub
    End If
    
    ' Najpierw dodaj nadgodziny z wiersza 2
    totalHours = Val(wsKalendarz.Cells(2, 3 + (employeeIndex * 2)).Value)
    
    ' Sumuj godziny dla pracownika
    For i = 3 To lastRow
        If IsDate(wsKalendarz.Cells(i, 1).Value) Then
            totalHours = totalHours + Val(wsKalendarz.Cells(i, 3 + (employeeIndex * 2)).Value)
        End If
    Next i
    
    ' Zaktualizuj sumę godzin w arkuszu Kalendarz
    wsKalendarz.Cells(godzRoboczeRow, 3 + (employeeIndex * 2)).Value = totalHours
    
    ' Zaktualizuj sumę godzin w arkuszu Harmonogram Pracy
    wsHarmonogram.Cells(employeeIndex + 4, daysInMonth + 3).Value = totalHours
    
    ' Pobierz wartość docelową z ComboBox F2
    Dim targetHours As Double
    Dim wsComboBox As Worksheet
    Set wsComboBox = ThisWorkbook.Sheets("ComboBox")
    targetHours = Val(wsComboBox.Range("F2").Value)
    
    ' Formatuj komórkę w zależności od tego, czy suma godzin jest równa wartości docelowej
    With wsHarmonogram.Cells(employeeIndex + 4, daysInMonth + 3)
        .Font.Bold = False
        If totalHours = targetHours Then
            ' Suma godzin jest równa wartości docelowej - formatuj na zielono i pogrubioną czcionką
            .Font.Color = RGB(0, 128, 0) ' Zielony
            .Font.Bold = True
        Else
            ' Suma godzin jest różna od wartości docelowej - formatuj na czerwono
            .Font.Color = RGB(255, 0, 0) ' Czerwony
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Wystąpił błąd w UpdateEmployeeHours: " & Err.Description, vbExclamation
    Resume Next
End Sub

' Funkcja do zmiany nazwy pliku Excel
Sub RenameWorkbook(newFileName As String)
    Dim currentPath As String
    Dim newPath As String
    Dim parts() As String
    Dim formattedFileName As String

    ' Sprawdź czy nazwa pliku zawiera format dzial-rok-miesiac.xlsm
    If InStr(newFileName, "-") > 0 Then
        parts = Split(newFileName, "-")
        If UBound(parts) >= 2 Then
            ' Format miesiąca jako dwucyfrowy
            Dim dzial As String
            Dim rok As String
            Dim miesiac As String
            Dim rozszerzenie As String

            dzial = parts(0)
            rok = parts(1)

            ' Sprawdź czy ostatnia część zawiera rozszerzenie
            If InStr(parts(2), ".") > 0 Then
                Dim miesiacIRozszerzenie() As String
                miesiacIRozszerzenie = Split(parts(2), ".")
                miesiac = miesiacIRozszerzenie(0)
                rozszerzenie = "." & miesiacIRozszerzenie(1)
            Else
                miesiac = parts(2)
                rozszerzenie = ".xlsm"
            End If

            ' Format miesiąca jako dwucyfrowy
            miesiac = Format(Val(miesiac), "00")

            ' Złóż nazwę pliku z powrotem
            formattedFileName = dzial & "-" & rok & "-" & miesiac & rozszerzenie
        Else
            formattedFileName = newFileName
        End If
    Else
        formattedFileName = newFileName
    End If

    ' Pobierz aktualną ścieżkę pliku
    currentPath = ThisWorkbook.FullName

    ' Dodaj obsługę błędów
    On Error Resume Next

    ' Jeśli plik nie został jeszcze zapisany, zapisz go w domyślnej lokalizacji
    If currentPath = "" Then
        ThisWorkbook.SaveAs Filename:=formattedFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
        ' Utwórz nową ścieżkę z tą samą lokalizacją, ale nową nazwą
        newPath = Left(currentPath, InStrRev(currentPath, "\")) & formattedFileName

        ' Zapisz plik z nową nazwą
        ThisWorkbook.SaveAs Filename:=newPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    End If

    ' Sprawdź czy wystąpił błąd (np. użytkownik kliknął Anuluj)
    If Err.Number <> 0 Then
        ' Ignoruj błąd - użytkownik prawdopodobnie kliknął Anuluj
        Err.Clear
    End If

    On Error GoTo 0
End Sub
