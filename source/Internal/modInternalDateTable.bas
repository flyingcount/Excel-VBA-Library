Attribute VB_Name = "modInternalDateTable"
Option Explicit

' Internal: build a calendar / date-dimension array.
' Column layout aligns with Power Query fxGenerateDateTable / fxAddDateParts.

Public Enum DateTableCol
    colDate = 1
    colYear = 2
    colQuarter = 3
    colMonthNumber = 4
    colMonth = 5
    colDay = 6
    colDayOfWeekNumber = 7
    colDayOfWeek = 8
    colWeek = 9
    colMmmYyyy = 10
    colMonthSort = 11
    colFiscalYear = 12
    colFiscalPeriod = 13
    colCount = 13
End Enum

''' Builds a 1-based 2D array: row 1 = headers, then one row per day inclusive.
Public Function BuildDateTableArray( _
    ByVal StartDate As Date, _
    ByVal EndDate As Date, _
    Optional ByVal FiscalYearStartMonth As Long = 1 _
) As Variant

    Dim dayCount As Long
    Dim data As Variant
    Dim r As Long
    Dim d As Date
    Dim monthNum As Long
    Dim fy As Long
    Dim fp As Long

    If EndDate < StartDate Then
        Err.Raise vbObjectError + 1001, "BuildDateTableArray", _
            "EndDate must be on or after StartDate."
    End If

    If FiscalYearStartMonth < 1 Or FiscalYearStartMonth > 12 Then
        Err.Raise vbObjectError + 1002, "BuildDateTableArray", _
            "FiscalYearStartMonth must be between 1 and 12."
    End If

    dayCount = EndDate - StartDate + 1
    ReDim data(1 To dayCount + 1, 1 To DateTableCol.colCount)

    data(1, DateTableCol.colDate) = "Date"
    data(1, DateTableCol.colYear) = "Year"
    data(1, DateTableCol.colQuarter) = "Quarter"
    data(1, DateTableCol.colMonthNumber) = "Month Number"
    data(1, DateTableCol.colMonth) = "Month"
    data(1, DateTableCol.colDay) = "Day"
    data(1, DateTableCol.colDayOfWeekNumber) = "Day Of Week Number"
    data(1, DateTableCol.colDayOfWeek) = "Day Of Week"
    data(1, DateTableCol.colWeek) = "Week"
    data(1, DateTableCol.colMmmYyyy) = "MMM-YYYY"
    data(1, DateTableCol.colMonthSort) = "Month sort"
    data(1, DateTableCol.colFiscalYear) = "Fiscal Year"
    data(1, DateTableCol.colFiscalPeriod) = "Fiscal Period"

    For r = 1 To dayCount
        d = StartDate + (r - 1)
        monthNum = Month(d)

        ' Fiscal year labeled by ending calendar year; period 1 = FiscalYearStartMonth
        If FiscalYearStartMonth = 1 Then
            fy = Year(d)
            fp = monthNum
        ElseIf monthNum >= FiscalYearStartMonth Then
            fy = Year(d) + 1
            fp = monthNum - FiscalYearStartMonth + 1
        Else
            fy = Year(d)
            fp = monthNum + (12 - FiscalYearStartMonth + 1)
        End If

        data(r + 1, DateTableCol.colDate) = d
        data(r + 1, DateTableCol.colYear) = Year(d)
        data(r + 1, DateTableCol.colQuarter) = DatePart("q", d)
        data(r + 1, DateTableCol.colMonthNumber) = monthNum
        data(r + 1, DateTableCol.colMonth) = Format$(d, "mmmm")
        data(r + 1, DateTableCol.colDay) = Day(d)
        ' Monday = 1 ... Sunday = 7 (matches PQ Day.Monday + 1 style)
        data(r + 1, DateTableCol.colDayOfWeekNumber) = Weekday(d, vbMonday)
        data(r + 1, DateTableCol.colDayOfWeek) = Format$(d, "dddd")
        data(r + 1, DateTableCol.colWeek) = DatePart("ww", d, vbMonday, vbFirstJan1)
        data(r + 1, DateTableCol.colMmmYyyy) = Format$(d, "mmm-yyyy")
        data(r + 1, DateTableCol.colMonthSort) = Format$(d, "mm") & " " & Format$(d, "mmm")
        data(r + 1, DateTableCol.colFiscalYear) = fy
        data(r + 1, DateTableCol.colFiscalPeriod) = fp
    Next r

    BuildDateTableArray = data
End Function
