Attribute VB_Name = "DateTemp"
Option Compare Database
Option Explicit
'
' DateTemp
' Version 1.0.0
'
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Callback
'
' Selected supporting functions for Callback demo from
' modules DateFind and DateText.
' Not to be used if the full modules from project VBA.Date is used.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)


' Common constants.

    Public Const DaysPerWeek            As Long = 7
    Public Const MonthsPerYear          As Integer = 12
'

' Returns first day of week according to the current Windows settings.
'
' 2017-05-03. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function SystemDayOfWeek() As VbDayOfWeek

    Const DateOfSaturday    As Date = #12:00:00 AM#
    
    Dim DayOfWeek   As VbDayOfWeek
    
    DayOfWeek = vbSunday + vbSaturday - Weekday(DateOfSaturday, vbUseSystemDayOfWeek)
    
    SystemDayOfWeek = DayOfWeek
    
End Function

' Returns the date of the weekday as specified by DayOfWeek
' following Date1.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateNextWeekday( _
    ByVal Date1 As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek) _
    As Date

    Dim Interval    As String
    Dim ResultDate  As Date
    
    Interval = "d"
    
    If DayOfWeek = vbUseSystemDayOfWeek Then
        DayOfWeek = Weekday(Date1)
    End If
    
    ResultDate = DateAdd(Interval, DaysPerWeek - (Weekday(Date1, DayOfWeek) - 1), Date1)
    
    DateNextWeekday = ResultDate
    
End Function

' Returns the date of the weekday as specified by DayOfWeek
' preceding Date1.
'
' Note: If DayOfWeek is omitted, the weekday of Date1 is used.
' If so, the date returned will always be Date1 - 7.
'
' 2019-06-29. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DatePreviousWeekday( _
    ByVal Date1 As Date, _
    Optional ByVal DayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek) _
    As Date

    Dim Interval    As String
    Dim ResultDate  As Date
    
    Interval = "d"
    
    If DayOfWeek = vbUseSystemDayOfWeek Then
        DayOfWeek = Weekday(Date1)
    End If
    
    ResultDate = DateAdd(Interval, -DaysPerWeek - ((Weekday(Date1, DayOfWeek) - DaysPerWeek - 1) Mod DaysPerWeek), Date1)
    
    DatePreviousWeekday = ResultDate
    
End Function

' Returns the primo date of the week of the date passed.
'
' 2016-01-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DateThisWeekPrimo( _
    ByVal DateThisWeek As Date, _
    Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday) _
    As Date

    Dim ResultDate  As Date
    
    ResultDate = DateAdd("d", 1 - Weekday(DateThisWeek, FirstDayOfWeek), DateThisWeek)
    
    DateThisWeekPrimo = ResultDate
    
End Function

' Obtain the system date format without API calls.
'
' 2021-01-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatSystemDate() As String

    Const TestDate  As Date = #1/2/3333#
    
    Dim DateFormat  As String
    
    DateFormat = Replace(Replace(Replace(Replace(Replace(Format(TestDate), "3", "y"), "1", "m"), "2", "d"), "0m", "mm"), "0d", "dd")

    FormatSystemDate = DateFormat
    
End Function

' Obtain the system date separator without API calls.
'
' 2021-01-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatSystemDateSeparator() As String

    Dim Separator   As String
    
    Separator = Format(Date, "/")

    FormatSystemDateSeparator = Separator
    
End Function

