VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CallbackDemoWeekdayDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()

    Const DefaultRowCount   As Integer = 16

    Dim FirstDate           As Date
    Dim RowCount            As Long
    Dim DayOfWeek           As VbDayOfWeek
    
    ' Assign the first day of the week as the default selection.
    Me!Weekdays1.Value = SystemDayOfWeek
    Me!StartDate1.Value = Date
    ' Assign the last day of the week as the default selection.
    Me!Weekdays2.Value = (SystemDayOfWeek - 1 - 1 + DaysPerWeek) Mod DaysPerWeek + 1
    Me!StartDate2.Value = Date
    
    ' Adjust count of rows to be listed if set in property Tag.
    RowCount = Val(Me!WeekdayDates1.Tag)
    If RowCount = 0 Then
        RowCount = DefaultRowCount
    End If
    Me!DateRows1.Value = RowCount
    
    RowCount = Val(Me!WeekdayDates2.Tag)
    If RowCount = 0 Then
        RowCount = DefaultRowCount
    End If
    Me!DateRows2.Value = RowCount
    
    ' Calculate end dates to display.
    FirstDate = DateNextWeekday(DateAdd("d", -1, Me!StartDate1.Value), Me!Weekdays1.Value)
    Me!EndDate1.Value = DateAdd("ww", RowCount - 1, FirstDate)

    FirstDate = DateNextWeekday(DateAdd("d", -1, Me!StartDate2.Value), Me!Weekdays2.Value)
    Me!EndDate2.Value = DateAdd("ww", RowCount - 1, FirstDate)
    
    ' Retrieve and display the default date format from the comboboxes' Format property.
    Me!FormatDate1.DefaultValue = """" & Me!WeekdayDates1.Format & """"
    Me!FormatDate2.DefaultValue = """" & Me!WeekdayDates2.Format & """"
    
    ' Check if default weekday of the weekday selector is different from SystemDayOfWeek.
    DayOfWeek = Me!Weekdays1.Value
    If DayOfWeek <> SystemDayOfWeek Then
        ' Reconfig the list/combobox as the weekday to list is different from SystemDayOfWeek.
        ConfigWeekdayDates Me!WeekdayDates1, DayOfWeek
    End If
    DayOfWeek = Me!Weekdays2.Value
    If DayOfWeek <> SystemDayOfWeek Then
        ' Reconfig the list/combobox as the weekday to list is different from SystemDayOfWeek.
        ConfigWeekdayDates Me!WeekdayDates2, DayOfWeek
    End If

    DayOfWeek = Weekday(Date)
    If DayOfWeek <> SystemDayOfWeek Then
        ' Reconfig the list/combobox as the weekday to list is different from SystemDayOfWeek.
        ConfigWeekdayDates Me!ListWeekdayDates, DayOfWeek
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ConfigWeekdayDates Me!ListWeekdayDates
    ConfigWeekdayDates Me!WeekdayDates1
    ConfigWeekdayDates Me!WeekdayDates2
    
End Sub

Private Sub DateRows1_AfterUpdate()

    DateRowsAfterUpdate Me!Weekdays1, Me!StartDate1, Me!EndDate1, Me!DateRows1, Me!WeekdayDates1
    
End Sub

Private Sub DateRows2_AfterUpdate()

    DateRowsAfterUpdate Me!Weekdays2, Me!StartDate2, Me!EndDate2, Me!DateRows2, Me!WeekdayDates2

End Sub

Private Sub EndDate1_AfterUpdate()

    EndDateAfterUpdate Me!Weekdays1, Me!StartDate1, Me!EndDate1, Me!DateRows1, Me!WeekdayDates1

End Sub

Private Sub EndDate2_AfterUpdate()

    EndDateAfterUpdate Me!Weekdays2, Me!StartDate2, Me!EndDate2, Me!DateRows2, Me!WeekdayDates2

End Sub

Private Sub FormatDate1_AfterUpdate()

    FormatDateAfterUpdate Me!FormatDate1, Me!WeekdayDates1
    
End Sub

Private Sub FormatDate2_AfterUpdate()

    FormatDateAfterUpdate Me!FormatDate2, Me!WeekdayDates2

End Sub

Private Sub StartDate1_AfterUpdate()

    StartDateAfterUpdate Me!Weekdays1, Me!StartDate1, Me!EndDate1, Me!DateRows1, Me!WeekdayDates1

End Sub

Private Sub StartDate2_AfterUpdate()

    StartDateAfterUpdate Me!Weekdays2, Me!StartDate2, Me!EndDate2, Me!DateRows2, Me!WeekdayDates2

End Sub

Private Sub Weekdays1_AfterUpdate()

    WeekdaysAfterUpdate Me!Weekdays1, Me!StartDate1, Me!EndDate1, Me!DateRows1, Me!WeekdayDates1
    
End Sub

Private Sub Weekdays2_AfterUpdate()

    WeekdaysAfterUpdate Me!Weekdays2, Me!StartDate2, Me!EndDate2, Me!DateRows2, Me!WeekdayDates2

End Sub

Private Sub DateRowsAfterUpdate( _
    ByRef MeWeekdays As ComboBox, _
    ByRef MeStartDate As TextBox, _
    ByRef MeEndDate As TextBox, _
    ByRef MeDateRows As TextBox, _
    ByRef MeWeekdayDates As Control)

    Dim ThisRowCount        As Long
    Dim PreviousRowCount    As Long
    Dim NewRowCount         As Long
    Dim FirstDate           As Date
    Dim LastDate            As Date
    
    Select Case MeWeekdayDates.ControlType
        Case acComboBox, acListBox
            ' OK.
        Case Else
            ' Nothing to do.
            Exit Sub
    End Select
        
    ThisRowCount = Nz(MeDateRows.Value, 0)
    PreviousRowCount = MeWeekdayDates.ListCount
    
    If ThisRowCount > MaximumRowCount Then
        NewRowCount = PreviousRowCount
    ElseIf ThisRowCount < MinimumRowCount Then
        NewRowCount = PreviousRowCount
    Else
        NewRowCount = ThisRowCount
    End If
    
    If ThisRowCount <> NewRowCount Then
        ' Adjust the displayed row count.
        MeDateRows.Value = NewRowCount
    End If
    
    If PreviousRowCount <> NewRowCount Then
        ' First date listed.
        FirstDate = DateNextWeekday(DateAdd("d", -1, MeStartDate.Value), MeWeekdays.Value)
    
        ' Check that the end date matches the row count.
        LastDate = DateAdd("ww", NewRowCount - 1, FirstDate)
        If DateDiff("d", MeEndDate.Value, LastDate) <> 0 Then
            ' Adjust the end date.
            MeEndDate.Value = LastDate
        End If
    
        ConfigWeekdayDates MeWeekdayDates, , MeStartDate.Value, NewRowCount
    End If
    
End Sub

Private Sub EndDateAfterUpdate( _
    ByRef MeWeekdays As ComboBox, _
    ByRef MeStartDate As TextBox, _
    ByRef MeEndDate As TextBox, _
    ByRef MeDateRows As TextBox, _
    ByRef MeWeekdayDates As Control)

    Dim StartDate           As Date
    Dim FirstDate           As Date
    Dim LastDate            As Date
    Dim RowCount            As Long
    Dim NewRowCount         As Long
    
    Select Case MeWeekdayDates.ControlType
        Case acComboBox, acListBox
            ' OK.
        Case Else
            ' Nothing to do.
            Exit Sub
    End Select
    
    ' No empty end date.
    If IsNull(MeEndDate.Value) Then
        MeEndDate.Value = MeStartDate.Value
    End If
    
    StartDate = MeStartDate.Value
    RowCount = MeDateRows.Value
    
    FirstDate = DateNextWeekday(DateAdd("d", -1, MeStartDate.Value), MeWeekdays.Value)
    Select Case DateDiff("d", FirstDate, MeEndDate.Value)
        Case Is > 0
            ' Keep the start date as is, and adjust the row count.
            NewRowCount = MinimumRowCount + DateDiff("ww", FirstDate, MeEndDate.Value)
        Case Is < 0
            ' Keep the row count as is, and move start date back.
            LastDate = DatePreviousWeekday(DateAdd("d", 1, MeEndDate.Value), MeWeekdays.Value)
            MeStartDate.Value = DateAdd("ww", 1 - RowCount, LastDate)
            NewRowCount = RowCount
        Case Else
            ' Keep start date as is, set row count to 1, and set end date to the first weekday.
            MeEndDate.Value = FirstDate
            NewRowCount = MinimumRowCount
    End Select
    
    If RowCount <> NewRowCount Or DateDiff("d", StartDate, MeStartDate.Value) <> 0 Then
        MeDateRows.Value = NewRowCount
        ConfigWeekdayDates MeWeekdayDates, , MeStartDate.Value, NewRowCount
    End If
    
End Sub

Private Sub FormatDateAfterUpdate( _
    ByRef MeFormatDate As TextBox, _
    ByRef MeWeekdayDates As Control)

    Dim Format              As String
    Dim Format1             As String
    Dim Format2             As String
    Dim Format3             As String
    Dim Separator           As String
    Dim FormatCode          As Long
    
    Select Case MeWeekdayDates.ControlType
        Case acComboBox, acListBox
            ' OK.
        Case Else
            ' Nothing to do.
            Exit Sub
    End Select
    
    Format = Nz(MeFormatDate.Value, FormatSystemDate)
    ConfigWeekdayDates MeWeekdayDates, , , , Format

End Sub

Private Sub StartDateAfterUpdate( _
    ByRef MeWeekdays As ComboBox, _
    ByRef MeStartDate As TextBox, _
    ByRef MeEndDate As TextBox, _
    ByRef MeDateRows As TextBox, _
    ByRef MeWeekdayDates As Control)

    Dim StartDate           As Date
    Dim FirstDate           As Date
    Dim LastDate            As Date
    
    Select Case MeWeekdayDates.ControlType
        Case acComboBox, acListBox
            ' OK.
        Case Else
            ' Nothing to do.
            Exit Sub
    End Select
    
    StartDate = Nz(MeStartDate.Value, #12:00:00 AM#)
    
    ' No empty start date.
    If IsNull(MeStartDate.Value) Then
        MeStartDate.Value = MeEndDate.Value
    End If

    ' First date listed.
    FirstDate = DateNextWeekday(DateAdd("d", -1, MeStartDate.Value), MeWeekdays.Value)
    
    ' Check that the end date will include the last date listed.
    LastDate = DateAdd("ww", MeDateRows.Value - 1, FirstDate)
    If DateDiff("d", MeEndDate.Value, LastDate) > 0 Then
        ' Increase end date.
        MeEndDate.Value = LastDate
    ElseIf DateDiff("d", MeEndDate.Value, DateAdd("ww", MeDateRows.Value, LastDate)) < 0 Then
        ' Reduce end date.
        MeEndDate.Value = LastDate
    End If

    If DateDiff("d", MeWeekdayDates.ItemData(0), FirstDate) <> 0 Then
        ConfigWeekdayDates MeWeekdayDates, , MeStartDate.Value
    End If

End Sub

Private Sub WeekdaysAfterUpdate( _
    ByRef MeWeekdays As ComboBox, _
    ByRef MeStartDate As TextBox, _
    ByRef MeEndDate As TextBox, _
    ByRef MeDateRows As TextBox, _
    ByRef MeWeekdayDates As Control)

    Dim FirstDate           As Date
    Dim LastDate            As Date
    
    Select Case MeWeekdayDates.ControlType
        Case acComboBox, acListBox
            ' OK.
        Case Else
            ' Nothing to do.
            Exit Sub
    End Select
    
    If IsNull(MeWeekdays.Value) Then
        MeWeekdays.Value = SystemDayOfWeek
    End If
    
    ' Include start date as the possible first date to list.
    FirstDate = DateNextWeekday(DateAdd("d", -1, MeStartDate.Value), MeWeekdays.Value)
    
    ' Check that the end date will include the last date listed.
    LastDate = DateAdd("ww", MeDateRows.Value - 1, FirstDate)
    If DateDiff("d", MeEndDate.Value, LastDate) > 0 Then
        ' Increase end date.
        MeEndDate.Value = LastDate
    End If
    
    ConfigWeekdayDates MeWeekdayDates, MeWeekdays.Value
    
End Sub

