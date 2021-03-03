VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CallbackDemoUltimoMonths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub DateRows_AfterUpdate()

    DateRowsAfterUpdate Me!DateRows, Me!UltimoMonthDates

End Sub

Private Sub DateRowsAfterUpdate( _
    ByRef MeDateRows As TextBox, _
    ByRef MeUltimoMonthDates As Control)

    Dim ThisRowCount        As Long
    Dim PreviousRowCount    As Long
    Dim NewRowCount         As Long
    
    Select Case MeUltimoMonthDates.ControlType
        Case acComboBox, acListBox
            ' OK.
        Case Else
            ' Nothing to do.
            Exit Sub
    End Select
        
    ThisRowCount = Nz(MeDateRows.Value, 0)
    PreviousRowCount = MeUltimoMonthDates.ListCount
    
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
        ConfigUltimoMonthDates MeUltimoMonthDates, , NewRowCount
    End If
    
End Sub

Private Sub Form_Load()

    Me!DateRows.Value = Me!UltimoMonthDates.ListCount

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ConfigUltimoMonthDates Me!UltimoMonthDates
    
End Sub

Private Sub UltimoSelect_AfterUpdate()

    Dim StartDate   As Date
    Dim RowCount    As Long
    Dim Enabled     As Boolean
    
    Select Case UltimoSelect.Value
        Case 0
            ' List dates from today.
            StartDate = Date
            ' Allow to adjust the count of months to list.
            RowCount = Me!DateRows.Value
            Enabled = True
        Case 1
            ' List the current year's ultimo month dates.
            StartDate = DateSerial(Year(Date), 1, 1)
            ' Lock row count to the count of months for a year.
            RowCount = MonthsPerYear
            Enabled = False
    End Select
    
    Me!DateRows.Value = RowCount
    Me!DateRows.Enabled = Enabled
    Me!DateRows.Locked = Not Enabled
    
    ConfigUltimoMonthDates Me!UltimoMonthDates, StartDate, RowCount
    
End Sub
