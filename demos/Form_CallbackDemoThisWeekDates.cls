VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CallbackDemoThisWeekDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()

    Me!WeekdayDate.Value = Date
    Me!ThisWeekDates.Value = Date
    
End Sub

Private Sub ThisWeekDates_AfterUpdate()

    Me!WeekdayDate.Value = Me!ThisWeekDates.Value
    
End Sub
