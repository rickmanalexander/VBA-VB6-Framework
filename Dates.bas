Attribute VB_Name = "Dates"
'@Folder("Framework.CommonMethods")

Option Explicit
Option Private Module

Public Function IsDateInCurrentMonth(ByVal dateIn As Date) As Boolean
    IsDateInCurrentMonth = (Month(dateIn) = Month(Date))
End Function

Public Function FirstOfMonth(ByVal dateIn As Date) As Date
    FirstOfMonth = DateSerial(Year(dateIn), Month(dateIn), 1)
End Function

Public Function EndOfMonth(ByVal dateIn As Date) As Date
    EndOfMonth = DateSerial(Year(dateIn), Month(dateIn) + 1, 0)
End Function

Public Function GetPrevMonthDate(ByVal dateIn As Date) As Date
    GetPrevMonthDate = DateAdd("m", -1, dateIn)
End Function
