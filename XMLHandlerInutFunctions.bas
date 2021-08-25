Option Explicit
Option Private Module

  ' Module supplements XMLHandler and provides a framework for handeling custom input functions
  ' and data modifications.

Public Function FunctionHandler(customFunctionName As String, inputValue As String) As String
       
    If customFunctionName = vbNullString Then
        FunctionHandler = vbNullString
        Exit Function
    End If
    
    Select Case customFunctionName
        Case Is = "zclConvertDate": FunctionHandler = XMLHandlerInutFunctions.zclConvertDate(CDate(inputValue))

    End Select
    
End Function


Private Function zclConvertDate(inputDate As Date _
                     , Optional lDateRange As Date = #1/1/1900# _
                     , Optional uDateRange As Date = #12/31/2999#) As String

    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
  
    If (inputDate < lDateRange Or inputDate > uDateRange) Then
        inputDate = Date
    End If
  
    strYear = Year(inputDate)
    strMonth = addLeadingZeros(Month(inputDate), 2)
    strDay = addLeadingZeros(Day(inputDate), 2)
  
    zclConvertDate = strYear & "-" & strMonth & "-" & strDay

End Function
