Option Explicit
Option Private Module


Private DBConnection                      As ADODB.Connection


Public Function ConnectionExists() As Boolean

    ' returns true if able to connect to the dabase specified. 
    
    Call DatabaseHandler.ConnectDatabase
    
    If Not DBConnection Is Nothing Then
        ConnectionExists = True
        Call DisconnectDatabase
        Exit Function
    End If

    MsgBox "No database connection.", vbCritical + vbOKOnly
    Exit Function

End Function


Private Sub ConnectDatabase(Optional ReadOnly As Boolean = True)
    
    ' creates an active database connection to the active workbook
    
    Dim DbSource                        As String
    Dim DbProperties                    As String
    Dim DbProvider                      As String
    
    DbProvider = "Microsoft.ace.OLEDB.16.0"
    DbSource = ThisWorkbook.FullName
    DbProperties = "Excel 12.0;HDR=Yes;IMEX=1;"
    
    If ReadOnly = False Then
        DbProperties = "Excel 12.0;HDR=Yes;ReadOnly=0;"
    End If
       
    Set DBConnection = New ADODB.Connection
    
    With DBConnection
        .Provider = DbProvider
        .Properties("Extended Properties").value = DbProperties
        .Open DbSource
    End With
        
End Sub
    

Private Sub DisconnectDatabase()
    
    ' disconnects the database
    
    If Not DBConnection Is Nothing Then
        If Not (DBConnection.State = 0) Then
            DBConnection.Close
        End If
    End If
    
    Set DBConnection = Nothing
        
End Sub


Public Function SelectSingleValue(FieldName As String, TableName As String, Optional Filter As String = vbNullString) As String

    ' function returns a single value from the specified table as a string. 
    
    Dim sqlQuery                      As String
    Dim rst                           As New ADODB.RecordSet
    
    sqlQuery = "SELECT [" & FieldName & "] FROM [" & TableName & "$]"
    
    If Filter <> vbNullString Then
        sqlQuery = sqlQuery & " WHERE " & Filter & ";"
    End If
    
    Call DatabaseHandler.ConnectDatabase([ReadOnly] = True)

    rst.Open sqlQuery, DBConnection
    
    If IsNull(rst.Fields(FieldName)) Then
        SelectSingleValue = vbNullString
        rst.Close
        Exit Function
    End If
    
    SelectSingleValue = rst.Fields(FieldName).value
    
    rst.Close
    
    Call DatabaseHandler.DisconnectDatabase
    
End Function


Public Function SelectMaxValue(FieldName As String, TableName As String) As String

    Dim sqlQuery                    As String
    Dim rst                         As New ADODB.RecordSet
    
    sqlQuery = "SELECT MAX([" & FieldName & "]) AS MAX" & FieldName & " FROM [" & TableName & "]"
    
    Call DatabaseHandler.ConnectDatabase([ReadOnly] = True)
    
    rst.Open sqlQuery, DBConnection
    
    If Not IsNull(rst.Fields("MAX" & FieldName)) Then
        SelectMaxValue = rst.Fields("MAX" & FieldName).value
        rst.Close
        Exit Function
    End If
    
    SelectMaxValue = vbNullString
    
    Call DatabaseHandler.DisconnectDatabase
    
End Function


Public Function GetRecordSet(sqlQuery As String) As ADODB.RecordSet 
    
    ' function returns a disconnected recordset to the calling procedure
    
    Set GetRecordSet = New ADODB.RecordSet
    
    Call DatabaseHandler.ConnectDatabase([ReadOnly] = True)
    
    With GetRecordSet
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open sqlQuery, DBConnection
        .ActiveConnection = Nothing
     End With
     
     Call DatabaseHandler.DisconnectDatabase
    
End Function


Public Sub UpdateSingleStringValue(NewValue As String, FieldName As String, TableName As String, Filter As String)

    Dim sqlQuery                    As String
    
    sqlQuery = "UPDATE [" & TableName & "$] SET [" & FieldName & "]='" & NewValue & "' WHERE " & Filter
    
    Call DatabaseHandler.ConnectDatabase([ReadOnly] = False)
    
    DBConnection.Execute sqlQuery
    
    Call DatabaseHandler.DisconnectDatabase
    
End Function


Public Sub UpdateSingleLongValue(NewValue As Long, FieldName As String, TableName As String, Filter As String )

    Dim sqlQuery                    As String
    
    sqlQuery = "UPDATE [" & TableName & "$] SET [" & FieldName & "]=" & CLng(NewValue) & " WHERE " & Filter
    
    Call DatabaseHandler.ConnectDatabase([ReadOnly] = False)
    
    DBConnection.Execute sqlQuery
    
    Call DatabaseHandler.DisconnectDatabase
    
End Function


Public Sub ExecuteSql(SqlString As String)
   
    ' careful with this as it effectively lets you execute any sql agaist the database
    
    Call DatabaseHandler.ConnectDatabase([ReadOnly] = False)
    
    DBConnection.Execute SqlString
    
    Call DatabaseHandler.DisconnectDatabase
    
End Sub

