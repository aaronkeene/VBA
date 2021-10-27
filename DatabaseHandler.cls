Option Explicit

' --------------------------------------------- '
' Enums
' --------------------------------------------- '

Public Enum DatabaseType
    MicrosoftExcel = 0
    MicrosoftAccess = 1
End Enum

' --------------------------------------------- '
' Constants
' --------------------------------------------- '

  Const dbExcelProvider As String = "Microsoft.ACE.OLEDB.16.0"
  Const dbExcelExtPropJet = "Excel 12.0"
  Const dbExcelExtPropHeaders As String = "Yes"
  Const dbExcelExtPropIMEX As String = "1"
  
  Const dbAccessProvider As String = "Microsoft.ACE.OLEDB.12.0"

' --------------------------------------------- '
' Public Variables
' --------------------------------------------- '

  Public dbType                 As DatabaseType
  Public dbSource               As String
  Public dbReadOnly             As Boolean

' --------------------------------------------- '
' Private Variables
' --------------------------------------------- '

  Private dbConnection            As ADODB.connection



Private Sub ConnectDatabase(dbType As DatabaseType, Optional dbReadOnly As Boolean = True)
    
    Dim dbConnProp As String
    
    Set dbConnection = New ADODB.connection
    
    Select Case dbType
    
        Case Is = DatabaseType.MicrosoftExcel
            
            dbConnProp = dbExcelExtPropJet & ";"
            dbConnProp = dbConnProp & "HDR=" & dbExcelExtPropHeaders & ";"
            dbConnProp = dbConnProp & "IMEX=" & dbExcelExtPropIMEX & ";"
            
            If Not dbReadOnly Then: dbConnProp = dbConnProp & "ReadOnly=0;"
            
            dbConnection.Provider = dbExcelProvider
            dbConnection.Properties("Extended Properties").value = dbConnProp
            dbConnection.Open Me.dbSource
            
            Exit Sub
            
            
        Case Is = DatabaseType.MicrosoftAccess
            
            dbConnection.Provider = dbAccessProvider
            dbConnection.Open Me.dbSource
            
            Exit Sub
    
    End Select
        
End Sub



Private Sub DisconnectDatabase()
    
    If Not dbConnection Is Nothing Then
        If Not (dbConnection.State = 0) Then
            dbConnection.Close
        End If
    End If
    
    Set dbConnection = Nothing
        
End Sub



Public Function ValidConnection() As Boolean
    
    Call ConnectDatabase(Me.dbType, Me.dbReadOnly)
    
    If Not dbConnection Is Nothing Then
        ValidConnection = True
        Call DisconnectDatabase
        Exit Function
    End If
    
    MsgBox "No database connection. Download a new copy of the tool before proceeding.", vbCritical + vbOKOnly
    
End Function



Public Function GetRecordSet(sqlquery As String, Optional ReadOnly As Boolean = True) As ADODB.RecordSet

    ' Returns disconnected recordset for read only queries and a connected recordset for update queries
    
    ' Exit if query is null
    If sqlquery = vbNullString Then
        Exit Function
    End If
    
    Set GetRecordSet = New ADODB.RecordSet
    
    Call ConnectDatabase(Me.dbType, ReadOnly)
    
    GetRecordSet.CursorLocation = adUseClient
    GetRecordSet.CursorType = adOpenStatic
    GetRecordSet.LockType = adLockBatchOptimistic
    GetRecordSet.Open sqlquery, dbConnection
    
    If Not ReadOnly Then: Exit Function
        
    GetRecordSet.ActiveConnection = Nothing
    
    Call DisconnectDatabase
    
End Function
    


Public Function SelectSingleValue(FieldName As String, TableName As String, Optional Filter As String = vbNullString) As String

    Dim sqlquery                As String
    Dim rst                     As New ADODB.RecordSet
    
    
    ' Format query correctly for Excel file queries
    Select Case dbType
        Case Is = DatabaseType.MicrosoftExcel
            sqlquery = "SELECT [" & FieldName & "] FROM [" & TableName & "$]"
    
        Case Is = DatabaseType.MicrosoftAccess
            sqlquery = "SELECT [" & FieldName & "] FROM [" & TableName & "]"
    
    End Select
    
    If Not Filter = vbNullString Then
        sqlquery = sqlquery & " WHERE " & Filter & ";"
    End If
    
    Call ConnectDatabase(Me.dbType)

    rst.Open sqlquery, dbConnection
    
    If IsNull(rst.Fields(FieldName)) Then
        SelectSingleValue = vbNullString
        rst.Close
        Set rst = Nothing
        Exit Function
    End If
    
    SelectSingleValue = rst.Fields(FieldName).value
    
    rst.Close
    Set rst = Nothing
    
End Function



Public Function SelectMaxValue(FieldName As String, TableName As String) As String

    Dim sqlquery                As String
    Dim rst                     As New ADODB.RecordSet
       
    ' Format query correctly for Excel file queries
    Select Case dbType
        Case Is = DatabaseType.MicrosoftExcel
            sqlquery = "SELECT MAX ([" & FieldName & "]) as MAX" & FieldName & " FROM [" & TableName & "$]"
    
        Case Is = DatabaseType.MicrosoftAccess
            sqlquery = "SELECT MAX ([" & FieldName & "]) as MAX" & FieldName & "  FROM [" & TableName & "]"
    
    End Select
    
    Call ConnectDatabase(Me.dbType)
    
    rst.Open sqlquery, dbConnection
    
    If Not IsNull(rst.Fields("MAX" & FieldName)) Then
        SelectMaxValue = rst.Fields("MAX" & FieldName).value
        rst.Close
        Set rst = Nothing
        Exit Function
    End If
    
    SelectMaxValue = vbNullString
    rst.Close
    Set rst = Nothing
    
End Function



Public Function UpdateStringValue(FieldName As String, TableName As String, NewStringValue As String, Optional Filter As String = "1=1") As String

    Dim sqlquery                As String
    
    
    ' Format query correctly for Excel file queries
    Select Case dbType
        Case Is = DatabaseType.MicrosoftExcel
            sqlquery = "UPDATE [" & TableName & "$] SET [" & FieldName & "]='" & NewValue & "' WHERE " & Filter
    
        Case Is = DatabaseType.MicrosoftAccess
            sqlquery = "UPDATE [" & TableName & "] SET [" & FieldName & "]='" & NewValue & "' WHERE " & Filter
    
    End Select
      
    Call ConnectDatabase(Me.dbType, dbReadOnly = False)
    
    dbConnection.Execute sqlquery
    
    Call DisconnectDatabase
    
End Function



Public Function UpdateLongValue(FieldName As String, TableName As String, NewLongValue As Long, Optional Filter As String = "1=1") As String

    Dim sqlquery                As String
    
    
    ' Format query correctly for Excel file queries
    Select Case dbType
        Case Is = DatabaseType.MicrosoftExcel
            sqlquery = "UPDATE [" & TableName & "$] SET [" & FieldName & "]=" & CLng(NewValue) & " WHERE " & Filter
    
        Case Is = DatabaseType.MicrosoftAccess
            sqlquery = "UPDATE [" & TableName & "] SET [" & FieldName & "]=" & CLng(NewValue) & " WHERE " & Filter
    
    End Select
    
    Call ConnectDatabase(Me.dbType, dbReadOnly = False)
    
    dbConnection.Execute sqlquery
    
    Call DisconnectDatabase
        
End Function



Public Sub ExecuteSql(query As String)
   
    Call ConnectDatabase(Me.dbType, dbReadOnly = False)
    
    dbConnection.Execute UpdateQuery
    
    Call DisconnectDatabase
    
End Sub
