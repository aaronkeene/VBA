Option Explicit
Option Private Module


  Private Function FolderPathExists(folderPath As String) As Boolean
    ' Returns true if file/folder exists
    
    On Error GoTo errHandler

    If Dir(folderPath, vbDirectory) <> "" Then
        FolderPathExists = True
        Exit Function
    End If

errHandler:
    'Expected error if file/folder not found.
    If Err.Number = 52 Then
        Err.Clear
        Exit Function
    End If
    
End Function


Private Function GetArrayValue(ByRef arr2D As Variant _
                             , searchString As String _
                             , Optional ArrayColumnToSearch As Long = 0 _
                             , Optional ArrayColumnToReturn As Long = 1) As String
  
    ' Function expects a 2D array with format of searchValue; returnValue
  
    Dim i As Long
    
    For i = LBound(arr2D, 2) To UBound(arr2D, 2)
        '// if setting matches setting text in column 0
        If StrComp(UCase(searchString), UCase(arr2D(ArrayColumnToSearch, i)), vbBinaryCompare) = 0 Then
            If Not IsNull(arr2D(ArrayColumnToReturn, i)) Then
                GetArrayValue = arr2D(ArrayColumnToReturn, i)
            End If
            Exit Function
        End If
    Next i
End Function


Public Function GetXPathElementText(ByVal doc As DOMDocument60, ByVal Namespace As String, ByVal RootNodeName As String, ByVal xpath As String) As String
    
    On Error GoTo errHandler
    
    Dim rootNode As IXMLDOMNode
    Dim node As IXMLDOMNode
    Dim nodeList As IXMLDOMNodeList
        
    doc.SetProperty "SelectionNamespaces", Namespace

    Set rootNode = doc.SelectSingleNode(RootNodeName)
    Set node = rootNode.SelectSingleNode(xpath)
    
    If Not node Is Nothing Then
        GetXPathElementText = node.Text
        Exit Function
    End If
    
errHandler:
    GetXPathElementText = ""
    
End Function


Public Function GetResponse(request As DOMDocument60, Url As String) As DOMDocument60
    ' Sends xml request to the specifed url and returns the xml document
    
    On Error GoTo ErrConnection
  
    Set GetResponse = New DOMDocument60
    Dim server As New MSXML2.XMLHTTP60
 
    server.Open "POST", Url, False
    server.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    server.send request
    
    If server.responseText <> vbNullString Then
        GetResponse.LoadXML server.responseText
        Exit Function
    End If
    
ErrConnection:
  Exit Function
  
End Function





Public Sub GoFaster(Optional ByVal opt As Boolean = False)
      
    'Optional settings used to speed up macro execution
        
    With Application
        If opt = False Then
            .Calculation = xlCalculationAutomatic
            '.Cursor = xlDefault
            '.DisplayAlerts = True
            '.DisplayStatusBar = True
            '.EnableAnimations = True
            '.EnableEvents = True
            .ScreenUpdating = True
            Exit Sub
        End If
        If opt = True Then
            .Calculation = xlCalculationManual
            '.Cursor = xlDefault
            '.DisplayAlerts = False
            '.DisplayStatusBar = False
            '.EnableAnimations = False
            '.EnableEvents = False
            .ScreenUpdating = False
        End If
  End With
End Sub


Private Function PrettyPrintXML(xml As String) As String

    '******************************************************************************
    ' Name:         PrettyPrintXML
    ' Description:  Formats the xml string to include proper indents
    ' Author:
    ' Contributors:
    ' Last Updated: 12/30/2015
    ' References:   http://stackoverflow.com/questions/1118576/how-can-i-pretty-print-xml-source-using-vb6-and-msxml
    ' Dependencies: None
    ' Known Issues: None
    '******************************************************************************

    Dim Reader                  As New SAXXMLReader60
    Dim Writer                  As New MXXMLWriter60

    Writer.indent = True
    Writer.standalone = False
    Writer.omitXMLDeclaration = False
    Writer.Encoding = "utf-8"

    Set Reader.contentHandler = Writer
    Set Reader.dtdHandler = Writer
    Set Reader.errorHandler = Writer
  
    Call Reader.putProperty("http://xml.org/sax/properties/declaration-handler", Writer)
    Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", Writer)
    Call Reader.Parse(xml)
  
    PrettyPrintXML = Writer.output

End Function



Public Sub SaveLogs(FileName As String _
                   , RequestXML As String _
                   , ResponseXML As String _
                   , Optional PrettyPrint As Boolean = True _
                   , Optional ReplaceFile As Boolean = True)
      
    '******************************************************************************
    ' Name:         SaveLogs
    ' Description:  Creates debug file with request and response XML
    ' Author:       Aaron Keene
    ' Contributors:
    ' Last Updated: 12/16/2015
    ' References:
    ' Dependencies: Microsoft Scripting Runtime; PrettyPrintXML function
    ' Limitations:  None
    '******************************************************************************
                             
  Dim stringPosition As Long
  
  '// output file should be a txt file type
  If Not Right(FileName, 4) = ".txt" Then
    
    If InStr(1, FileName, ".") Then
    
      FileName = Left(FileName, InStr(1, FileName, ".") - 1) & ".txt"
    
    Else
      FileName = FileName & ".txt"
    
    End If
  End If
   
  'Replace file if it exists
  If ReplaceFile And Dir(FileName) <> "" Then
    Kill (FileName)
   ' Application.Wait Now + ((1 / 86400) * 2)
  End If
                 
  'Make the xml output pretty (indented)
  If PrettyPrint Then
    If Not RequestXML = "" Then: RequestXML = PrettyPrintXML(RequestXML)
    If Not ResponseXML = "" Then: ResponseXML = PrettyPrintXML(ResponseXML)
  End If
  
  'Write the data to the file
  Open FileName For Append As #1
  Print #1, RequestXML
  Print #1, vbNewLine, vbNewLine
  Print #1, ResponseXML
  Close #1
  
exitMe:
  Exit Sub
  
errHandler:
  Resume exitMe
                        
End Sub


                                
Public Function ShowMessage(MessageName As String, Optional ModuleName As String = vbNullString, Optional AdditionalInfoText As String = vbNullString)
    
    ' Queries the error table and returns message to user.
    
    Dim sqlQuery                  As String
    Dim rst                       As New ADODB.Recordset
    Dim appName                   As String
    Dim appVersion                As String
    
   
    If AdditionalInfoText = vbNullString Then
        AdditionalInfoText = "Not available"
    End If
    
    If ModuleName = vbNullString Then
        ModuleName = "Unknown module"
    End If
        
    appName = DatabaseHandler.SelectSingleValue("application_name", "cfgAppConfig", "[id]=1")
    appVersion = DatabaseHandler.SelectSingleValue("application_version", "cfgAppConfig", "[id]=1")
    
    sqlQuery = "SELECT * FROM [" & tblErrorMessages & "] WHERE [error_name]='" & MessageName & "'"
    
    Set rst = DatabaseHandler.GetRecordSet(sqlQuery)
    
    If rst.BOF And rst.EOF Then
        Exit Function
    End If
  
    If rst.Fields("error_class") = "Info" Then
       MsgBox rst.Fields("error_text"), vbOKOnly, appName & "  v" & appVersion
       Exit Function
    End If
  
    MsgBox rst.Fields("error_text") & vbNewLine & vbNewLine & _
        "Error Module:" & vbTab & ModuleName & vbNewLine & _
        "Additional Info:" & vbTab & AdditionalInfoText & vbNewLine _
        , vbOKOnly + vbCritical _
        , appName & "  v" & appVersion
    
    ' Reset excel options if an error is encountered
    Call Utilities.GoFaster(False)
    
    Exit Function
            
End Function


Public Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    
    'Return true if sht is not nothing
    WorksheetExists = Not sht Is Nothing

End Function


Public Function FindLastRow(Optional ColumnNumber As Long = 1) As Long

    With ActiveSheet
        FindLastRow = .Cells(.Rows.Count, ColumnNumber).End(xlUp).Row
    End With

End Function


Public Function FindLastCol(Optional RowNumber As Long = 1) As Long
    
    With ActiveSheet
        FindLastCol = .Cells(RowNumber, .Columns.Count).End(xlToLeft).Column
    End With

End Function


Public Function CleanString(InputString As String) As String

    Dim oRegEx                  As Object
    Dim replaceChar             As String

    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.MultiLine = True
    oRegEx.IgnoreCase = False
    oRegEx.Pattern = "[\\/:?<>|\*""]"
  
    CleanString = Replace(oRegEx.Replace(InputString, "_"), Chr(10), "_")
    
End Function


Public Sub DeleteWorksheet(WorksheetName As String)
    
    Dim ws As Worksheet
    
    If WorksheetName = vbNullString Then
        Exit Sub
    End If
    
    For Each ws In Worksheets
        If ws.Name = WorksheetName Then
            Application.DisplayAlerts = False
            Sheets(WorksheetName).Delete
            Application.DisplayAlerts = True
            Exit Sub
        End If
    Next
End Sub

