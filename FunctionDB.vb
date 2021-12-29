'-----------------------------
' БД 
Public Function DBConnect(StrCon)
'DBConnect - подключение к базе данных
'On Error Resume Next  

    set oConn = CreateObject("ADODB.Connection")
    oConn.CommandTimeout = 0
    oConn.CursorLocation = 3
    oConn.Open Strcon
    DBConnect = True
    
    IF Err.Number <> 0 THEN 
        DBConnect = False
        For Each ADOErr In oConn.Errors     
            WriteLogErr  "Error.Description: " & ADOErr.Description  & vbCrLf & _
                         "Error.Number: "      & ADOErr.Number       & vbCrLf & _
                         "Error.Source: "      & ADOErr.Source       & vbCrLf & _
                         "Error.SQLState: "    & ADOErr.SQLState     & vbCrLf & _
                         "Error.NativeError: " & ADOErr.NativeError   
        NEXT		
    END IF     
'On Error Goto 0    
End Function

'-----------------------------
Public Function DBdisConnect
'DBdisConnect - закрытие соединения к базе данных
    oRst.close
    oConn.close

    set oRst  = nothing
    set oConn = nothing
End Function

'-----------------------------
Public Function DBExecute(SqlText)
'DBExecute - выполнение запроса
On Error Resume Next
Err.clear  
  dim ADOErr
    DBExecute = True
	
	WriteLog "DBExecute"
    WriteLog "SQL Query: " & vbCrLf & SqlText
	IF len(trim(SqlText)) = 0 THEN EXIT FUNCTION
    
    oConn.Execute (SqlText)
   
	IF Err.Number <> 0 THEN 
	    DBExecute = False
        
        For Each ADOErr In oConn.Errors     
            WriteLogErr "Error.Function: DBExecute" & vbCrLf & _
    		            "Error.Description: " & ADOErr.Description  & vbCrLf & _
                        "Error.Number: "      & ADOErr.Number       & vbCrLf & _
                        "Error.Source: "      & ADOErr.Source       & vbCrLf & _
                        "Error.SQLState: "    & ADOErr.SQLState     & vbCrLf & _
                        "Error.NativeError: " & ADOErr.NativeError  
        NEXT 
     
        'If Err.Number = -2147217908 then DBdisConnect: WScript.Quit
    END IF 
Err.clear    
On Error GoTo 0 	
End Function
