'-----------------------------
PUBLIC FUNCTION FolderExists(SourceFolderName)
'FileUnlocked - проверка наличия файла
'  SourceFolderName - путь к папке
FolderExists = False
If CreateObject("Scripting.FileSystemObject").FolderExists(SourceFolderName)="True" Then 
    FolderExists = True
end if
END FUNCTION

'-----------------------------
PUBLIC FUNCTION WriteLog(text)
'WriteLog   -   Функция логирования
'    
  if writeLog_ then
     call WriteLogging (text)
  end if    
END FUNCTION
'-----------------------------
PUBLIC FUNCTION WriteLogErr(text)
'WriteLog   -   Функция логирования
'    
  if writeLog_ or writeLogErr_ then
     call WriteLogging (text)
  end if    
END FUNCTION

'-----------------------------
PUBLIC FUNCTION WriteLogging(text)
'WriteLogging - Функция логирования
'    
dim oFLog
dim pLogFolder, pLogName

    ' Задаем имя лога
    pLogFolder= ""
    pLogName = "Log\Log_" & environment & "_" & Date
    ' Заменяем в имени все знаки на подчеркивания
    pLogName = Replace(pLogName, ".", "")
    pLogName = Replace(pLogName, ":", "")
    pLogName = Replace(pLogName, "/", "")
    'pLogName = Replace(pLogName, "-", "")
    pLogName = pLogFolder & pLogName
    
    ' Создаем файл
    Set oFLog = CreateObject("Scripting.FileSystemObject").OpenTextFile(pLogName & ".txt",8, true) ' для добавления - 8, тру - пересодавать файл лога если его нет
    
    oFLog.WriteLine  cstr(now) & " - " & text 
    oFLog.Close   
  
    set oFLog = nothing    
END FUNCTION

'-----------------------------
Public sub PrintErr(AErr)
'PrintErr - 
    WriteLogErr "Error.sub: PrintErr" & vbCrLf & _
                "Error.Description: " & AErr.Description  & vbCrLf & _
                "Error.Number: "      & AErr.Number       & vbCrLf & _
                "Error.Source: "      & AErr.Source       & vbCrLf & _
                "Error.HelpContext : "& AErr.HelpContext  & vbCrLf & _
                "Error.HelpFile : "   & AErr.HelpFile  
                
    'If AErr.Number = -2147217908 then DBdisConnect: WScript.Quit                
    AErr.Clear                
End sub