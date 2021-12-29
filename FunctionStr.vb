
Public Function getSubStr(pS, pSymbolBegin, pSymbolEnd, pStep)  
'getSubStr - Получение подстроки
'   pStep - номер вхождения
On Error Resume Next
  getSubStr = Split(Split(pS, pSymbolBegin)(pStep), pSymbolEnd)(0)
err.Clear  
End Function

Private Function getSubCount(pStr, pSubStr)
'getSubCount - Количество вхождений строки в подстроку
dim i 
dim k 
  i=0:k=0  
  
  i= InStr(1, pStr, pSubStr)
  do while i > 0 
       k = k + 1  
       i = InStr(i+1, pStr, pSubStr)
  loop     
  getSubCount = k
End Function
