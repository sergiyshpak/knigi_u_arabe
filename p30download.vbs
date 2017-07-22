URL1="http://p30download.com/fa/entry/"

start_num=67000
end_num=73000

htmlName="res_"+CStr(start_num)+"_"+CStr(end_num)+".html"

On Error Resume Next
 
set xmlhttp = createobject ("msxml2.xmlhttp.3.0")


Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile(htmlName,True,True)  

resFile.write ("<html><head><script src=sorttable.js></script></head><body><table  border=1 class=sortable>" & vbCrLf)
resFile.write ("<tr> <th>title</th> <th>number</th> <th>link</th>  </tr> " & vbCrLf)

for j = start_num to end_num

	URL=URL1+CStr(j)

	'MsgBox URL
	xmlhttp.open "get", URL, false
	xmlhttp.send
	MyText= xmlhttp.responseText

	startpos=1

    fromPosSt=InStr(startpos, MyText,"<title>")
	'MsgBox fromPosSt
    fromPosEnd=InStr(fromPosSt, MyText,"</title>")
    fromStr=Mid(MyText, fromPosSt+7, fromPosEnd-fromPosSt)

	checkStr=Mid(MyText, fromPosSt+12, 3)
	'MsgBox checkStr
	
	if checkStr<>"404" then
	  	resFile.write ("<tr> <td>"&fromStr &"</td> <td>"&CStr(j) &"</td> <td>"&URL &"</td>  </tr> " & vbCrLf)
	end if 
	
	WScript.Sleep 650
	
next

resFile.write ("</table></body></html>")
resFile.Close


set shell = WScript.CreateObject("WScript.Shell")
shell.Run "cmd /c  start " + htmlName

