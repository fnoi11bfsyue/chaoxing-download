'说明，计数器文件在1.txt中，网址文件在2.txt中
Dim WshShell, FileObj, TextObj, OpenNo
Dim i, url
'第一步，打开计数器文件，获取计数值，并更新
Set FileObj = CreateObject("Scripting.FileSystemObject")
Set TextObj = FileObj.OpenTextFile("1.txt", 1, False)
if TextObj.AtEndOfStream then
  OpenNo = 0
else
  OpenNo = TextObj.ReadLine
End if
OpenNo = OpenNo + 1
TextObj.Close
Set TextObj = FileObj.OpenTextFile("1.txt", 2, True)
TextObj.WriteLine OpenNo
TextObj.Close
OpenNo = OpenNo - 1
'msgbox "这是第" & openno & "次双击"
'第二步，打开网址文件，打开10个网页
Set WshShell = WScript.CreateObject("WScript.Shell")
Set TextObj = FileObj.OpenTextFile("2.txt", 1, False)
'跳过前面双击时已经打开的网页
for i=1 to OpenNo * 10
  if TextObj.AtEndOfStream then exit for
  url = TextObj.ReadLine
next
'开始打开网址
for i=1 to 10
  if TextObj.AtEndOfStream then exit for
  url = TextObj.ReadLine
  'msgbox "即将打开的网址是：" & url
  WshShell.Run url
next
TextObj.Close
'第三步、判断网址文件是否已经读完，如果完了就修改打e799bee5baa6e58685e5aeb931333335316533开次数为0
if i<10 then
  Set TextObj = FileObj.OpenTextFile("1.txt", 2, True)
  TextObj.WriteLine 0
  TextObj.Close
end if