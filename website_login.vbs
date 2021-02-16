
'Dim A
set A=wscript.createobject("wscript.Application")
'For Firefox.exe;IE=Iexplore.exe; google chrome= Chrome.exe;.....
A.run "Chrome.exe"
wscript.sleep(700)
A.Sendkeys "{F6}"
A.Sendkeys"{www.facebook.com}"
A.Sendkeys"{ENTER}"
WSCript.sleep(1500)
A.Sendkeys"(skriyang@yahoo.com)"
WSCript.sleep(1000)
A.Sendkeys"(Skr321usa)"
A.Sendkeys"{ENTER}"
'shellExecute "Chrome.exe","http://www.facebook.com","","",1
