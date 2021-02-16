
Dim A
set A= createobject("shell.Application")
'For Firefox.exe;IE=Iexplore.exe; google chrome= Chrome.exe;.....
A.shellExecute "Chrome.exe","http://www.facebook.com","","",1
A.Sendkeys"{ENTER}"
WSCript.sleep(1500)
A.Sendkeys"(skriyang@yahoo.com)"
WSCript.sleep(1000)
A.Sendkeys"(Skr321usa)"
A.Sendkeys"{ENTER}"