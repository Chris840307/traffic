<%
function haiwaocde(zfstr)
 zf = lcase("*"&zfstr&"*")
 zf = replace(zf,"0","_|_|__||_||_|")
 zf = replace(zf,"1","_||_|__|_|_||")
 zf = replace(zf,"2","_|_||__|_|_||")
 zf = replace(zf,"3","_||_||__|_|_|")
 zf = replace(zf,"4","_|_|__||_|_||")
 zf = replace(zf,"5","_||_|__||_|_|")
 zf = replace(zf,"7","_|_|__|_||_||")
 zf = replace(zf,"6","_|_||__||_|_|")
 zf = replace(zf,"8","_||_|__|_||_|")
 zf = replace(zf,"9","_|_||__|_||_|")
 zf = replace(zf,"a","_||_|_|__|_||")
 zf = replace(zf,"b","_|_||_|__|_||")
 zf = replace(zf,"c","_||_||_|__|_|")
 zf = replace(zf,"d","_|_|_||__|_||")
 zf = replace(zf,"e","_||_|_||__|_|")
 zf = replace(zf,"f","_|_||_||__|_|")
 zf = replace(zf,"g","_|_|_|__||_||")
 zf = replace(zf,"h","_||_|_|__||_|")
 zf = replace(zf,"i","_|_||_|__||_|")
 zf = replace(zf,"j","_|_|_||__||_|")
 zf = replace(zf,"k","_||_|_|_|__||")
 zf = replace(zf,"l","_|_||_|_|__||")
 zf = replace(zf,"m","_||_||_|_|__|")
 zf = replace(zf,"n","_|_|_||_|__||")
 zf = replace(zf,"o","_||_|_||_|__|")
 zf = replace(zf,"p","_|_||_||_|__|")
 zf = replace(zf,"r","_||_|_|_||__|")
 zf = replace(zf,"q","_|_|_|_||__||")
 zf = replace(zf,"s","_|_||_|_||__|")
 zf = replace(zf,"t","_|_|_||_||__|")
 zf = replace(zf,"u","_||__|_|_|_||")
 zf = replace(zf,"v","_|__||_|_|_||")
 zf = replace(zf,"w","_||__||_|_|_|")
 zf = replace(zf,"x","_|__|_||_|_||")
 zf = replace(zf,"y","_||__|_||_|_|")
 zf = replace(zf,"z","_|__||_||_|_|")
 zf = replace(zf,"-","_|__|_|_||_||")
 zf = replace(zf,"*","_|__|_||_||_|")
 zf = replace(zf,"/","_|__|__|_|__|")
 zf = replace(zf,"%","_|_|__|__|__|")
 zf = replace(zf,"+","_|__|_|__|__|")
 zf = replace(zf,".","_||__|_|_||_|")
 haiwaocde = zf
end function
function dragcode(ccode)
 code_H = 20
 code_W = 1
 c = ccode
 c = replace(c,"_","<span style='height:"&code_H&";width:"&code_w&";background:#FFFFFF'></span>")
 c = replace(c,"|","<span style='height:"&code_H&";width:"&code_w&";background:#000000'></span>")
 dragcode = c
end function
function dragtext(ccode)
 c = ccode
 dragtext = ""
 for i=1 to len(c)
 dragtext = dragtext&"<span style='width:26;text-align:center'>"&mid(c,i,1)&"</span>"
 next
 dragtext = dragtext
end function

Function CheckExp(patrn,str)
Set regEx=New RegExp
regEx.Pattern=patrn
regEx.IgnoreCase=true
regEx.Global=True
CheckExp = regEx.test(str)  
End Function
%>