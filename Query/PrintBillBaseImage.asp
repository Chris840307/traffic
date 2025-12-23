<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> ¹H³W¬Û¤ù¦C¦L </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY><%
	response.write "<br>¡@¡@¡@<img src="""&request("ImagePatha")&""" name=""imgB1"" width=""700"">"
	if not ifnull(request("ImagePathb")) then
		response.write "<br>¡@¡@¡@<img src="""&request("ImagePathb")&""" name=""imgB2"" width=""700"">"
	end If
	if not ifnull(request("ImagePathc")) then
		response.write "<br>¡@¡@¡@<img src="""&request("ImagePathc")&""" name=""imgB2"" width=""700"">"
	end If
	if not ifnull(request("ImagePathd")) then
		response.write "<br>¡@¡@¡@<img src="""&request("ImagePathd")&""" name=""imgB2"" width=""700"">"
	end if
%>
</BODY>
<script language="JavaScript">
window.print();
</script>
</HTML>
