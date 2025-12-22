<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fname= "car.txt"
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"

response.write "1,汽車"&vbnewline
response.write "2,拖車"&vbnewline
response.write "3,重機"&vbnewline
response.write "4,輕機"&vbnewline
response.write "6,臨時車牌"
%>