<!-- #include file="AllFunction.inc"-->
<%
Dim SolarMonth
SolarMonth = Array(31,28,31,30,31,30,31,31,30,31,30,31)
Function SolarDays(y,m)
   SolarDays = SolarMonth(m-1)
   If m = 2 Then
      If (y Mod 4 = 0) And (y Mod 100 <> 0) Or (y Mod 400 = 0) Then SolarDays = 29
   End If
End Function

Function selectlist(s,e,v) 
     selectlist = "<select NAME=TZ style=font-size=9pt; onchange='changeyear()'>"
	  for i=cint(s) to cint(e)
	    if i = cint(v) then
         selectlist = selectlist & "<OPTION VALUE=" & i & " selected>" & Right("0" & i - 1911,3) & vbCRLF  
		else
		 selectlist = selectlist & "<OPTION VALUE=" & i & ">" & Right("0" & i - 1911,3)   & vbCRLF  
        end if 
      next
     selectlist = selectlist & "</select>"
End Function

Function Daylist(StartDay,MonthDay) 
     Dim Start
	 Start = cint(StartDay)
     Daylist = "<Tr bgcolor=#FFCC99>"
     if StartDay <> 0 then
       for j = 0 to (Start-1)
          Daylist = Daylist & "<TD>&nbsp;</TD>" 
       next
	 end if 
	 for i=1 to cint(MonthDay)
	   if ((i+Start-1) mod 7) = 6 then
		  if i = MonthDay then
             Daylist = Daylist & "<TD align=center><A HREF=javascript:setDate("& i &") >" & i &"</A></TD>" & vbCRLF  
          else
		     Daylist = Daylist & "<TD align=center><A HREF=javascript:setDate("& i &") >" & i &"</A></TD></tr>" & vbCRLF & "<Tr bgcolor=#FFCC99>"
		  end if
	   else
          Daylist = Daylist & "<TD align=center><A HREF=javascript:setDate("& i &") >" & i &"</A></TD>"
	   end if
		  x = ((i+Start-1) mod 7)
	 next
	     
     for j = 1 to (6 - cint(x))
          Daylist = Daylist & "<TD>&nbsp;</TD>" 
     next

     Daylist = Daylist & "</tr>" & vbCRLF 
End Function

Function DayOfWeek(y,m,d)
        DayOfWeek = (((3*y-(7*(y+(m+9)\12))\4+(23*m)\9+d+17-((y+(m<3))\100+1)*3\4) Mod 7))
End Function

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<script language=javascript src='../js/date.js'></script>
<script language="javascript">
  var y,m,ClickName
function CheckField(x) {

  var f = document.forms["search"];
   
   m = f.elements['month_h'].value*1;
   y = f.elements['Year_h'].value*1;
   ClickName = f.elements['ClickName1'].value;
   if (x == "next")
     {
	    if(m == 12)
		 {
		   m = 1;
		   y = y + 1;
         }
		 else 
		  m = m + 1
	 }
   if(x == "prev") 
    {
	    if(m == 1)
		  {
		    m = 12;
		    y = y - 1;
		    if(y<= 0)
		      y =1;
          }
        else 
		  m = m - 1


	}
  return true
}


function onWhich(x) {
   var f = document.forms["search"];
   
  switch (x) {
     case "next":
	            CheckField(x);
                document.forms["search"].action="date.asp?m=" + m + "&y=" + y +"&ClickName1=" +  ClickName
                break;
     case "prev":
	            CheckField(x);
                document.forms["search"].action="date.asp?m=" + m + "&y=" + y +"&ClickName1=" +  ClickName
                break;
  }
  document.forms["search"].submit();
}
function changeyear()
{
   var m = document.forms["search"].elements['month_h'].value*1;
   var y = document.forms["search"].TZ.value;
   var ClickName = document.forms["search"].elements['ClickName1'].value;
   document.forms["search"].action="date.asp?m=" + m + "&y=" + y +"&ClickName1=" +  ClickName;
   document.forms["search"].submit();
}



</script>
<BODY BGCOLOR="#33CCFF" leftmargin=0 topmargin=0  >
<%
   
   if request("m") <> "" then 
     month_num = request("m")
   else 
     month_num = Month(Date) 
   end if    
   if request("y") <> "" then 
     Year_num = request("y")
   else 
     Year_num = Year(Date) 
   end if    
   if request("D") <> "" then 
     Day_num = request("D")
   else 
     Day_num = Day(Date) 
   end if  
   if request("ClickName") <> "" then
     ClickName = request("ClickName") 
   else
     ClickName = request("ClickName1") 
   end if
      
  
   response.write "<form name=""search"" method=post action=""date.asp"">"
   response.write "<input type=hidden  name=ClickName1  id=ClickName1 value='"& ClickName & "'>"
   response.write "<input type=hidden  name=month_h  value="& month_num & ">"
   response.write "<input type=hidden  name=Year_h  value="& Year_num & ">"
   response.write "<input type=hidden  name=Day_h  value="& Day_num & ">"
   response.write "<TABLE bgcolor=#FFdFFF width=100% height=100% ><TR bgcolor=#FFCC99><TD width=270 height=54 COLSPAN=7 align=center><A HREF=""javascript:onWhich('prev')""><IMG SRC=../image/green_prev.gif WIDTH=54 HEIGHT=13 BORDER=0 ALT=""上個月""></A>&nbsp;&nbsp;" & selectlist(Year_num-5,Year_num+5,Year_num)&"&nbsp;&nbsp;年&nbsp;&nbsp;"&month_num&"&nbsp;&nbsp;月"&"&nbsp;&nbsp;<A HREF=""javascript:onWhich('next')""><IMG SRC=../image/green_next.gif WIDTH=54 HEIGHT=13 BORDER=0 ALT=""下個月""></A></TD></tr>"
   response.write "<TR bgcolor=#FFCC99 ><TD>&nbsp;日&nbsp;</TD><TD>&nbsp;一&nbsp;</TD><TD>&nbsp;二&nbsp;</TD><TD>&nbsp;三&nbsp;</TD><TD>&nbsp;四&nbsp;</TD><TD>&nbsp;五&nbsp;</TD><TD>&nbsp;六&nbsp;</TD></tr>"
   response.write Daylist(DayOfWeek(Year_num,month_num,1),SolarDays(Year_num,month_num))  
   response.write "<tr  bgcolor=#FFCC99><td COLSPAN=7 align=left>&nbsp;&nbsp;&nbsp;今天是&nbsp;<font color=red><B>" & gInitDT(Date()) & "</B></font></td></tr>" 
   response.write "</TABLE>"
   response.write "</from>"

%>
</BODY>
</HTML>
