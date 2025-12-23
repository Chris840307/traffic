<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/OlddbAccess.ini"-->
<!--#include virtual="Traffic/Common/AllFunction_oldData.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<%
    SelUnitList=""
    if request("DB_Selt")="Selt" then
        totalCnt = Trim(Request("totalCnt"))
       'response.Write totalCnt      
       'response.End
        ListLength = Int(Request("checkedNum"))
        For i = 1 To Int(totalCnt)
            fldName = "item_" & i
            strValue = trim(Request(fldName))
            if strValue <> "" then         	  
                strText = trim(Request("text_" & i))
                if SelUnitList="" then
                    SelUnitList=strValue & "_" & strText 
                else
                    SelUnitList=SelUnitList & "," & strValue & "_" & strText 
                end if  
                j = j + 1
            end if
        Next
	response.Write "<script language=""JavaScript"">" & CHR(10) 
	response.Write "opener.myForm.UnitList.value='" & SelUnitList & "'" & CHR(10) 
	response.Write "opener.myForm.submit();" & CHR(10)
   	response.Write "window.close();" & CHR(10) 
	response.Write "</script>" & CHR(10)	
   end if

   'response.Write SelUnitList
   'response.End
   
   ' 
   

 %>
<script language="JavaScript">

    function setItemValue(obj,str){
	    var tmpCounter;
	    if (obj.checked==true){
		    tmpCounter = eval(document.myForm2.checkedNum.value) + 1 ;
		    document.myForm2.checkedNum.value = tmpCounter ;
	      obj.value = str;
	    }else {
		    tmpCounter = eval(document.myForm2.checkedNum.value) - 1 ;
		    document.myForm2.checkedNum.value = tmpCounter ;
	      obj.value = "";		  
	    }
    }	   
   
   function closeWindow() {
        opener.myForm.UnitList.value=myForm2.UnitList.value;
        opener.myForm.submit();
        window.close();
    }  
   	window.focus();
	
</script>


<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>單位代碼查詢</title>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm2" method="post" onsubmit="return select_street();">  
	    <input type="Hidden" name="UnitList">	
			
		<input type="hidden" name="checkedNum" value="0">		
		<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4">
				    單位列表
				    <input type="button" name="btnSelt" value="加入" onclick="compareUnitList();"/>
				</td>
			</tr>
			<tr bgcolor="#EBFBE3">
			   <td width="1%" align="center">選取</td>
				<td width="15%" align="center">單位代碼</td>
				<td width="25%" align="center">單位名稱</td>
			</tr>
<%
	strProject="select * from accnew order by Acc_no"
	set rsProject=conn.execute(strProject)
	p = 1
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof

%>
			<tr <%lightbarstyle 1 %>>
			    <td><input type="checkbox" name="item_<%=p%>" onClick="setItemValue(this,'<%=rsProject("Acc_no")%>');" ></td>
			    <input type="hidden" name="text_<%=p%>" value="<%=rsProject("acc_nm")%>">					 
				<td align="center"><%=trim(rsProject("Acc_no"))%>&nbsp;</td>
				<td><%=trim(rsProject("acc_nm"))%>&nbsp;</td>				
			</tr>
			
<%	
    p = p + 1
    rsProject.MoveNext
	Wend
	rsProject.close
	set rsProject=nothing
%>
            <INPUT TYPE="HIDDEN" NAME="totalCnt" VALUE="<%=p%>" />
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
					<input type="button" name="close" value="關閉視窗" onclick="window.close();">
				</td>
			</tr>
		</table>
		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
 
<script language="JavaScript">

function Inert_Data(SCode,SStreet){
	<%if Stype="U" then%>
		opener.myForm.BillUnitID.value=SCode;
		opener.Layer6.innerHTML=SStreet;
		opener.TDUnitErrorLog=0;
		window.close();
	<%elseif Stype="S" then%>
		opener.myForm.MemberStation.value=SCode;
		opener.Layer5.innerHTML=SStreet;
		opener.TDStationErrorLog=0;
		window.close();
	<%else%>
		opener.myForm.BillUnitID.value=SCode;
		opener.Layer6.innerHTML=SStreet;
		opener.TDUnitErrorLog=0;
		window.close();
	<%end if%>	
}

function compareUnitList()
{
    myForm2.DB_Selt.value="Selt";
    myForm2.submit(); 
}

<%if request("DB_Selt")="Selt" then%>
<%response.Write "opener.myForm.UnitList.value='" & SelUnitList & "'" & CHR(10) %>
<%response.Write "opener.myForm.submit();" & CHR(10) %>
//<%response.end%>
<%end if %>

</script>

</html>
