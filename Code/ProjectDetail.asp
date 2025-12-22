<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
		'權限
           '1:查詢 ,2:新增 ,3:修改 ,4:刪除

           if CheckPermission(226,1)=false then
              iPermission1= "disabled"
           else
              iPermission1= " "
		   end if

           if CheckPermission(226,2)=false then
              iPermission2= "disabled"
           else
              iPermission2= " "
		   end if

           if CheckPermission(226,3)=false then
              iPermission3= "disabled"
           else
              iPermission3= " "
		   end if

           if CheckPermission(226,4)=false then
              iPermission4= "disabled"
           else
              iPermission4= " "
		   end if

%>


 
  
<%


 
GetLawItem=Trim(request("lawItem") & "")
if request("select999")="999" then GetLawItem=GetLawItem &",999"
CarTypeList=Trim(request("CarTypeListValue") & "")


GetBillType=trim(request("BillType"))
if GetBillType="3" then 
  GetBillType="1,2"
end if


If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	
iTag= Request.querystring("tag")
iProjectID= Request.querystring("ProjectID")
Set gadoRS = Server.CreateObject("ADODB.Recordset")

	if iTag="del" then


				iDelMember="null"
					
			gsql="update project set "
			gsql=gsql & " RecordStateID='-1' "
			gsql=gsql & " ,RecordDate=sysdate "
			gsql=gsql & " ,RecordMemberID='"&Session("User_ID")&"' "
			gsql=gsql & " ,DelMemberID ="&iDelMember
			gsql=gsql & " where ProjectID ='"&trim(iProjectID)&"'"
		
			
		
			conn.execute (gsql)
				'response.write gsql
				'response.end
				response.redirect "project.asp"		
	end if
	if Request.querystring("Save")="Y" then
	
		if iTag="new" then
			if trim(Request("RecordStateID"))="0" then
				iDelMember="null"
			else
				iDelMember="'"&Session("User_ID")&"'"
			end if
			
			gsql="select projectid from project where projectid ='"&replace(trim(Request("ProjectID")),"'","''")&"' "
			set gadoRS=conn.execute(gsql)
			if gadoRS.eof then

				gsql="insert into Project ( ProjectID , Name , StartDate , EndDate , RecordStateID, RecordDate , RecordMemberID , DelMemberID ,BILLTYPEIDLIST ,LAWIDLIST,CarTypeIDList)"
				gsql=gsql & " values ('"&replace(trim(Request("ProjectID")),"'","''")&"' ,'"&replace(trim(Request("Name")),"'","''")&"',"
				gsql=gsql & funGetDate(gOutDT(trim(Request("StartDate"))),0) &","&funGetDate(gOutDT(trim(Request("EndDate"))),0)&",'"&trim(Request("RecordStateID"))&"',sysdate,'"&Session("User_ID")&"',"&iDelMember&",'"&GetBillType&"','"&GetLawItem &"','"&CarTypeList&"') "
		'---------------------------------------------------------------------------------------		
		'response.write gsql
		'response.end
				conn.execute gsql
			
				response.redirect "project.asp"
			else
				   response.write "<script>alert('專案代碼已存在');history.back() ;</script>"
				   response.write "<script>history.back() ;</script>"				   
				 gadoRS.close  
			
			end if
	
		else
		
			if trim(Request("RecordStateID"))="0" then
				iDelMember="null"
			else
				iDelMember="'"&Session("User_ID")&"'"
			end if
					
			gsql="update project set name ='"&replace(trim(Request("name")),"'","''")&"' "
			gsql=gsql & " , StartDate ="&funGetDate(gOutDT(trim(Request("StartDate"))),0)
			gsql=gsql & " , EndDate ="&funGetDate(gOutDT(trim(Request("EndDate"))),0)
			gsql=gsql & " ,RecordStateID='"&trim(Request("RecordStateID"))&"' "
			gsql=gsql & " ,RecordDate=sysdate "
			gsql=gsql & " ,RecordMemberID='"&Session("User_ID")&"' "
			gsql=gsql & " ,DelMemberID ="&iDelMember
			gsql=gsql & " ,BILLTYPEIDLIST ='" & GetBillType & "'"
			gsql=gsql & " ,LAWIDLIST='" & GetLawItem & "'"
			gsql=gsql & " ,CarTypeIDList='" & CarTypeList & "'"
			gsql=gsql & " where ProjectID ='"&trim(Request("mdyProjectID"))&"'"
		'response.write gsql
			conn.execute (gsql)
				'
				'response.end
				response.redirect "project.asp"		
		end if
	
	end if
		
	if itag="mdy" and iProjectID <>"" and Request.querystring("Save")="" then

		gsql="select ProjectID , Name ,StartDate,EndDate,RecordStateID ,(decode (RecordStateID,'0','開始','停止') )as RecordStateDesc,RecordDate,LawIDList,BillTypeIDList,CarTypeIDList  "
		gsql=gsql & " from Project where ProjectID='"&iProjectID&"'"

        iBillTypeIDList=""
		iLawIDList=""
		iCarTypeList=""
		
		set gadoRS=conn.execute(gsql)
		if not gadoRS.eof then
			iName=gadoRS("Name")
			iStartDate=gadoRS("StartDate")
			iEndDate=gadoRS("EndDate")
			iRecordStateID=gadoRS("RecordStateID")
			iRecordStateDesc=gadoRS("RecordStateDesc")			
            iLawIDList=gadoRS("LawIDList")
			iBillTypeIDList=gadoRS("BillTypeIDList")
			iCarTypeList=gadoRS("CarTypeIDList")
		end if
		gadoRS.close
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>專案資料檔維護</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {font-size: 14px}
.style3 {font-size: 15px}
-->
</style></head>


<%
gsql="select ProjectID from Project "
gadoRS.open gsql,conn,3,1


%>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<script language="javascript">

var id = new Array(<%=gadoRS.recordcount%>); 


<%
	if not gadoRS.eof then
		k=0
		while not gadoRS.eof
%>
		id[<%=k%>]="<%=gadoRS("ProjectID")%>"; 
<%
			k=k+1
			gadoRS.movenext
		wend
	end if
	gadoRS.close
%>

function  datacheck(){
Project.lawItem.value=GetLawItem();
Project.CarTypeListValue.value=GetCarTypeList();


	if (document.Project.ProjectID.value==""){
		alert ('請輸入專案代碼...!!');
		return false
	}

		if (document.Project.name.value==""){
		alert ('請輸入專案名稱...!!');
		return false
	}

	if (document.Project.select3.length==0){
		alert ('請輸入法條...!!');
		return false
	}

	if (document.Project.StartDate.value==""){
		alert ('請輸入專案施行期間...!!');
		return false
	}

	if (document.Project.EndDate.value==""){
		alert ('請輸入專案施行期間...!!')
		return false
	}	

var DateS =new String(document.Project.StartDate.value)	;
var DateE =new String(document.Project.EndDate.value) ;
DateS =new Date(DateS)
DateE =new Date(DateE)
	if (DateS > DateE){	
		alert ('起始日期大於結束日期...!!');
		return false
	}	


	if (document.Project.ProjectID.value!=""){
           for( j=0; j<id.length; j++ )
            {		
				if (id[j]==document.Project.ProjectID.value){
				alert ('專案代碼已存在...!!');				
				return false ;
				}
			}
	}


}

function openQryLaw(){
	 window.open("../Report/QueryLaw.asp?qryType=1&reportId=REPORTBASE0010","tmpWindow","width=600,height=355,left=0,top=0,resizable=yes,scrollbars=yes");
}	

function openQryLaw2(){
	 window.open("../Report/QueryDCICar.asp?qryType=1&reportId=REPORTBASE0010","tmpWindow","width=600,height=355,left=0,top=0,resizable=yes,scrollbars=yes");
}

function remove(){
   obj = document.all.select3 ;
   objIndex = obj.selectedIndex;   
   if (objIndex != -1) {
      obj.remove(obj.selectedIndex);
   }
}   

function remove2(){
   obj = document.all.CarTypeList ;
   objIndex = obj.selectedIndex;   
   if (objIndex != -1) {
      obj.remove(obj.selectedIndex);
   }
}   

function GetBillType()
{
var valueStr;

 for(i=0;i<document.all.BillType.length;i++){
    if(document.all.BillType[i].checked){
		valueStr=document.all.BillType[i].value;
	}
  }
  return valueStr
       
}

function GetLawItem()
{
var tmplaw="";
var tmpStr="";
    if (document.all.select3.length!=0) 
    {
	    for(i=0;i<document.all.select3.length;i++){
	    	  if(i==0){
				 tmplaw = document.all.select3.options[i].value ;
	    	  	 tmpStr = document.all.select3.options[i].value ;		  	 
	    	  }else{
				 if(tmplaw != document.all.select3.options[i].value){
					tmplaw = document.all.select3.options[i].value ;
	    	  		tmpStr = tmpStr + "," + document.all.select3.options[i].value ;
				 }
	    	  }
       }
    }       
	
     return tmpStr + "";
       
}

function GetCarTypeList()
{
var tmplaw="";
var tmpStr="";
   if (document.all.CarTypeList.length!=0)
   {
	    for(i=0;i<document.all.CarTypeList.length;i++){
	    	  if(i==0){
				 tmplaw =  document.all.CarTypeList.options[i].value ;
	    	  	 tmpStr =  document.all.CarTypeList.options[i].value ;
		  	 
	    	  }else{
				 if(tmplaw != document.all.CarTypeList.options[i].value){
					tmplaw =  document.all.CarTypeList.options[i].value ;
	    	  		tmpStr = tmpStr + "," + document.all.CarTypeList.options[i].value ;
				 }
	    	  }
       }
   }       
     return tmpStr + "" ;       
}

</script>

<body>
<FORM NAME="Project" ACTION="ProjectDetail.asp?tag=<%=iTag%>&Save=Y" METHOD="POST" onSubmit="return datacheck();">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style3">專案資料檔維護&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><a href="Project.jpg" target="_blank" class="style2">專案新增簡易說明</a>	</td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC">
    <table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF" height="310">
      <tr>
      	<%	if iProjectID<>"" then 
      			idisabled=" disabled"
      	
      		end if
      	%>
        <td width="11%" bgcolor="#FFFFCC" height="21"><div align="right" class="style3">專案代碼</div></td>
        <td width="89%" height="21"><span class="style3">
          <input name="ProjectID" type="text" value="<%=iProjectID%>" size="20" maxlength="25" <%=idisabled%>  class="btn1">
          <input name="mdyProjectID" type="hidden" value="<%=iProjectID%>" size="20"   class="btn1">          
          </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC" height="21"><div align="right"><span class="style3">專案名稱</span></div></td>
        <td height="21"><span class="style3">
          <input name="name" type="text" value="<%=iName%>" size="20" maxlength="25"  class="btn1"> 
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC" height="25"><div align="right"><span class="style3">專案施行期間</span></div></td>
        <td height="25"><span class="style3">
<input type='text' size='10' id='StartDate' class="btn1" name='StartDate' value="<%=gInitDT(iStartDate)%>" readonly onclick="OpenWindow('StartDate')"><input type=button value="..." name='btnDateS'  onclick="OpenWindow('StartDate')">      ~
<input type='text' size='10' id='EndDate'  class="btn1"name='EndDate' value="<%=gInitDT(iEndDate)%>" readonly onclick="OpenWindow('EndDate')"><input type=button value="..." name='btnDateE'  onclick="OpenWindow('EndDate')">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC" height="16">
        <p align="right">適用車別</td>
        <td height="16">
          <input type="button" name="btnAdd" value="加入" onClick="openQryLaw2(this.value);">
          <input type="button" name="btnDel" onClick="remove2();" value="刪除">　慢車行人時，不選車別<p>
          <select name="CarTypeList" multiple size=9>
            <%
            If trim(iCarTypeList)<>"" Then 
            
              tempStr=""
            
                iCarTypeList=split(iCarTypeList,",")
                
              for i=0 to Ubound(iCarTypeList)
                if i=Ubound(iCarTypeList) then 
                  tempStr=tempStr & "'" & iCarTypeList(i) & "'"
                else
                  tempStr=tempStr & "'" & iCarTypeList(i) & "',"
                end if 
              next 
            
              RuleStr="select ID,Content from dcicode where typeid=5 and content is not null and ID in ("&tempstr&")"
   	          set Rule=conn.execute(RuleStr)
     	      while Not Rule.eof
                response.write "<option value="&Rule("ID")&">"&Rule("Content")&"</option>"
	 	      Rule.moveNext
              wend
	          Rule.close
	    	  Set Rule=nothing	  
	    	  
			End If 
			
			%>
          </select></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC" height="109">
        <p align="right"><span class="style3">法條</span></td>
        <td height="109">
          <input type="button" name="btnAdd" value="加入" onClick="openQryLaw(this.value);">
          <input type="button" name="btnDel" onClick="remove();" value="刪除"><p>
			不包含以下法條<input type="checkbox" name="select999" value="999" <% if instr(iLawIDList,"999")>0 then response.write "checked"%>></p>
			<p>
          <select name="select3" multiple size=9>
		    <%
			 
			If Trim(iLawIDList)<>"" Then 
               RuleStr="select distinct ItemID,IllegalRule from law where version=2 and ItemID in ("&iLawIDList&")"
			   set Rule=conn.execute(RuleStr)
	         	while Not Rule.eof
                  response.write "<option value="&Rule("ItemID")&">"&Rule("IllegalRule")&"</option>"
		          Rule.moveNext
	            wend
	        	Rule.close
	    	Set Rule=nothing	  
			End If 
			%>
          </select></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC" height="20">
        <p align="right"><span class="style3">舉發單別</span></td>
        <td height="20">

		<%If iBillTypeIDList="1" Then %>
		  <input type="radio" value="1" checked name="BillType">攔停　
          <input type="radio" name="BillType" value="2">逕舉　
          <input type="radio" name="BillType" value="3">全部
        <%End if%>

    	<%If iBillTypeIDList="2" Then %>
		  <input type="radio" value="1" name="BillType">攔停　
          <input type="radio" checked name="BillType" value="2">逕舉　
          <input type="radio" name="BillType" value="3">全部
        <%End if%>

    	<%If iBillTypeIDList="1,2" Then %>
		  <input type="radio" value="1" name="BillType">攔停　
          <input type="radio" name="BillType" value="2">逕舉　
          <input type="radio" checked name="BillType" value="3">全部
        <%End if%>

    	<%If iBillTypeIDList="" Then %>
		  <input type="radio" checked value="1" name="BillType">攔停　
          <input type="radio" name="BillType" value="2">逕舉　
          <input type="radio" name="BillType" value="3">全部
        <%End if%>

		<p>「拖吊」、「慢車行人」舉發單別選「全部」</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC" height="21"><div align="right"><span class="style3">狀態</span></div></td>
        <td height="21"><span class="style3">
          <select name="RecordStateID">
          	<%if trim(iRecordStateID&" ")="-1" then %>
              <option value='-1'>停止</option>            
              <option value='0'>開始</option>            

            <%else%>  
              <option value='0'>開始</option>
              <option value='-1'>停止</option>
            
            <%end if%>  
            </select>
        </span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
    </a>
        <input type="submit" name="Submit423" value="確 定" >
        <input type="hidden" name="lawItem" value="">
        <input type="hidden" name="CarTypeListValue" value="">
        

        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="button" name="btnBack" value="關 閉" onclick="javascript:document.location.href='Project.asp'">
</p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    </td></tr>
</table>

<%
	'gadoRS.close
	Set gadoRS=nothing
	Conn.close
	Set Conn=nothing

%>
</body>
</html>