<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
    <head>
        <meta   http-equiv="Content-Type"   content="text/html;   charset=big5"> 
        <script language="JavaScript">
	        window.focus();
        </script> 
         
        <title>明細查詢</title>
        <!--#include virtual="Traffic/Common/css.txt"-->
        <!--#include virtual="traffic/Common/db.ini"-->
        <!--#include virtual="Traffic/Common/AllFunction.inc"-->
        <%
            RealSpeed1="0" '實際車速1
            LimitSpeed1="0" '限制車速1 
            RealSpeed2="0" '實際車速2
            LimitSpeed2="0" '限制車速2 
            function QuotedStr(Str)
                QuotedStr="'"+Str+"'"
            end function
           
            '判斷如果是0的話回傳&nbsp;
            function ReplaceSpace(value)
                if value="" then
                    ReplaceSpace =  "&nbsp;"
                else
                    ReplaceSpace = value 
                end if 
            end Function
Function GetReturnResultName(ReturnResultID)
	tmp=""
	tmp2=Trim(ReturnResultID)
	If tmp2="N" Then 
		tmp="找不到"
	ElseIf tmp2="C" Then 
		tmp="已結案"
	ElseIf tmp2="P" Then 
		tmp="已公示"
	ElseIf tmp2="S" Then 
		tmp="成功"
	Else
		tmp="成功"
	End If
	GetReturnResultName=tmp
End Function

Function GetReturnReasonName(ReturnreasonID)
	tmp=""
	tmp2=Trim(ReturnreasonID)
	If tmp2="1" Then 
		tmp="遷移不明"
	ElseIf tmp2="2" Then 
		tmp="查無此人"
	ElseIf tmp2="3" Then 
		tmp="地址欠詳"
	ElseIf tmp2="4" Then 
		tmp="查無此址"
	ElseIf tmp2="5" Then 
		tmp="招領逾期"
	ElseIf tmp2="6" Then 
		tmp="拒收"
	ElseIf tmp2="7" Then 
		tmp="投箱待領逾期"
	ElseIf tmp2="8" Then 
		tmp="其他"
	Else
		tmp=ReturnreasonID
	End If
	GetReturnReasonName=tmp
End function
            
Function GetCarTypeName(CarTypeID)
	tmp=""
	tmp2=Trim(CarTypeID)
	If tmp2="1" Then 
		tmp="自大客車"
	ElseIf tmp2="4" Then 
		tmp="營大客車"
	ElseIf tmp2="7" Then 
		tmp="營小客車"
	ElseIf tmp2="A" Then 
		tmp="營交通車"
	ElseIf tmp2="G" Then 
		tmp="大型重機"
	ElseIf tmp2="P" Then 
		tmp="併裝車"
	ElseIf tmp2="W" Then 
		tmp="自小客"
	ElseIf tmp2="2" Then 
		tmp="自大貨車"
	ElseIf tmp2="5" Then 
		tmp="營大貨車"
	ElseIf tmp2="8" Then 
		tmp="租賃小客"
	ElseIf tmp2="B" Then 
		tmp="貨櫃曳引"
	ElseIf tmp2="E" Then 
		tmp="外賓小客"
	ElseIf tmp2="H" Then 
		tmp="重機"
	ElseIf tmp2="Q" Then 
		tmp="500cc重機"
	ElseIf tmp2="X" Then 
		tmp="動力機械"
	ElseIf tmp2="3" Then 
		tmp="自小客貨"
	ElseIf tmp2="6" Then 
		tmp="營小貨車"
	ElseIf tmp2="9" Then 
		tmp="遊覽客車"
	ElseIf tmp2="C" Then 
		tmp="自用拖車"
	ElseIf tmp2="F" Then 
		tmp="外賓大客"
	ElseIf tmp2="L" Then 
		tmp="輕機"
	ElseIf tmp2="V" Then 
		tmp="自小貨"
	ElseIf tmp2="Y" Then 
		tmp="租賃小貨"
	End If
	If tmp="" Then tmp=CarTypeID
	GetCarTypeName=tmp

End function
                        
            '查詢某一個欄位
            function SelectFld(TableName,Fld,Cond)  
                QuerySql="Select " & Fld & " from " & TableName & " where " & Cond
                set QueryRS=conn.execute(QuerySql)
                if  not QueryRS.Eof then
                    SelectFld = QueryRS(Fld)
                end if 
                QueryRS.close
            end  function   
                     
            function SetEngDate(tDate)
	            if len(tDate)=7 then
		            SetEngDate=left(tDate,3)&"年"& mid(tDate,4,2)&"月"& Right(tDate,2)&"日"
	            else
		            SetEngDate="&nbsp;"
	            end if
            end function
           
            function SetchinaDate(tDate)
	            if len(tDate)=6 then
		            SetchinaDate=mid(tDate,1,2) &"年"& mid(tDate,3,2)&"月"& mid(tDate,5,2)&"日"
	            else
		            SetchinaDate="&nbsp;"
	            end if
            end function  

            function GetTime(ttime)
              W=""
              H=""
              H=left(ttime,2)
              N=right(ttime,2)
              if cdbl(H)=12 then
                W="中午"
              elseif cdbl(H)<6  then
                W="凌晨"
              elseif cdbl(H)>5 and cdbl(H)<12 then
                W="早上"
              elseif cdbl(H)>12 and cdbl(H)<18 then
                W="下午"
              elseif cdbl(H)>17 then
                W="晚上"
              end if

              SH=0

              if H>12 then SH=cdbl(H)-12 else SH=H

	            if len(ttime)=4 then
		            GetTime=W&" "&right("00"&SH,2)&"點"&N&"分"
	            else
		            GetTime="&nbsp;"
	            end if
            end function  
                       
           	strSQL="Select Billno,carno,CarTypeID,Driver,BirthDay,IDdata,DriverAddress,owner,owneraddress,illegaldate,illTime,illegalPlaceID,illegalplace,rule1,rule1real,rule1limit,rule2,rule2real,rule2limit,billunitid,billmemid1,recorddate,deallinedate,station,MoveDate from oldbillbase where  billno="  & QuotedStr(request("BillNo"))  
            set rs1=conn.execute(strSQL) 
        %> 
    </head>
    <body>
        <table width='100%' border='1' cellpadding="2" id="table1">
		<tr bgcolor="#FFCC33">
			<td><strong>舊告發單詳細資料</strong></td>
		</tr>
		</table> 
        <table width="100%"   border='1' cellpadding="2" style="border-top-style: groove; border-right-style: groove; border-left-style: groove; border-bottom-style: groove" >
            <tr>
                <td style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>單&nbsp;號</strong>
                </td> 
                <td style="width: 15%" colspan="8" align="center">
                    <%=ReplaceSpace(rs1("Billno"))%>
                </td>
               <td  style="width: 15%" colspan="5" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;號</strong>
                </td>  
                <td style="width: 17%" colspan="10" align="center">
                    <%=ReplaceSpace(rs1("carno"))%>
                </td>
                <td  style="width: 8%" colspan="8" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;別</strong>
                </td>  
                <td style="width: 10%" colspan="7" align="center">
                    <%=ReplaceSpace(GetCarTypeName(Trim(rs1("CarTypeID"))))%>
                </td> 
            </tr>
           <tr>
               <td   style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                      <strong>違&nbsp;規&nbsp;人</strong>
               </td> 
               <td  style="width: 15%" colspan="8" align="center">
                   <%=ReplaceSpace(trim(rs1("Driver"))) %>
               </td>
                 

                <td  style="width: 10%" colspan="5" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>出&nbsp;生&nbsp;日&nbsp;期</strong> 
               </td>
               <td style="width: 17%" colspan="10" align="center">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("BirthDay"))))%>      
               </td> 
               <td   style="width: 10%" colspan="8" bgcolor="#FFFF99" align="center"  height="35">
                      <strong>證&nbsp;號</strong>
               </td> 
               <td  style="width: 15%" colspan="7" align="center">
                    <%=ReplaceSpace(trim(rs1("IDdata"))) %>
               </td>
           </tr> 
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;人&nbsp;地&nbsp;址</strong> 
                </td>
               <td colspan="38" height="35">
                    <%=ReplaceSpace(trim(rs1("DriverAddress")))%>      
               </td> 
           </tr> 
           <tr>
               <td   style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                      <strong>車&nbsp;主</strong>
               </td> 
               <td  style="width: 15%" colspan="8" align="center">
                   <%=ReplaceSpace(trim(rs1("owner"))) %>
               </td>
                 
                <td   style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                       <strong>車&nbsp;主&nbsp;地&nbsp;址</strong>
                </td> 
                <td  style="width: 15%" colspan="27" align="left">
                    <%=ReplaceSpace(trim(rs1("owneraddress"))) %>
                </td>
           </tr> 
           <tr>
               <td  style="width: 10%; height: 41px;" colspan="	6" bgcolor="#FFFF99" align="center">
                    <strong>違&nbsp;規&nbsp;時&nbsp;間</strong> 
                </td>
               <td colspan="8" style="height: 41px">
                    <%=Trim(ReplaceSpace(SetchinaDate(rs1("illegaldate"))) & " " & left(rs1("illTime"),2)&":"&Right(rs1("illTime"),2))%>      
               </td>
                <td  style="width: 15%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>違&nbsp;規&nbsp;地&nbsp;點&nbsp;</strong> 
                </td>
               <td colspan="27" style="height: 41px">
                    <%=ReplaceSpace(trim(rs1("illegalPlaceID")) &"  " & trim(rs1("illegalplace")))%>      
               </td>
           </tr> 
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;條&nbsp;款&nbsp;一</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("rule1")))%>      
               </td> 
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>實際車速車重</strong>
               </td> 
               <td style="width: 10%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("rule1real")))%>     
               </td>
               <td  style="width: 10%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>限制車速車重</strong>
               </td> 
               <td style="width: 10%" colspan="6">
                    <%=ReplaceSpace(trim(rs1("rule1limit")))%>   
               </td> 
           </tr>
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;條&nbsp;款&nbsp;二</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("rule2")))&"&nbsp;"%>      
               </td> 
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>實際車速車重</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("rule2real")))%>   
               </td>
               <td  style="width: 10%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>限制車速車重</strong>
               </td> 
               <td style="width: 22%" colspan="6">
                    <%=ReplaceSpace(trim(rs1("rule2limit")))%>     
               </td> 
           </tr>
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>舉&nbsp;發&nbsp;單&nbsp;位</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("billunitid")))%>
               </td>  
               <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>舉&nbsp;發&nbsp;員&nbsp;警</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("billmemid1")))%>
               </td>         
               <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>建&nbsp;檔&nbsp;日</strong>
               </td> 
               <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("recorddate"))))%>
               </td>          
           </tr>
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>到&nbsp;案&nbsp;日</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("deallinedate"))))%>
                </td>  
                <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>到&nbsp;案&nbsp;單&nbsp;位</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("station")))%>
                </td>         
                <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>移&nbsp;送&nbsp;日&nbsp;期</strong>
                </td> 
                <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("MoveDate"))))%>
                </td>          
           </tr>    
		   <tr>
		   <%
           	strSQL="Select billunitID,billunitname,recorddate,illegaldate,station,movedate,returnReason,returnresult,publicStoreRecordDate,PublicStoreRecordmemID,arriveDate,StartDate,ArriveWord,FirstReturnDate,CarNo,CarTypeID,Owner,OwnerAddr from oldbillReturn where  billno="  & QuotedStr(request("BillNo"))  
            set rsRe=conn.execute(strSQL) 
			while Not rsRe.eof
		   %>
		   <td colspan="50" bgcolor="#FFC000">
			&nbsp;&nbsp;單&nbsp;&nbsp;退&nbsp;&nbsp;資&nbsp;&nbsp;料
		   </td>
		   <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>建&nbsp;檔&nbsp;單&nbsp;位&nbsp;代&nbsp;號</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace((trim(rsRe("billunitID"))))%>
                </td>  
                <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>建&nbsp;檔&nbsp;單&nbsp;位&nbsp;名&nbsp;稱</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rsRe("billunitname")))%>
                </td>         
                <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>建&nbsp;檔&nbsp;日</strong>
                </td> 
                <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace(SetchinaDate(trim(rsRe("recorddate"))))%>
                </td>          
           </tr>    
		   <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;日</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(SetchinaDate(trim(rsRe("illegaldate"))))%>
                </td>  
                <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>到&nbsp;案&nbsp;地</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rsRe("station")))%>
                </td>         
                <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>單退移送日</strong>
                </td> 
                <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace(SetchinaDate(trim(rsRe("movedate"))))%>
                </td>          
           </tr>    
		   <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>單&nbsp;退&nbsp;原&nbsp;因</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(GetReturnReasonName(trim(rsRe("returnReason"))))%>
                </td>  
                <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>單&nbsp;退&nbsp;結&nbsp;果</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(GetReturnResultName(trim(rsRe("returnresult"))))%>
                </td>         
                <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>(公示/寄存)<br>建檔日</strong>
                </td> 
                <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace(SetchinaDate(trim(rsRe("publicStoreRecordDate"))))%>
                </td>          
           </tr>    
		   <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>(公示/寄存)<br>建檔人</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace((trim(rsRe("PublicStoreRecordmemID"))))%>
                </td>  
                <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>刊載日/送達日期</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(SetchinaDate(trim(rsRe("arriveDate"))))%>
                </td>         
                <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>裁罰日/生效日期</strong>
                </td> 
                <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace(SetchinaDate(trim(rsRe("StartDate"))))%>
                </td>          
           </tr>    
		   <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>刊載碼/送達文號</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace((trim(rsRe("ArriveWord"))))%>
                </td>  
                <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>第一次郵退日</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(SetchinaDate(trim(rsRe("FirstReturnDate"))))%>
                </td>         
                <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;號</strong>
                </td> 
                <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace((trim(rsRe("CarNo"))))%>
                </td>          
           </tr>    
		   <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;種</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(GetCarTypeName(trim(rsRe("CarTypeID"))))%>
                </td>  
                <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;主</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rsRe("Owner")))%>
                </td>         
                <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;主&nbsp;地&nbsp;址</strong>
                </td> 
                <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace((trim(rsRe("OwnerAddr"))))%>
                </td>          
           </tr>    
		   <tr>
<%
			rsRe.movenext
		wend
		
%>
</table>
<table>
		   <%
		   '\\10.102.114.234\ret_recnew\96\04\01\B032370911S.bmp
		   '\\10.102.114.235\ret_recnew\96\04\01\B032370911S.bmp
		   'http://10.102.114.234/ret_recnew/96/04/01/B032370911S.bmp
		   '
			If CDbl(rs1("illegaldate"))>960101 Then 
			' tmp="F:\oldImage"
			' tmp2="http://10.102.114.158/OldImage"
			Else
			' tmp="\\10.102.114.235"
			' tmp2="http://10.102.114.235"
			End if
			tmp="G:"
			tmp2="http://10.102.114.158/OldImage"
		   set fs=CreateObject("Scripting.FileSystemObject")  

		   For i=1 To 3
		   If fs.FileExists(tmp&"\ret_recnew\"&mid(rs1("illegaldate"),1,2)&"\"&mid(rs1("illegaldate"),3,2)&"\"&mid(rs1("illegaldate"),5,2)&"\"&Trim(rs1("billno"))&i&"S.bmp") then %>
		   <td colspan="50" bgcolor="#FFFF99">
			送達證書
		   </td>
		   <TR>
		   <td colspan="50">
		   <img src="<%=tmp2%>/ret_recnew/<%=mid(rs1("illegaldate"),1,2)%>/<%=mid(rs1("illegaldate"),3,2)%>/<%=mid(rs1("illegaldate"),5,2)%>/<%=Trim(rs1("billno"))&i%>S.bmp">
		   </td>
		   <TR>
			<%End If
			Next

			For i=1 To 3
		    If fs.FileExists(tmp&"\vil_errnew\"&mid(rs1("illegaldate"),1,2)&"\"&mid(rs1("illegaldate"),3,2)&"\"&mid(rs1("illegaldate"),5,2)&"\"&Trim(rs1("billno"))&i&"E.JPG") then %>
		   <td colspan="50" bgcolor="#FFFF99">
			註銷案件
		   </td>
		   <TR>
		   <td colspan="50">
		   <img src="<%=tmp2%>/vil_errnew/<%=mid(rs1("illegaldate"),1,2)%>/<%=mid(rs1("illegaldate"),3,2)%>/<%=mid(rs1("illegaldate"),5,2)%>/<%=Trim(rs1("billno"))&i%>E.JPG">
		   </td>
		   <TR>
			<%End If
			Next

			For i=1 To 3
			If fs.FileExists(tmp&"\vil_recnew\"&mid(rs1("illegaldate"),1,2)&"\"&mid(rs1("illegaldate"),3,2)&"\"&mid(rs1("illegaldate"),5,2)&"\"&Trim(rs1("billno"))&i&"V.JPG") then %>
		   <td colspan="50" bgcolor="#FFFF99">
			入案案件
		   </td>
		   <TR>
		   <td colspan="50">
		   <img src="<%=tmp2%>/vil_recnew/<%=mid(rs1("illegaldate"),1,2)%>/<%=mid(rs1("illegaldate"),3,2)%>/<%=mid(rs1("illegaldate"),5,2)%>/<%=Trim(rs1("billno"))&i%>V.JPG">
		   </td>
		   <TR>
			<%End If
			Next%>

           </tr> 
        </table> 
        <center>
            <input type="button" name="Submit4233" onClick="javascript:window.print();" value="列 印"> 
            <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉">
        </center>
    </body>
</html>
<%
Set fs=nothing
conn.close
set conn=nothing
%>