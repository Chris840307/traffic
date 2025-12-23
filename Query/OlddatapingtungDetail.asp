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
        <!--#include virtual="traffic/Common/Olddb_Pingtungdata.ini"-->
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
            end function
                        
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
                       
           	strSQL="Select * from (" &_
	          "Select * from 93 where 1=1 and 單號="  & QuotedStr(request("BillNo"))  &_
	          " union all " &_
	          "Select * from 94 where 1=1 and 單號="  & QuotedStr(request("BillNo"))  &_
	          " union all " &_
	          "Select * from 95 where 1=1 and 單號="  & QuotedStr(request("BillNo"))  &_
	          " union all " &_
	          "Select * from 96 where 1=1 and 單號="  & QuotedStr(request("BillNo"))  &_
	          " )" 
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
                    <%=ReplaceSpace(rs1("單號"))%>
                </td>
               <td  style="width: 15%" colspan="5" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;號</strong>
                </td>  
                <td style="width: 17%" colspan="10" align="center">
                    <%=ReplaceSpace(rs1("車號"))%>
                </td>
                <td  style="width: 8%" colspan="8" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;別</strong>
                </td>  
                <td style="width: 10%" colspan="7" align="center">
                    <%=ReplaceSpace(rs1("車種"))%>
                </td> 
            </tr>
           <tr>
               <td   style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                      <strong>違&nbsp;規&nbsp;人</strong>
               </td> 
               <td  style="width: 15%" colspan="8" align="center">
                   <%=ReplaceSpace(trim(rs1("違規人"))) %>
               </td>
                 

                <td  style="width: 10%" colspan="5" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>出&nbsp;生&nbsp;日&nbsp;期</strong> 
               </td>
               <td style="width: 17%" colspan="10" align="center">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("出生年月日"))))%>      
               </td> 
               <td   style="width: 10%" colspan="8" bgcolor="#FFFF99" align="center"  height="35">
                      <strong>證&nbsp;號</strong>
               </td> 
               <td  style="width: 15%" colspan="7" align="center">
                    <%=ReplaceSpace(trim(rs1("證號"))) %>
               </td>
           </tr> 
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;人&nbsp;地&nbsp;址</strong> 
                </td>
               <td colspan="38" height="35">
                    <%=ReplaceSpace(trim(rs1("違規人地址")))%>      
               </td> 
           </tr> 
           <tr>
               <td   style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                      <strong>車&nbsp;主</strong>
               </td> 
               <td  style="width: 15%" colspan="8" align="center">
                   <%=ReplaceSpace(trim(rs1("車主"))) %>
               </td>
                 
                <td   style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                       <strong>車&nbsp;主&nbsp;地&nbsp;址</strong>
                </td> 
                <td  style="width: 15%" colspan="27" align="left">
                    <%=ReplaceSpace(trim(rs1("車主地址"))) %>
                </td>
           </tr> 
           <tr>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>違&nbsp;規&nbsp;時&nbsp;間</strong> 
                </td>
               <td colspan="10" style="height: 41px">
                    <%=ReplaceSpace(SetchinaDate(rs1("違規日")) & " " & GetTime(rs1("違規時間")))%>      
               </td>
                <td  style="width: 15%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>違&nbsp;規&nbsp;地&nbsp;點&nbsp;</strong> 
                </td>
               <td colspan="27" style="height: 41px">
                    <%=ReplaceSpace(trim(rs1("違規地點代號")) &"  " & trim(rs1("違規地點")))%>      
               </td>
           </tr> 
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;條&nbsp;款&nbsp;一</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("違規法條一")))%>      
               </td> 
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>實際車速車重</strong>
               </td> 
               <td style="width: 10%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("違規法條一實際數")))%>     
               </td>
               <td  style="width: 10%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>限制車速車重</strong>
               </td> 
               <td style="width: 10%" colspan="6">
                    <%=ReplaceSpace(trim(rs1("違規法條一限制數")))%>   
               </td> 
           </tr>
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;條&nbsp;款&nbsp;二</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("違規法條二")))%>      
               </td> 
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>實際車速車重</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("違規法條二實際數")))%>   
               </td>
               <td  style="width: 10%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>限制車速車重</strong>
               </td> 
               <td style="width: 22%" colspan="6">
                    <%=ReplaceSpace(trim(rs1("違規法條二限制數")))%>     
               </td> 
           </tr>
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>舉&nbsp;發&nbsp;單&nbsp;位</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("舉發單位")))%>
               </td>  
               <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>舉&nbsp;發&nbsp;員&nbsp;警</strong>
               </td> 
               <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("舉發員警")))%>
               </td>         
               <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>建&nbsp;檔&nbsp;日</strong>
               </td> 
               <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("建檔日"))))%>
               </td>          
           </tr>
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>到&nbsp;案&nbsp;日</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("到案日"))))%>
                </td>  
                <td  style="width: 15%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>到&nbsp;案&nbsp;單&nbsp;位</strong>
                </td> 
                <td style="width: 15%" colspan="8">
                    <%=ReplaceSpace(trim(rs1("到案地")))%>
                </td>         
                <td  style="width: 15%" colspan="10" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>移&nbsp;送&nbsp;日&nbsp;期</strong>
                </td> 
                <td style="width: 15%" colspan="6">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("移送日"))))%>
                </td>          
           </tr>      
           </tr> 
        </table> 
        <center>
            <input type="button" name="Submit4233" onClick="javascript:window.print();" value="列 印"> 
            <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉">
        </center>
    </body>
</html>
<%
conn.close
set conn=nothing
%>