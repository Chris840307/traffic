<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
    <head>
        <meta   http-equiv="Content-Type"   content="text/html;   charset=big5"> 
        <script language="JavaScript">
	        window.focus();
        </script> 
         
        <title>歷史查詢</title>
        <!--#include virtual="Traffic/Common/css.txt"-->
        <!--#include virtual="traffic/Common/DB.ini"-->
        <!--#include virtual="Traffic/Common/AllFunction.inc"-->
        <%
            RealSpeed1="0" '實際車速1
            LimitSpeed1="0" '限制車速1 
            RealSpeed2="0" '實際車速2
            LimitSpeed2="0" '限制車速2 
            function QuotedStr(Str)
                QuotedStr="'"+Str+"'"
            end function

            '查詢某一個欄位
            function SelectFld(TableName,Fld,Cond,conn)  
                QuerySql="Select " & Fld & " from " & TableName & " where " & Cond
                set QueryRS=conn.execute(QuerySql)
                if  not QueryRS.Eof then
                    SelectFld = QueryRS(Fld)
                 end if 
                QueryRS.close
            end  function 
            
            Function DateSet(value)
				if trim(value) <> "" then
					DateSet = gInitDT(trim(value))&" "&right("00"&hour(value),2)&":"&right("00"&minute(value),2)    
				end if
            end Function
                               
            sql="Select * from BillBase a,BillBaseDCIReturn b where b.ExchangeTypeID='W' and b.BillNo= " & QuotedStr(request("BillNo")) & " and a.BillNo=b.Billno and b.CarNo=" & QuotedStr(request("CarNo"))
            set rs1=conn.execute(sql)
        %> 
    </head>
    <body>
        <table width="100%"   border='1' cellpadding="2" style="border-top-style: groove; border-right-style: groove; border-left-style: groove; border-bottom-style: groove" >
            <tr>	
			    <td colspan="44" bgcolor="#00FFFF" height="20">
			        <b>&nbsp;&nbsp;結案詳細資料</b>
			    </td> 
			</tr> 
            <tr>
                <td style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>單&nbsp;號</strong>
                </td> 
                <td style="width: 17%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & trim(rs1("BillNo"))%>
                </td>
                 
                <td style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>車&nbsp;號</strong>
                </td> 
                <td style="width: 15%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & trim(rs1("CarNo"))%>
                </td>
                <td style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>車&nbsp;別</strong>
                </td>
                <td style="width: 10%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & SelectFld("DciCode","Content"," TypeID=5 and ID=" & QuotedStr(trim(rs1("DCIReturnCarType"))),conn)%>
                </td>
            </tr>
           <tr>
                <td  style="width: 7%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>車&nbsp;輛&nbsp;顏&nbsp;色</strong>
                </td>  
                <td style="width: 17%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & SelectFld("DCICode","Content"," TypeID=4 and ID=" & QuotedStr(trim(rs1("DCIReturnCarColor"))),conn)%>
                </td>
                <td  style="width: 8%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>車&nbsp;輛&nbsp;廠&nbsp;牌</strong>
                </td>  
                <td style="width: 10%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & trim(rs1("A_Name"))%>
                </td>
                <td   style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>應&nbsp;到&nbsp;案&nbsp;處&nbsp;所</strong>
                </td> 
                <td  style="width: 15%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & SelectFld("Station","DCIStationName"," DCIStationID=" & QuotedStr(trim(rs1("DCIReturnStation"))),conn)%>
                </td>
           </tr> 
           
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>駕&nbsp;駛&nbsp;人</strong> 
                </td>
                <td style="width: 17%" colspan="9" align="center">
                    <%="&nbsp" & trim(rs1("DRIVER"))%>      
                </td>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>駕&nbsp;駛&nbsp;人&nbsp;生&nbsp;日</strong> 
                </td>
                <td style="width: 10%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & trim(DateSet(rs1("DRIVERBIRTH")))%>
                    &nbsp;&nbsp;
                </td>
                
                <td   style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>駕&nbsp;駛&nbsp;人&nbsp;證&nbsp;號</strong>
                </td> 
                <td  style="width: 15%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & trim(rs1("DRIVERID")) %>
                </td>
           </tr>
           
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>駕&nbsp;駛&nbsp;人&nbsp;地&nbsp;址</strong> 
                </td>
               <td colspan="15" style="height: 41px">
                    <%="&nbsp" & trim(rs1("DRIVERADDRESS"))%>      
               </td>
                <td colspan="24" align="left" style="height: 41px">&nbsp;</td>
           </tr> 
           
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>車&nbsp;主</strong> 
                </td>
                <td style="width: 17%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & trim(rs1("OWNER"))%>      
                </td>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>車&nbsp;主&nbsp;證&nbsp;號</strong> 
                </td>
                <td style="width: 10%; height: 41px;" colspan="9" align="center">
                    <%="&nbsp" & DateSet(rs1("OWNERID"))%>
                    &nbsp;&nbsp;
                </td>
                <td   style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>&nbsp;&nbsp;</strong>
                </td> 
                <td  style="width: 15%; height: 41px;" colspan="9" align="center">
                    &nbsp;&nbsp;
                </td>
           </tr>
           
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>車&nbsp;主&nbsp;地&nbsp;址</strong> 
                </td>
               <td colspan="15" style="height: 41px">
                    <%="&nbsp" & trim(rs1("OWNERADDRESS"))%>      
               </td>
                <td colspan="24" align="left" style="height: 41px">&nbsp;</td>
           </tr> 
           
           <tr>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>違&nbsp;規&nbsp;條&nbsp;款&nbsp;一</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%="&nbsp" & trim(rs1("RULE1"))%>      
               </td> 
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>實際車速車重</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="2">
                    <%="&nbsp" & trim(rs1("ILLEGALSPEED"))%>            
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>限制車速車重</strong>
               </td> 
               <td style="width: 22%; height: 41px;" colspan="2">
                    <%="&nbsp" & trim(rs1("RULESPEED"))%>            
               </td> 
               <td colspan="2" style="height: 41px"><%="&nbsp" %></td>  
		
                <td style="width: 15%; height: 41px;" colspan="14" >
                    <%="&nbsp" & trim(rs1("FORFEIT1")) & " " & "元"%>      
               </td> 
           </tr>
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;事&nbsp;實&nbsp;一</strong>
               </td>  
                <td  colspan="40" >
                    <% ="&nbsp" & SelectFld("Law","IllegalRule"," ITEMID=" & QuotedStr(trim(rs1("RULE1"))),conn)%>
                </td> 
           </tr>
           
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;條&nbsp;款&nbsp;二</strong>
               </td> 
               <td style="width: 15%" colspan="9">
                    <%="&nbsp" & trim(rs1("RULE2"))%>       
               </td> 
              <td colspan="2"><%="&nbsp" %></td>  
                <td style="width: 15%" colspan="28">
                    <%="&nbsp" & trim(rs1("FORFEIT2")) & " " & "元"%>      
               </td> 
           </tr>
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;事&nbsp;實&nbsp;二</strong>
               </td>  
                <td  colspan="42" >
                    <% ="&nbsp" & SelectFld("Law","IllegalRule"," ITEMID=" & QuotedStr(trim(rs1("RULE2"))),conn)%>
                </td> 
                
           </tr>   
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>舉&nbsp;發&nbsp;單&nbsp;位</strong>
               </td> 
               <td style="width: 15%" colspan="9">
                    <%="&nbsp" & SelectFld("UnitInfo","UnitName"," UnitID=" & QuotedStr(trim(rs1("BILLUNITID"))),conn)%>      
               </td>  
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>舉&nbsp;發&nbsp;員&nbsp;警</strong>
               </td> 
               <td style="width: 15%" colspan="25">
                    <%="&nbsp" & trim(rs1("BILLMEMID1")) & "  " & trim(rs1("BILLMEM1"))%>
                    &nbsp;       
                    <%="&nbsp" & trim(rs1("BILLMEMID2")) & "  " & trim(rs1("BILLMEM2"))%>
                    &nbsp; 
                    <%="&nbsp" & trim(rs1("BILLMEMID3")) & "  " & trim(rs1("BILLMEM3"))%>
               </td>         
           </tr>
                      
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>到&nbsp;案&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%="&nbsp" & DateSet(trim(rs1("DEALLINEDATE")))%>      
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>建&nbsp;檔&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="24">
                    <%="&nbsp" & DateSet(trim(rs1("BILLFILLDATE")))%>   
               </td>
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