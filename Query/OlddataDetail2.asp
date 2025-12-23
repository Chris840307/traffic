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
        <!--#include virtual="traffic/Common/OldData2.INI"-->
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
                if trim(value)="" then
                    ReplaceSpace =  "&nbsp;"
                else
                    ReplaceSpace = value 
                end if 
            end function
             
           '判斷法條第八碼，將其組合起來
           function composeLaw(value1,value2)
                if trim(value2) <> "" then
                    composeLaw = value1 & value2
                else
                    composeLaw = value1
                end if  
           end function
             
           '組小數點位數
            function composeDot(value1,value2)  
                if (trim(value1) <> "") and (trim(value2) <> "") then
                    composeDot = value1 & "." & value2
                end if 
                if (trim(value1) <> "") and (trim(value2) = "") then
                    composeDot = value1
                end if
            end  function
            
            '查詢單位名稱
            function QueryUnitName(value)
                UnitSql="Select Acc_NM from accnew where ACC_No=" & QuotedStr(trim(value))
                set UnitRs=conn.execute(UnitSql)
                if  not UnitRs.Eof then
                    QueryUnitName = UnitRs("Acc_NM")
                end if      
                UnitRs.close
            end function    
            
            '查詢操作人員名稱
            function QueryOperat(value)  
                OperatSql="Select OPName from operat where Operat=" & QuotedStr(trim(value))
                set OperatRs=conn.execute(OperatSql)
                if  not OperatRs.Eof then
                    QueryOperat = OperatRs("OPName")
                end if 
                OperatRs.close
            end  function   
            
            '查詢某一個欄位
            function SelectFld(TableName,Fld,Cond)  
                QuerySql="Select " & Fld & " from " & TableName & " where " & Cond
                set QueryRS=conn.execute(QuerySql)

                if  not QueryRS.Eof then
                    SelectFld = QueryRS(Fld)
                end if 
                QueryRS.close
            end  function   
            
            '選擇勾選條件
            function chkPS(value)  
                if trim(value)="0" then
                    chkPS="郵繳"
                else
                    chkPS="到案" 
                end if 
            end  Function
            
						'查詢DCI狀態
            function QryDCIState(value)  
                if trim(value)="00" then
                  QryDCIState="未寫入資料庫"
                ElseIf trim(value)="Y" then
									QryDCIState="寫入資料庫"
								ElseIf trim(value)="N" then
									QryDCIState="未寫入資料庫"
								ElseIf trim(value)="S" then
									QryDCIState="違規人已先繳結案"
								ElseIf trim(value)="L" then
									QryDCIState="已入案過"
								ElseIf trim(value)="n" then
									QryDCIState="不可寫入,監理單位已入案"
                end if 
            end  Function

            '查詢法條內容
            function QueryLawContent(value,RealSpeed,LimitSpeed) 
                LawSql="Select * from rule_n where Rule_C=" & QuotedStr(trim(value))
                set LawRs=conn.execute(LawSql)
                if not LawRs.Eof then
                    if (mid(LawRs("Rule_c"),1,3)="293") and (LawRs("A_DESC")="1") then
                        QueryLawContent = replace(LawRs("Rule_D"),"重量 噸","重量 " & LimitSpeed & " 噸")
                        QueryLawContent = replace(QueryLawContent,"過磅 噸","過磅 " & RealSpeed & " 噸")
                        QueryLawContent = replace(QueryLawContent,"超載 噸","超載 " & RealSpeed-LimitSpeed & " 噸")
                    elseif (mid(LawRs("Rule_c"),1,4)="4010") and (LawRs("A_DESC")="1") then
                        QueryLawContent = replace(LawRs("Rule_D"),"限速 公里","限速 " & LimitSpeed & " 公里")
                        QueryLawContent = replace(QueryLawContent,"時速 公里","時速 " & RealSpeed & " 公里")
                        QueryLawContent = replace(QueryLawContent,"超速 公里","超速 " & RealSpeed-LimitSpeed & " 公里")
                    else
                         QueryLawContent = LawRs("Rule_D")
                    end if 
                end if 

								If Trim(QueryLawContent) <> "" Then
									LawSql="Select * from rule_C where Rule_C=" & QuotedStr(trim(value))
									set LawRs=conn.execute(LawSql)
									if not LawRs.Eof then
											if (mid(LawRs("Rule_c"),1,3)="293") and (LawRs("A_DESC")="1") then
													QueryLawContent = replace(LawRs("Rule_D"),"重量 噸","重量 " & LimitSpeed & " 噸")
													QueryLawContent = replace(QueryLawContent,"過磅 噸","過磅 " & RealSpeed & " 噸")
													QueryLawContent = replace(QueryLawContent,"超載 噸","超載 " & RealSpeed-LimitSpeed & " 噸")
											elseif (mid(LawRs("Rule_c"),1,4)="4010") and (LawRs("A_DESC")="1") then
													QueryLawContent = replace(LawRs("Rule_D"),"限速 公里","限速 " & LimitSpeed & " 公里")
													QueryLawContent = replace(QueryLawContent,"時速 公里","時速 " & RealSpeed & " 公里")
													QueryLawContent = replace(QueryLawContent,"超速 公里","超速 " & RealSpeed-LimitSpeed & " 公里")
											else
													 QueryLawContent = LawRs("Rule_D")
											end if 
									end if 
								End if
            end  function
             
            '檢查保險證
            Function chkissure(value)
                 if trim(value)="0" then
                    chkissure="正常"
                 elseif  trim(value)="1" then
                    chkissure="未帶"
                 elseif  trim(value)="2" then
                    chkissure="肇事且未帶"
                 elseif  trim(value)="3" then
                    chkissure="逾期且未保"
                 elseif  trim(value)="4" then
                    chkissure="肇事且逾期或未帶"
                 end if     
            end function
           
            '檢查簽收情形
            Function chksigner(value)
                 if trim(value)="0" then
                    chksigner="正常"
                 elseif  trim(value)="1" then
                    chksigner="拒簽"
                 elseif  trim(value)="2" then
                    chksigner="拒收"
                 elseif  trim(value)="3" then
                    chksigner="拒簽拒收"
                 end if     
            end function   
            
            function chkBillType(BillTypeID)
                if trim(BillTypeID) <> "" then
                    Select Case  trim(BillTypeID)
                        Case "1" chkBillType="攔停"
                        Case "2" chkBillType="逕舉"
                        Case "3" chkBillType="逕舉手開單" 
                        Case "4" chkBillType="拖吊" 
                        Case "5" chkBillType="慢車行人"   
                        Case "6" chkBillType="肇事"   
                        Case "7" chkBillType="掌-攔停"   
                        Case "8" chkBillType="掌-行人"   
                        Case "9" chkBillType="掌電拖吊"   
                        Case "H" chkBillType="人工移送"   
                        Case "M" chkBillType="郵寄處理"   
                        Case "N" chkBillType="攔停逕行(未開單)"   
                        Case "D" chkBillType="註銷"   
                        Case "R" chkBillType="單退"   
                        Case "V" chkBillType="掌電拖吊(補開單)"   
                    end select       
                end if 
            end function
            
            '選取單退原因
            Function GetReturnCode(Code)
                 if trim(Code)="1" then
                    GetReturnCode="遷移不明"
                 elseif  trim(Code)="2" then
                    GetReturnCode="查無此人"
                 elseif  trim(Code)="3" then
                    GetReturnCode="地址欠詳"
                 elseif  trim(Code)="4" then
                    GetReturnCode="查無地址"
                 elseif  trim(Code)="5" then
                    GetReturnCode="招領逾期"
                elseif  trim(Code)="6" then
                    GetReturnCode="拒收"
                elseif  trim(Code)="7" then
                    GetReturnCode="投箱待領逾期"
                elseif  trim(Code)="8" then
                    GetReturnCode="其他"
                 end if     
            end function     
            
            '選取單退結果
            Function GetReturnResult(Code)
                 if trim(Code)="S" then
                    GetReturnResult="成功"
                 elseif  trim(Code)="N" then
                    GetReturnResult="找不到資料"
                 elseif  trim(Code)="n" then
                    GetReturnResult="己結案"
                 elseif  trim(Code)="k" then
                    GetReturnResult="已送達不可做未達註記"
                 elseif  trim(Code)="Y" then
                    GetReturnResult="撤銷送達"
                elseif  trim(Code)="h" then
                    GetReturnResult="已開裁決書"
                elseif  trim(Code)="B" then
                    GetReturnResult="無此車號/無此證號"
                elseif  trim(Code)="E" then
                    GetReturnResult="日期錯誤"
                 end if     
            end function       
            
            '查詢車種 
            function GetCarType(CarTypeID)
                if trim(CarTypeID) <> "" then
                    Select Case  trim(CarTypeID)
                        Case "1" GetCarType="自大客車"
                        Case "2" GetCarType="自大貨車"
                        Case "3" GetCarType="自小客(貨)" 
                        Case "4" GetCarType="營大客車" 
                        Case "5" GetCarType="營大貨車"   
                        Case "6" GetCarType="營小貨車"   
                        Case "7" GetCarType="營小客車"   
                        Case "8" GetCarType="租賃小客"   
                        Case "9" GetCarType="遊覽客車"   
                        Case "A" GetCarType="營交通車"   
                        Case "B" GetCarType="貨櫃曳引"   
                        Case "C" GetCarType="自用拖車"   
                        Case "D" GetCarType="營業拖車"   
                        Case "E" GetCarType="外賓小客"   
                        Case "F" GetCarType="外賓大客"   
                        Case "H" GetCarType="普通重機"    
                        Case "L" GetCarType="輕機"   
                        Case "p" GetCarType="併裝車"   
                        Case "x" GetCarType="動力機械"      
                        Case "Y" GetCarType="租賃小貨車"
                        Case "W" GetCarType="自小客"
                        Case "V" GetCarType="自小貨"
                        Case "G" GetCarType="大型重機250CC"
                        Case "Q" GetCarType="大型重機550CC"    
                    end select       
                end if 
            end function
            
            '組地址字串
            function composeAddress(Address,lane,alley,No,Dash)
                composeAddress=""
                composeAddress=Address
                if  trim(lane) <> "" then
                    composeAddress=composeAddress & lane & "巷"
                end if
                if  trim(alley) <> "" then
                    composeAddress=composeAddress & alley & "弄"
                end if
                if  trim(No) <> "" then
                    composeAddress=composeAddress & No & "號"
                end if
                if  trim(Dash) <> "" then
                    composeAddress=composeAddress & "之" & Dash
                end if
            end function
            
            function SetEngDate(tDate)
	            if len(tDate)=7 then
		            SetEngDate=left(tDate,3)&"年"& mid(tDate,4,2)&"月"& Right(tDate,2)&"日"
	            else
		            SetEngDate="&nbsp;"
	            end if
            end function
           
            function SetchinaDate(tDate)
	            if len(trim(tDate))=6 then
		            SetchinaDate=mid(tDate,1,2) &"年"& mid(tDate,3,2)&"月"& mid(tDate,5,2)&"日"
	            else
		            SetchinaDate="&nbsp;"
	            end if
            end function  

            function GetTime(ttime)
		if trim(ttime) <> "" and not isnull(ttime)  then
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

		if H>12 then
			SH=cdbl(H)-12
		else
			SH=H		
		end if
		if len(ttime)=4 then
			GetTime=W&" "&right("00"&SH,2)&"點"&N&"分"
		else
			GetTime="&nbsp;"
		end if
		end if
	end function  
	sql="Select * from Vil_Rec where  1=1  and tkt_no=" & QuotedStr(request("BillNo"))

	set rs1=conn.execute(sql) 
'response.write rs1("vil_tm")
'response.end	
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
			    <td colspan="42" bgcolor="#00FFFF" height="20">
			        <b>&nbsp;&nbsp;舉發單基本資料</b>
			    </td> 
			</tr> 
            <tr>
                <td style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>舉&nbsp;發&nbsp;類&nbsp;型</strong>
                </td> 
                <td style="width: 17%; height: 41px;" colspan="9" align="center">
                    <%=ReplaceSpace(chkBillType(trim(rs1("acc_tp"))))%>
                </td>
                 
                <td style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>單&nbsp;號</strong>
                </td> 
                <td style="width: 15%; height: 41px;" colspan="9" align="center">
                    <%=ReplaceSpace(rs1("tkt_no"))%>
                </td>
                <td style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>勾&nbsp;記&nbsp;位&nbsp;置</strong>
                </td>
                <td style="width: 10%; height: 41px;" colspan="9" align="center">
                    <%=ReplaceSpace(rs1("chk_ps") &"   "& chkPS(rs1("chk_ps")))%>
                </td>
            </tr>
           <tr>
                <td  style="width: 7%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;號</strong>
                </td>  
                <td style="width: 17%" colspan="9" align="center">
                    <%=ReplaceSpace(rs1("plt_no"))%>
                </td>
                <td  style="width: 8%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>車&nbsp;別</strong>
                </td>  
                <td style="width: 10%" colspan="9" align="center">
                    <%=ReplaceSpace(trim(rs1("car_tp")) & "   " & GetCarType(trim(rs1("car_tp"))))%>
                </td>
                <td   style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>證&nbsp;號</strong>
                </td> 
                <td  style="width: 15%" colspan="9" align="center">
                    <%=ReplaceSpace(trim(rs1("id_num"))) %>
                </td>
           </tr> 
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>出&nbsp;生&nbsp;日&nbsp;期</strong> 
               </td>
               <td style="width: 17%" colspan="9" align="center">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("birthd"))))%>      
               </td>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>扣&nbsp;件</strong> 
                </td>
                <td colspan="27"  height="37" >
                    <table style="width:100%">
                        <tr>
                            <td style="width: 30%; height: 26px;">
                                <%=ReplaceSpace(mid(rs1("hold_c"),1,1) & "  " & SelectFld("hold_c","Hold_D","Hold_C=" & QuotedStr(mid(rs1("hold_c"),1,1))))%>
                            </td>
                            <td style="width: 30%; height: 26px;">
                                <%=ReplaceSpace(mid(rs1("hold_c"),2,1) & "  " & SelectFld("hold_c","Hold_D","Hold_C=" & QuotedStr(mid(rs1("hold_c"),2,1))))%>
                            </td>
                            <td style="width: 30%; height: 26px;">
                                <%=ReplaceSpace(mid(rs1("hold_c"),3,1) & "  " & SelectFld("hold_c","Hold_D","Hold_C=" & QuotedStr(mid(rs1("hold_c"),3,1))))%>
                            </td>
                        </tr>
                    </table>
                </td>
           </tr> 
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>簽&nbsp;收&nbsp;情&nbsp;形</strong> 
                </td>
                <td style="width: 17%" colspan="9" align="center">
                    <%=ReplaceSpace(trim(rs1("signer") &"  "& chksigner(rs1("signer"))))%>      
                </td>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>保&nbsp;險&nbsp;證</strong> 
                </td>
                <td colspan="9" align="left">
                    <%=ReplaceSpace(rs1("issure") &"  "& chkissure(rs1("issure")))%>
                    &nbsp;&nbsp;
                </td>
								<td colspan="6"  bgcolor="#FFFF99" align="center"  height="35">
										<strong>入&nbsp;案&nbsp;狀&nbsp;態</strong>
								</td>
								<td colspan="9" align="left">
									<%
										DCIState = Mid(trim(rs1("rece02")),5,1) & "," &QryDCIState(Mid(trim(rs1("rece02")),5,1))
                    response.Write ReplaceSpace(DCIState)
									%>
									&nbsp;
								</td>
           </tr>
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>違&nbsp;規&nbsp;時&nbsp;間</strong> 
                </td>
               <td colspan="15" style="height: 41px">

                    <%=ReplaceSpace(SetchinaDate(rs1("vil_dt")) & " " & GetTime(rs1("vil_tm")))%>      
               </td>
                <td colspan="24" align="left" style="height: 41px">&nbsp;</td>
           </tr> 
           <tr>
                <%
                    Address1=""
                    Address2="" 
				    if trim(rs1("vilad1")) <> "" then
				        Address1 = SelectFld("addr_c","Addr_D"," addr_c= " & QuotedStr(trim(rs1("vilad1")))) 
                    end if                     
                   	if trim(rs1("vilad2")) <> "" then
                        Address2 = SelectFld("addr_c","Addr_D"," addr_c= " & QuotedStr(trim(rs1("vilad2")))) 
                    end if  
                 %>        
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>違&nbsp;規&nbsp;地&nbsp;點&nbsp;1</strong> 
                </td>
               <td colspan="39" style="height: 41px">
                    <%=ReplaceSpace(trim(rs1("vilad1")) &"  " & composeAddress(Address1,trim(rs1("vil_a1")),trim(rs1("vil_b1")),trim(rs1("vil_c1")),trim(rs1("vil_d1"))))%>      
               </td>
           </tr> 
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>違&nbsp;規&nbsp;地&nbsp;點&nbsp;2</strong> 
                </td>
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(rs1("vilad2"))) &"     "%>      
               </td> 
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>方&nbsp;向</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="22">
                    <%=ReplaceSpace(trim(rs1("vil_dr")))%>      
               </td>  
           </tr> 
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;條&nbsp;款&nbsp;一</strong>
               </td> 
               <td style="width: 15%" colspan="9">
                    <%=ReplaceSpace(composeLaw(trim(rs1("Rule_1")),Mid(rs1("rece03"),3,1)))%>      
               </td> 
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>實際車速車重</strong>
               </td> 
               <td style="width: 15%" colspan="2">
                    <%
                        RealSpeed1 = composeDot(trim(rs1("R1_SB2")),Mid(trim(rs1("rece01")),3,2))
                        response.Write ReplaceSpace(RealSpeed1)
                    %>      
               </td>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>限制車速車重</strong>
               </td> 
               <td style="width: 22%" colspan="2">
                    <%
                        LimitSpeed1 = composeDot(trim(rs1("R1_SB1")),Mid(trim(rs1("rece01")),1,2))
                        response.Write ReplaceSpace(LimitSpeed1)
                    %>       
               </td> 
               <td colspan="2"></td>  
		
                <td style="width: 15%" colspan="7" >
                    <%=ReplaceSpace(trim(rs1("money1")) & " " & "元")%>      
               </td> 
           </tr>
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;事&nbsp;實&nbsp;一</strong>
               </td>  
                <td  colspan="36" >
                    <% =ReplaceSpace(QueryLawContent(composeLaw(trim(rs1("Rule_1")),Mid(rs1("rece03"),3,1)),RealSpeed1,LimitSpeed1))%>
                </td> 
           </tr>
           
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;條&nbsp;款&nbsp;二</strong>
               </td> 
               <td style="width: 15%" colspan="9">
                    <%=ReplaceSpace(composeLaw(trim(rs1("Rule_2")),Mid(rs1("rece03"),4,1)))%>      
               </td> 
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>實際車速車重</strong>
               </td> 
               <td style="width: 15%" colspan="2">
                    <%
                        RealSpeed2 = composeDot(trim(rs1("R2_SB2")),Mid(trim(rs1("rece02")),3,2))
                        response.Write ReplaceSpace(RealSpeed2)
                    %>      
               </td>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>限制車速車重</strong>
               </td> 
               <td style="width: 22%" colspan="2">
                    <%
                        LimitSpeed2 = composeDot(trim(rs1("R2_SB1")),Mid(trim(rs1("rece02")),1,2))
                        response.Write ReplaceSpace(LimitSpeed2)
                    %>      
               </td> 
              <td colspan="2"></td>  
                <td style="width: 15%" colspan="7">
                    <%=ReplaceSpace(trim(rs1("money2")) & " " & "元")%>      
               </td> 
           </tr>
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>違&nbsp;規&nbsp;事&nbsp;實&nbsp;二</strong>
               </td>  
                <td  colspan="29" >
                    <% =ReplaceSpace(QueryLawContent(composeLaw(trim(rs1("Rule_2")),Mid(rs1("rece03"),4,1)),RealSpeed2,LimitSpeed2))%>
                </td> 
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>特&nbsp;殊&nbsp;專&nbsp;案</strong>
                </td>
                <td>
                    <%=ReplaceSpace(trim(Mid(rs1("rece03"),1,2)))%>      
                </td>
           </tr>   
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>舉&nbsp;發&nbsp;單&nbsp;位</strong>
               </td> 
               <td style="width: 15%" colspan="9">
                    <%=trim(rs1("acc_no")) & "  " & QueryUnitName(trim(rs1("acc_no")))%>      
               </td>  
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>舉&nbsp;發&nbsp;員&nbsp;警</strong>
               </td> 
               <td style="width: 15%" colspan="25">
                    <%=ReplaceSpace(trim(rs1("police")) & "  " & SelectFld("polnew","P_Name"," Police=" & QuotedStr(trim(rs1("police"))) & " and Life_X=0" ))%>

                    &nbsp;       
                    <%=ReplaceSpace(trim(rs1("polic2")) & "  " & SelectFld("polnew","P_Name"," Police=" & QuotedStr(trim(rs1("polic2"))) & " and Life_X=0" ))%> 
                    &nbsp; 
                    <%=ReplaceSpace(trim(rs1("polic3")) & "  " & SelectFld("polnew","P_Name"," Police=" & QuotedStr(trim(rs1("polic3"))) & " and Life_X=0" ))%> 
                    &nbsp; 
                    <%=ReplaceSpace(trim(rs1("polic4")) & "  " & SelectFld("polnew","P_Name"," Police=" & QuotedStr(trim(rs1("polic4"))) & " and Life_X=0" ))%>   
               </td>         
           </tr>
           <tr>
                <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>到&nbsp;案&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("ARV_DT"))))%>      
               </td>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>到&nbsp;案&nbsp;地&nbsp;點</strong>
               </td> 
               <td style="width: 15%" colspan="22">
                    <%=ReplaceSpace(trim(rs1("arvadd")) & "  " & SelectFld("arvadd","ARV_NM","ARVADD=" & QuotedStr(trim(rs1("arvadd")))))%>      
               </td>
           </tr>
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>建&nbsp;檔&nbsp;人</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(rs1("operat")) & "  " & QueryOperat(trim(rs1("operat"))))%>      
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>建&nbsp;檔&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("kin_dt"))))%>      
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#FFFF99" align="center">
                    <strong>紅&nbsp;單&nbsp;序&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(rs1("tktser")))%>      
               </td>                                 
           </tr>
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>付&nbsp;郵&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(SelectFld("ret_rec","mail1d","tkt_no=" & QuotedStr(trim(rs1("tkt_no")))))))%>      
               </td> 
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>大&nbsp;宗&nbsp;掛&nbsp;號</strong>
               </td> 
               <td style="width: 15%" colspan="22">
                    <%=ReplaceSpace(SelectFld("ret_rec","bigser","tkt_no=" & QuotedStr(trim(rs1("tkt_no")))))%>      
               </td>  
           </tr>
           <tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>移&nbsp;送&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(rs1("out_dt"))))%>      
               </td>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
                    <strong>移&nbsp;送&nbsp;批&nbsp;號</strong>
               </td> 
               <td style="width: 15%" colspan="22">

                    <%=ReplaceSpace(trim(rs1("PRN_NO")))%>           
               </td>  
           </tr>
          	<%
			    returnsql="Select * from ret_rec where  1=1  and tkt_no=" & QuotedStr(request("BillNo"))
                set returnrs=conn.execute(returnsql)
                if not (returnrs.EOF and returnrs.BOF) then
			 %> 
            <tr>	
			    <td colspan="42" bgcolor="#00FFFF" height="20">
			        <b>&nbsp;&nbsp;單退資料</b>
			    </td> 
			</tr>

			<tr>
			    <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
			        <b>單&nbsp;退&nbsp;原&nbsp;因</b>     
			    </td>
			    <td style="width: 15%" colspan="9">
			        <%=ReplaceSpace(trim(returnrs("retwhy")) & " " & GetReturnCode(trim(returnrs("retwhy")))) %>
			   </td>
			   <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
			        <b>單&nbsp;退&nbsp;結&nbsp;果</b>     
			    </td>
			    <td style="width: 15%" colspan="9">
			        <%=ReplaceSpace( trim(returnrs("retend")) & " " & GetReturnResult(trim(returnrs("retend")))) %>
			   </td> 
			   <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
			        <b>大&nbsp;宗&nbsp;掛&nbsp;號</b>     
			    </td>
			    <td style="width: 15%" colspan="9">
			        <%=ReplaceSpace(trim(returnrs("bigser"))) %>
			   </td> 
			</tr>
			<tr>
			    <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
			        <b>刊&nbsp;載&nbsp;日</b>     
			    </td>
			   <td style="width: 15%" colspan="9">
			        <%=ReplaceSpace(SetchinaDate(trim(returnrs("mndate"))))%>
			   </td>  
			   <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
			        <b>裁&nbsp;罰&nbsp;日</b>     
			   </td>
			   <td style="width: 15%" colspan="9">
			        <%=ReplaceSpace(SetchinaDate(trim(returnrs("mnacti"))))%>
			   </td>
			   <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
			        <b>刊&nbsp;載&nbsp;碼</b>     
			   </td>
			   <td style="width: 15%" colspan="9">
			        <%=ReplaceSpace(trim(returnrs("mncode")))%>
			   </td>    
			</tr>
			<tr>
               <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
			        <b>寄&nbsp;存&nbsp;期&nbsp;滿&nbsp;日&nbsp;期</b>     
			   </td>
			   <td style="width: 15%" colspan="9">
			        <%=ReplaceSpace(SetchinaDate(trim(returnrs("jrdate"))))%>
			   </td>
			   <td  style="width: 10%" colspan="6" bgcolor="#FFFF99" align="center"  height="35">
			        <b>操&nbsp;作&nbsp;人&nbsp;員</b>     
			   </td>
			   <td style="width: 15%" colspan="9">
			        <%=ReplaceSpace(QueryOperat(trim(returnrs("operat"))))%>
			   </td>
			</tr>
			<%
			end if
			 %>
            <tr>	
			    <td colspan="42" bgcolor="#00FFFF" height="20">
			        <b>&nbsp;&nbsp;裁罰資料</b>
			    </td> 
			</tr> 

            
<%
'response.write QuotedStr(request("BillNo"))
'response.end
strTmp="select * from PEO_New where tkt_no=" & QuotedStr(request("BillNo"))
set rsDetail2=conn.execute(strTmp)
    if  not rsDetail2.Eof then
        judeDate = rsDetail2("DES_DT")
        judeKeyinDate = rsDetail2("DKINDT")
		judeNo = rsDetail2("DESCNO")
		judePay = rsDetail2("DESPAY")
		judePubDate = rsDetail2("DSD_DT")
		judePubNo = rsDetail2("DSD_NO")
		judePubSN = rsDetail2("DPRNNO")
		judeOper = rsDetail2("D_OPER")
		judeSN = rsDetail2("DINSER")
		
		UrgeDate = rsDetail2("HUR_DT")
		UrgeKeyInDate = rsDetail2("HKINDT")
		UrgeNo = rsDetail2("HURYNO")
		UrgePay = rsDetail2("HURPAY")
		UrgePubDate = rsDetail2("HSD_DT")
		UrgePubNo = rsDetail2("HSD_NO")
		UrgePubSN = rsDetail2("HPRNNO")
		UrgeOper = rsDetail2("H_Oper")
		UrgeSN = rsDetail2("HINSER")
		
		MoveDate = rsDetail2("REM_DT")
		MoveKeyInDate = rsDetail2("RKINDT")
		MoveNo = rsDetail2("REM_NO")
		MovePay = rsDetail2("REMPAY")
		MovePubDate = rsDetail2("RSD_DT")
		MovePubNo = rsDetail2("RSD_NO")
		MovePubSN = rsDetail2("RPRNNO")
		MoveOper = rsDetail2("R_Oper")
		MoveSN = rsDetail2("RINSER")
    end if   
rsDetail2.close
set rsDetail2=nothing

%>            
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>裁&nbsp;決&nbsp;日</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(SetchinaDate(trim(judeDate)))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>建&nbsp;檔&nbsp;日</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(judeKeyinDate)))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>裁&nbsp;決&nbsp;字&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(judeNo))%>     
               </td>                                 
           </tr>
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>裁&nbsp;罰&nbsp;金&nbsp;額</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(trim(judePay))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>發&nbsp;文&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(judePubDate)))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>發&nbsp;文&nbsp;文&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(judePubNo))%>     
               </td>                                 
           </tr>
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>發&nbsp;文&nbsp;批&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(trim(judePubSN))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>建&nbsp;檔&nbsp;員</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(judeOper)) & " " & QueryOperat(trim(judeOper))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccff99" align="center">
                    <strong>建&nbsp;檔&nbsp;序&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(judeSN))%>     
               </td>                                 
           </tr>
           
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>催&nbsp;繳&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(SetchinaDate(trim(UrgeDate )))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>建&nbsp;檔&nbsp;日</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(UrgeKeyInDate )))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>催&nbsp;繳&nbsp;字&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(UrgeNo ))%>     
               </td>                                 
           </tr>
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>催&nbsp;繳&nbsp;金&nbsp;額</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(trim(UrgePay ))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>發&nbsp;文&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(UrgePubDate )))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>發&nbsp;文&nbsp;文&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(UrgePubNo ))%>     
               </td>                                 
           </tr>
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>催&nbsp;繳&nbsp;批&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(trim(UrgePubSN ))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>建&nbsp;檔&nbsp;員</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(UrgeOper )) & " " & QueryOperat(trim(UrgeOper))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ffcccc" align="center">
                    <strong>建&nbsp;檔&nbsp;序&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(UrgeSN))%>     
               </td>                                 
           </tr>
           
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>移&nbsp;送&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(SetchinaDate(trim(MoveDate  )))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>建&nbsp;檔&nbsp;日</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(MoveKeyInDate  )))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>移&nbsp;送&nbsp;字&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(MoveNo  ))%>     
               </td>                                 
           </tr>
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>移&nbsp;送&nbsp;金&nbsp;額</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(trim(MovePay  ))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>發&nbsp;文&nbsp;日&nbsp;期</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(SetchinaDate(trim(MovePubDate  )))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>發&nbsp;文&nbsp;文&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(MovePubNo ))%>     
               </td>                                 
           </tr>
           <tr>
                <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>移&nbsp;送&nbsp;批&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                     <%=ReplaceSpace(trim(MovePubSN ))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>建&nbsp;檔&nbsp;員</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(MoveOper )) & " " & QueryOperat(trim(MoveOper))%>    
               </td>
               <td  style="width: 10%; height: 41px;" colspan="6" bgcolor="#ccccff" align="center">
                    <strong>建&nbsp;檔&nbsp;序&nbsp;號</strong>
               </td> 
               <td style="width: 15%; height: 41px;" colspan="9">
                    <%=ReplaceSpace(trim(MoveSN))%>     
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
