<%

strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

strSql="select Value from Apconfigure where ID=31"
set rslawt=conn.execute(strSql)
        City=trim(rslawt("Value") & "")
Set rslawt=Nothing

strSql="select Value from Apconfigure where ID=35"
set rslawt=conn.execute(strSql)
        Sys_theUnit=trim(rslawt("Value") & "")
Set rslawt=Nothing

strSql="select Value from Apconfigure where ID=52"
set rslawt=conn.execute(strSql)
        Sys_BookUnit=trim(rslawt("Value") & "")
Set rslawt=Nothing

UnitNo=Session("Unit_ID")

strUnit="select UnitName from UnitInfo where UnitID='"&UnitNo&"'"
set rsUnit=conn.execute(strUnit)
UnitName=rsUnit("UnitName")
rsUnit.close
set rsUnit=nothing

strState="select a.billunitid,a.Driver,a.DealLineDate,a.rule1,Substr(a.rule1,1,2) as Rule1_1,Substr(a.rule1,3,1) as Rule1_2,Substr(a.rule1,4,2) as Rule1_3,a.rule2,Substr(a.rule2,1,2) as rule2_1,Substr(a.rule2,3,1) as rule2_2,Substr(a.rule2,4,2) as rule2_3,nvl(forfeit1,0) forfeit1,nvl(forfeit2,0) forfeit2,to_char(a.IllegalDate,'HH24') as tHH,to_char(a.IllegalDate,'MI') as tMI,b.Level1,b.Level2,a.IllegalDate,a.IllegalAddress,c.OpenGovNumber,c.JudeDate from Passerbase a,law b,PasserJude c where a.rule1=b.itemid and b.version=2 and a.Sn=c.BillSn(+) and SN="&BillSN(i)
set rsState=conn.execute(strState)

Sys_ForFeit1=0:Sys_ForFeit2=0
Sys_rule2="":Sys_rule2_1="":Sys_rule2_2="":Sys_rule2_3=""
if not rsState.eof Then
	UOpenGovNumber=trim(rsState("OpenGovNumber"))
	Sys_Driver=trim(rsState("Driver"))
	Sys_rule1=trim(rsState("Rule1"))
	Sys_rule1_1=trim(rsState("Rule1_1"))
	Sys_rule1_2=trim(rsState("Rule1_2"))
	Sys_rule1_3=trim(rsState("Rule1_3"))

	If not ifnull(trim(rsState("Rule2"))) Then
		Sys_Rule2=trim(rsState("Rule2"))
		Sys_Rule2_1=trim(rsState("Rule2_1"))
		Sys_Rule2_2=trim(rsState("Rule2_2"))
		Sys_Rule2_3=trim(rsState("Rule2_3"))

		strSQL="select ItemID,Level1,Level2,Level3,Level4 from law where version="&RuleVer&" and itemid='"&trim(rsState("Rule2"))&"'"
		set rslaw=conn.execute(strSQL)
		If not rslaw.eof Then
			If DateDiff("d",CDate(date),trim(rsState("DealLineDate")))>-1 Then 
			  Sys_ForFeit2=trim(rslaw("Level1"))
			Else
			  Sys_ForFeit2=trim(rslaw("Level2"))
			End If 
		End if 
		rslaw.close		
	End if 

	If not ifnull(rsState("JudeDate")) Then

		UJudeDate=split(gArrDT(rsState("JudeDate")),"-")
	else
		UJudeDate=split("--","-")
	
	End if 
	
	Sys_ForFeit1=0:Sys_ForFeit2=0

    Sys_ForFeit1=rsState("ForFeit1")
	Sys_ForFeit2=rsState("ForFeit2")

    If trim(rsState("tHH"))<>"" then
		Sys_tHH=cint(rsState("tHH"))
	Else
		Sys_tHH=trim(rsState("tHH"))
	End If
	If trim(rsState("tMI"))<>"" then
		Sys_tMI=trim(rsState("tMI"))
	Else
		Sys_tMI=trim(rsState("tMI"))
	End If 

    sum_ForFeit=cdbl(Sys_ForFeit1)+cdbl(Sys_ForFeit2)

	Sys_IllegalAddress=trim(rsState("IllegalAddress"))
    Sys_Date=split(gArrDT(trim(rsState("IllegalDate"))),"-")	

end if
rsState.close
set rsState=Nothing

strLaw="":strLaw2=""
sqlLaw="select illegalrule from law where version="&RuleVer&" and itemid='"&Sys_rule1&"'"
set rsLaw=conn.execute(sqlLaw)
If not rsLaw.eof Then strLaw=rsLaw("illegalrule")
rsLaw.Close

If not ifnull(Sys_rule2) Then
	sqlLaw="select illegalrule from law where version="&RuleVer&" and itemid='"&Sys_rule2&"'"
	set rsLaw=conn.execute(sqlLaw)
	If not rsLaw.eof Then strLaw2=rsLaw("illegalrule")
	rsLaw.Close
End if 

chkLaw=0
chkLaw2=0
%>
<BR><BR>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="599" id="AutoNumber1" height="85%">
  <tr>
    <td width="686" height="31" colspan="4" valign="middle">
    <p align="center"><font size="5" face="標楷體"><%=Sys_theUnit%><%=UnitName%>違反道路交通管理事件簽辦單</font></td>
  </tr>
  <tr>
    <td width="686" height="31" colspan="4" valign="middle">
    <p align="center"><font size="5" face="標楷體">違　　　反　　　事　　　實</font></td>
  </tr>
  <tr>
    <td width="686" height="535" colspan="4"><span align="center"><font face="標楷體" size="4">受處分人<%=Sys_Driver%>於　<%=Sys_Date(0)%>年　<%=cdbl(Sys_Date(1))%>月　<%=cdbl(Sys_Date(2))%>日　<%=Sys_tHH%>時　<%=cdbl(Sys_tMI)%>分，在<%=replace(trim(Sys_IllegalAddress),"台","臺")%> </font></span>
    <p><font face="標楷體">
    <br><font size="4">&nbsp;　
			<%If Sys_rule1="7410101" Then 
				response.write "&#9632;" 
				chkLaw=1
			elseif Sys_Rule2="7410101" then
				response.write "&#9632;" 
				chkLaw2=1
			Else 
				response.write "□"
			end if%> 一、慢車駕駛人不服從交通警察指揮或不依標誌、標線、號誌之指示。

    <br>　　<%If Sys_rule1="7410301" Then 
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="7410301" then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 二、慢車不依規定擅自穿越快車道者。

    <br>　　<%If Sys_rule1="81101011" Or Sys_rule1="81101021" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="81101011" Or Sys_Rule2="81101021" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 三、於鐵路，公路車站或交通頻繁處所，違規攬客或妨害交通序者。

<br>	　　<%If Sys_rule1="8210101" Then
				response.write "&#9632;" 
				chkLaw=1
			elseif Sys_Rule2="8210101" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else 
				response.write "□"
			end if%> 四、在道路堆積、置放、設置或拋擲足以妨礙交通之物者。

    <br>　　<%If Sys_rule1="8210201" Then 
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8210201" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else 
				response.write "□"
			end if%> 五、在道路兩旁附近燃燒物品發生濃煙足以妨礙行車路線。
    <br>　　<%If Sys_rule1="8210301" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8210301" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 六、利用道路作為工作場所。
    <br>　　<%If Sys_rule1="8210401" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8210401" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else 
				response.write "□"
			end if%> 七、利用道路放置拖車、貨櫃或動力機械者。
    <br>　　<%If Sys_rule1="8210501" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8210501" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
			response.write "□"
			end if%> 八、興修房屋使用道路未經許可或經許可超出限制。
    <br>　　<%If Sys_rule1="8210601" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8210601" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 九、經主管機關許可挖掘道路而不依規定樹立警告標誌或於事後未將
    <br>　　　　 　障礙物清除。
    <br>　　<%If Sys_rule1="8210701" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8210701" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 十、擅自設置或變更交通標誌、標線或其類似之標識。
    <br>　　<%If Sys_rule1="8210801" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8210801" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 十一、未經許可在道路置石碑、廣告牌、綵坊或其他類似物。
    <br>　　<%If Sys_rule1="8210901" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8210901" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 十二、未經許可在道路舉行賽會或擺設筵席、演戲、拍攝電影或其他類
    <br>　　　　　 　似行為。
    <br>　　<%If Sys_rule1="8211001" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8211001" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 十三、未經許可在道路擺設攤位。
    <br>　　<%If Sys_rule1="8310101" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8310101" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 十四、在車道或交通島上散發廣告物、宣傳單、或其他相類之物。
    <br>　　<%If Sys_rule1="8310201" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8310201" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 十五、在車道上、車站內、高速公路服務區休息站，任意販賣物品妨礙
    <br>　　　　 　　交通。
    <br>　　<%If Sys_rule1="8410101" Or Sys_rule1="8400001" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="8410101" Or Sys_Rule2="8400001" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 十六、疏縱或牽繫禽、畜、寵物在道路奔走妨害交通。
	<br>　　<%If Sys_rule1="7810301" Then
				response.write "&#9632;"
				chkLaw=1
			elseif Sys_Rule2="7810301" Then
				response.write "&#9632;" 
				chkLaw2=1
			Else
				response.write "□"
			end if%> 十七、行人不依規定擅自穿越車道。
    <br>　　<%If chkLaw = 0 Then
				chkLaw=1
				response.write "&#9632; 十八、其他："&strLaw
			elseif Sys_Rule2<>"" and chkLaw2 = 0 Then
				chkLaw2=1
				response.write "&#9632; 十八、其他："&strLaw2
			Else
				response.write "□ 十八、其他："
			end if%>
    <br>　　<%If chkLaw = 0 Then
				chkLaw=1
				response.write "&#9632; 十九、其他："&strLaw
			elseif Sys_Rule2<>"" and chkLaw2 = 0 Then
				chkLaw2=1
				response.write "&#9632; 十九、其他："&strLaw2
			end if%>
	</font></font><br>
    <font face="標楷體" size="4">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    <br>　　
    <br>　　
    <br>　　</font><br><br></td>
  </tr>
  <tr>
  <td colspan="4" width="686" height="44" >
 <table border="0" cellspacing="0" width="640" cellpadding="0" bordercolor="#111111" style="border-collapse: collapse" height="62"> 
    <td width="81" height="21" align="center" style="border-bottom-style: solid; border-bottom-width:thin; border-right-style:solid; border-right-width:thin">
      <font size="4" face="標楷體">適用法條</font><font size="4"> </font>
    </td>
    <td width="549" height="21" colspan="3" style="border-bottom-style: solid; border-bottom-width:thin;">
    <font face="標楷體" size="4">依道路交通管理處罰條例<%
		If Sys_rule1="81101011" Or Sys_rule1="81101021" Then
			Response.Write "第　"&Sys_rule1_1&" 條之1　"
		else
			Response.Write "第　"&Sys_rule1_1&" 條第　"&cdbl(Sys_rule1_2)&"　項"

			If cdbl(Sys_rule1_3) > 0 Then Response.Write "第 "&cdbl(Sys_rule1_3)&"　款"

		End if 
		if Sys_Rule2<>"" then 
			Response.Write "與第"&Sys_rule2_1&"條第"&Sys_rule2_2&"項"
			If cdbl(Sys_rule2_3) > 0 Then Response.Write "第"&cdbl(Sys_rule2_3)&"款"
		end If 
	%>規定裁決。</font></td>
  </tr>
  <tr>
    <td width="81" height="22" align="center" style="border-bottom-style: solid; border-bottom-width:thin; border-right-style:solid; border-right-width:thin">
    <font face="標楷體" size="4">擬處意見</font></td>
    <td width="549" height="22" colspan="3" style="border-bottom-style: solid; border-bottom-width:thin; ">
    <font face="標楷體" size="4">擬處罰鍰新臺幣<%=sum_ForFeit%> 元，□攤架、□招牌沒入。</font></td>
  </tr>
  <tr>
    <td width="81" height="19" align="center" nowrap style=" border-right-style:solid; border-right-width:thin">
    <font face="標楷體" size="4">裁決書文號</font></td>
    <td width="212" height="19" style="border-right-style:solid; border-right-width:thin">

    <font face="標楷體" size="3">
<%if City="澎湖縣" and UnitNo="D01" then 
response.write "馬警"
end if
%>
<%=BillPageUnit%><%=UrgeNo%><%=UOpenGovNumber%>號</font></td>
    <td width="151" height="19" align="center" style="border-right-style:solid; border-right-width:thin">
    <font face="標楷體" size="4">裁決日期</font></td>
    <td width="178" height="19" ><font face="標楷體" size="4">　<%=UJudeDate(0)%>年<%=UJudeDate(1)%>月<%=UJudeDate(2)%>日</font></td>
</table>    
</td>
  </tr>
  <tr>
    <td width="160" height="19" align="center"><font face="標楷體" size="4">承　　辦　　人</font></td>
    <td width="160" height="19" align="center"><font face="標楷體" size="4">組　　　　　長</font></td>
    <td width="160" height="19" align="center"><font face="標楷體" size="4">副　分　局　長</font></td>
    <td width="160" height="19" align="center"><font face="標楷體" size="4">批　　　　示</font></td>
  </tr>
  <tr>
    <td width="160" height="69" align="center">　</td>
    <td width="160" height="69" align="center">　</td>
    <td width="160" height="69" align="center">　</td>
    <td width="160" height="69" align="center">　</td>
  </tr>
</table>