<%
'strSql="select * from PasserBase where SN="&BillSN(i)
'set rsSql=conn.execute(strSql)
'Sys_Driver=trim(rsSql("Driver"))
'Sys_DriverAddress=trim(rsSql("DriverAddress"))
'rsSql.close
'strState="select UrgeDate from PasserUrge where BillSN="&BillSN(i)
'set rsState=conn.execute(strState)
'if not rsState.eof then
'	Sys_UrgeDate=split(gArrDT(trim(rsState("UrgeDate"))),"-")
'else
'	Sys_UrgeDate=split(gArrDT(trim("")),"-")
'end if
'rsState.close

Sys_Driver=trim(rssum("Driver"))
Sys_DriverID=trim(rssum("DriverID"))
Sys_DriverSex="女"
if trim(rssum("DriverSex"))="1" then Sys_DriverSex="男"
Sys_DriverAddress=trim(rssum("DriverAddress"))
Sys_DriverZip=trim(rssum("DriverZip"))
Sys_affair=trim(rssum("affair"))
Sys_FORFEIT1=trim(rssum("FORFEIT1"))
Sys_affair=trim(rssum("affair"))
%>
<table width="645" height="90%" border="1" cellspacing=0 cellpadding=0>
  <tr>
    <td height="67" colspan="4"><span class="style14">台 端 違 反 道 路 交 通 管 理 處 罰 條 例<br>
    罰 鍰 未 繳 ， 寄 存 送 達 通 知 書 。</span></td>
  </tr>
  <tr>
    <td width="104" height="43"><span class="style18">事　　由</span></td>
    <td colspan="3" class="style18">違反道路交通管理事件處罰案</td>
  </tr>
  <tr>
    <td height="47"><span class="style18">送達文件</span></td>
    <td colspan="3"><span class="style18">裁決書</span></td>
  </tr>
  <tr>
    <td height="49"><span class="style18">受送達人<br>
    姓　　名</span></td>
    <td colspan="3"><span class="style18"><%=Sys_Driver%></span></td>
  </tr>
  <tr>
    <td height="41"><span class="style18">送達處所</span></td>
    <td colspan="3"><span class="style18"><%=Sys_DriverZip&Sys_DriverAddress%></span></td>
  </tr>
  <tr>
    <td height="46"><span class="style18">送達時間</span></td>
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="102" colspan="4"><span class="style16">因未獲會晤本人亦無受領文書之同居人、受雇人，請將此通知書貼於門牌號旁並拍照，照片交回，本書不用繳回</span></td>
  </tr>
  <tr valign="top">
    <td height="223" colspan="4"><span class="style17">因未獲會晤本人亦無受領文書之同居人、受雇人或應送達處所之接收郵件人，已將該送達文者寄存於<br>
    請於　　日內連絡領取。</span></td>
  </tr>
  <tr>
    <td height="92"><div align="center"><span class="style15">送 達<br>
    單 位</span></div></td>
    <td width="191">&nbsp;</td>
    <td width="118"><div align="center"><span class="style15">送達人</span></div></td>
    <td width="213">&nbsp;</td>
  </tr>
  <tr>
    <td height="145" colspan="4"><p class="style15">&nbsp;</p>
    <p class="style15">中華民國　　　　　年　　　　　月　　　　　日</p></td>
  </tr>
</table>
