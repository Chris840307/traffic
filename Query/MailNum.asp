<body>
<form name="form1" action="https://ipost.post.gov.tw/CSController?cmd=POS0000_3&_MENU_ID=189&_ACTIVE_ID=189&_SYS_ID=D" method="post">
<input type="hidden" name="LookUptype" value="domestic_bundle_register">
<input type="hidden" name="apID" value="LIQD" >
<input type="hidden" name="kind" value="15" >
                <td><input type="hidden" name="MailNum_1" value="<%=trim(request("Mailnum"))%>"></td>
</form >
</body>
<script language="javascript">
//http://220.128.140.193/traffic/Query/MailNum.asp?MailNum=2608254000018
	document.form1.submit();
</script>
