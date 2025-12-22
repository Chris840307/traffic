<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--"traffic/include/ActiveXInclude.htm"-->
<!--"traffic/include/FunctionInclude.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="traffic/Common/css.txt"-->
<title>宏謙科技實業有限公司-入案管理系統</title>
<script type="text/javascript" src="/js/date.js"></script>
<script type="text/javascript" src="/js/form.js"></script>
<%
Response.AddHeader "X-XSS-Protection", "1; mode=block"

strAuthority = GenPkiTicket
session("Ticket") = strAuthority

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

'	if sys_City="台中縣" then
'		strDelLog="delete from Log where SN>133676"
'		conn.execute strDelLog
'	end if

if UseIpSelectUrlLocationFlag=1 then
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

	dim fso1 
	set fso1=Server.CreateObject("Scripting.FileSystemObject")
	UpdateBillFile = Server.MapPath("/traffic/Common/ActionIp.ini")	' 找出記數檔案在硬碟中的實際位置
	Set Out1= fso1.OpenTextFile(UpdateBillFile, 1, FALSE)	' 開啟唯讀檔案
	if not Out1.atEndOfStream then
		strUpdateBillUser = Out1.ReadAll	' 從檔案讀出記數資料
	end if

	Out1.Close	' 關閉檔案
	set fso1=nothing
	'如果使用者ip沒在文件上就轉址到ap
	goOtherUrl=1
	if strUpdateBillUser<>"" then
		UpdBUserArray=split(strUpdateBillUser,vbNewLine)
		for SU=0 to ubound(UpdBUserArray)
			if UpdBUserArray(SU)=trim(userip) then
				goOtherUrl=0
				exit for
			end if
		next
	end if
	'如果Server IP=ap ip就不轉
	if instr(UseIpSelectUrlLocation,trim(Request.ServerVariables("Local_ADDR") ))>0 then
		goOtherUrl=0
	end if
	if goOtherUrl=1 then
		response.Redirect UseIpSelectUrlLocation
	end if
end if

if trim(request("ErrorS"))<>"" then
	errstring=Replace(Replace(Replace(replace(replace(replace(request("ErrorS"),"<",""),">",""),"/",""), """" ,""), "'" ,""), "&" ,"")
	if instr(trim(request("ErrorS")),"<")>0 or instr(trim(request("ErrorS")),">")>0 or instr(trim(request("ErrorS")),"/")>0 or instr(trim(request("ErrorS")),"&")>0  then

	else
%>
<script language="JavaScript">
    alert ("<%=errstring%>!!");
<%
	
if trim(request("UpdM")&"")="1" then
	'Response.Redirect "UserDataEdit.asp"
%>
	location.href="UserDataEdit.asp";
	//window.open("UserDataEdit.asp","UserDataEdit","width=800,height=500,left=50,top=10,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no");
<%
	end if
end if 
%>
</script>
<%
end if
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<style type="text/css">
<!--
.style4 {font-family: "新細明體"}
#Layer2 {
	position:absolute;
	width:177px;
	height:26px;
	z-index:2;

}
.style5 {
	font-size: 20px;
	line-height:24px;
	font-weight: bold;
}
#Layer3 {
	position:absolute;
	width:177px;
	height:26px;
	z-index:2;

}
-->
</style>
</head>




<!DOCTYPE html>
<html lang="zh-cmn-Hant">

<head>

  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  <meta name="description" content="">
  <meta name="author" content="">

  <title></title>

  <meta name="csrf-param" content="authenticity_token" />
<meta name="csrf-token" content="BrDsC0Qw5Ms4jcQPZOI4xE7tWa4mACO3P+lHoM58CElTBiIY1OxWbFaGuib0n/ZETjRupxXjUzvX4N43gBSUHQ==" />


  <link rel="stylesheet" media="all" href="admin-92e995e469ea98c880e61710f498cb7c0dddcd185d591b92bc985fb93e14d29a.css" />
<script src='https://www.google.com/recaptcha/api.js'></script>
</head>

  <body class="bg-dark" style="background-color:#0abab5 !important">
    <div class="container">
      <div class="report_session card card-login mx-auto mt-5">
  <div class="card-header text-center" style="background-color:white!important;border-style: none;border-radius:150px 150px 150px 150px;">
    <h4 style="color:#0abab5 !important;font-weight: bold;">宏謙科技實業有限公司<br>入案管理系統</h4>
  </div>
  <div class="card-body">
     <form name="myForm" method="post" action="UserLogin_Contral.asp" onsubmit="return User_Login();"><input name="utf8" type="hidden" value="&#x2713;" /><input type="hidden" name="authenticity_token" value="fOaOkWnGBKG6VzKZy/wnj8sj8hhFG8UoFViWZjua0KejBr8zXYc7UJsuxcsbeW4UE++syLM/VIy1TUMBuoS/OA==" />
      <div class="form-group">
        <div class="form-label-group">
<input name="MemberID" type="text" value="" size="18" maxlength="10" onkeyup="value=value.toUpperCase()" onkeydown="if(event.keyCode==13)event.keyCode=9" class="form-control string required form-control" autofocus="autofocus" required="required" aria-required="true">
          
          <label>帳號名稱</label>
        </div>
      </div>
      <div class="form-group">
        <div class="form-label-group">
<input name="MemberPW" type="password" value="" size="18" class="form-control password required form-control" required="required" aria-required="true">
          <font color="red">密碼長度須8個字元以上、並含四種字元（英文大小寫、數字、特殊符號）中的3種</font></small>
          <label>密碼</label>
        </div>
      </div>
      <div class="form-group">
        <div class="checkbox">
          <label>
            <input value="1" name="admin_user[remember_me]" type="checkbox" data-parsley-multiple="checkbox-signup" class="peer">
            記住我一週
          </label>
        </div>
      </div>
      
      <button class="btn btn-primary btn-block btn-login">登入</button>
</form>  </div>
</div>


<style type="text/css">
  span.hint{
    display: none;
  }
</style>



    </div>

    <script src="admin-7e641842b7678866dba9f029b1984fc78978fbe3c300f84802bb98e82b1f6905.js"></script>

  </body>



</html>


<%
 	

	if sys_City="台南市" then
%>
    <div align="center">
		<a href="分局執法系統使用者申請表.doc" target="_blank" >分局執法系統使用者申請表</a> &nbsp;
		<a href="交大執法系統使用者申請表.doc" target="_blank" >交大執法系統使用者申請表</a>
		<br>
		 
    </div>
<%
	end if
	if sys_City="彰化縣" then
%>
    <div align="center">
		<a href="智慧型交通執法管理系統-使用者(異動)申請表-1121002.doc" target="_blank" style="font-size: 20pt;line-height:26px;" >智慧型交通執法管理系統-使用者(異動)申請表.doc</a> <br />
		<a href="智慧型交通執法系統帳號管理表-1121002.doc" target="_blank" style="font-size: 20pt;line-height:26px;" >智慧型交通執法系統帳號管理表.doc</a>
		<br>

    </div>
<%
	end if
%>

<%if sys_City="基隆市" then%>
	<div align="center">
		<a style="font-size: 30pt;line-height:40px;color: #FF0000;">** 使用者帳號改為知識聯網帳號，如無法登入請洽詢交通隊或分局承辦人 **</a>
    </div>
<%end if%>
<div>&nbsp;</div>
	<!--<div align="center">
		<a href="Edgeset.asp" target="_blank"  style="font-size: 20pt;line-height:26px;">** Edge瀏覽器相容性設定方式 **</a>
    </div>

	<div align="center">
		<a href="EdgeSafeSet.asp" target="_blank"  style="font-size: 20pt;line-height:26px;">** Edge瀏覽器安全性設定方式 **</a>
    </div>

	<div align="center">
		<a href="建檔系統處理方法.pdf" target="_blank"  style="font-size: 20pt;line-height:26px;">** 建檔系統超速法條帶入問題處理方式 **</a>
    </div>
	 <div align="center">
		<a href="setprint01.html" target="_blank"  style="font-size: 20pt;line-height:26px;">** Edge舉發單列印設定方式 **</a>
    </div> -->
	<input type="Hidden" name="PKICarchk" value="">

</body>
<script language="JavaScript">
function User_Login(){
  error=0;
  if (myForm.MemberID.value==""){
    alert ("請輸入身份證號碼!!");
    error=1;
<%if sys_City="高雄市" or sys_City="台南市" then%>
  }else if(myForm.MemberID.value.indexOf('\'')>=0 || myForm.MemberID.value.indexOf('&')>=0 || myForm.MemberID.value.indexOf('<')>=0 || myForm.MemberID.value.indexOf('>')>=0 || myForm.MemberID.value.indexOf('"')>=0 || myForm.MemberID.value.indexOf('--')>=0){
	alert("身份證號碼不可使用下列特殊符號『& , < , > , \" , ' , --』");
	error=1;
<%end if %>
  }else if(myForm.MemberPW.value==""){
	alert ("請輸入密碼!!");
    error=1;
  }<%if sys_City="高雄市" or sys_City="台南市" then%>else if(myForm.MemberPW.value.indexOf('\'')>=0 || myForm.MemberPW.value.indexOf('&')>=0 || myForm.MemberPW.value.indexOf('<')>=0 || myForm.MemberPW.value.indexOf('>')>=0 || myForm.MemberPW.value.indexOf('"')>=0 || myForm.MemberPW.value.indexOf('--')>=0){
	alert("使用者密碼不可使用下列特殊符號『& , < , > , \" , ', --』");
	error=1;
  }
 <%end if %>
  if (error==0){
    return true;
  }else{
    myForm.MemberID.focus();
    return false;
  }
}


function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
}
</script>
<script language="VBScript">
Sub SignICCard()
'	Set atxCert = createobject("AresPKIAtx.AtxCertificate")
'	Set atxUtility = createobject("AresPKIAtx.AtxUtility")
'	Set cms = createobject("AresPKIAtx.AtxCmsSignedData")
'	AresPKIClient.setLicense "9a6d220031dad1702592b900e497c95df5c864d0848b0e26659fc6721d4979b62373ccfb46ed64a7e7ebfa6f80e9a498d1f70268e58d39042bb0282861b991ac8ad0a331a241f450b1cc0f8c270335a1e97f115a834ac5ba455095e38cb318dffa6e1db9662e22406ec1dc7aa3770d4adb091798170f1dc380fc3d7783de375a","聯宏科技股份有限公司"
'    ticket = "<%=strAuthority%>"
'    nRet = AresPKIClient.Init()    
'    if nRet <> 0 then
'		msgbox(AresPKIClient.GetErrorMessage())
'		AresPKIClient.Finalize()
'		exit sub
'    end if
'	nRet = AresPKIClient.EncodeP7SignedData(ticket,"",pSignedData,"")
'    if nRet <> 0 then
'		msgbox(AresPKIClient.GetErrorMessage())
'		AresPKIClient.Finalize()
'		exit sub
'    end if     
'    AresPKIClient.Finalize()
'        
'    If pSignedData <> "" Then
'        pSignedData = AresPKIClient.HexStringToB64(pSignedData)
'	    'msgbox pSignedData
'    End If
'
'	encodeData = atxUtility.BSTR_B64ToBin(atxUtility.BSTR_WideCharToMultiByte(pSignedData))
'	result = cms.InitDecode(encodeData,"")
'	binary=cms.Decode()
'	'For j = lbound(binary) To ubound(binary)
'		'msgbox(binary(j))
'	'Next
'
'	'取得原始資料
'	'msgbox "GetDecodeContent()=================================="
'	'msgbox(atxUtility.BSTR_MultiByteToWideChar(cms.GetDecodeContent()))
'
'	'取得憑證
'	certs=cms.GetDecodeCertificates()
'	cert = certs(0)
'	cms.FinalDecode()
'
'	'驗證憑證
'	atxCert.BinaryCert = cert
'	'msgbox("---有效日期自(double byte string)：")
'	'msgbox(atxCert.FromDate)
'	'msgbox("---有效日期自(long)：")
'	'msgbox(atxCert.FromDateBinary)
'	'msgbox("---有效日期至(double byte string)：")
'	'msgbox(atxCert.ToDate)
'	'msgbox("---有效日期至(long)：")
'	'msgbox(atxCert.ToDateBinary)
'	if now>atxCert.ToDate then
'		msgbox("此卡片已過期")
'		exit sub
'	end if
'
'	nRet = AresPKIClient.Init("aetpkss1.dll")
'	nRet = AresPKIClient.GetCertificate(0,cert)
'	AresPKIClient.Finalize()
'	nRet = AresPKIClient.DecodeCertificate(cert)
'	certSN = AresPKIClient.GetCertSubjectSN()
'	certSN = certSN + AresPKIClient.GetCertIssuerName()
'	certHex = AresPKIClient. B64ToHexString(AresPKIClient.B64Encode(certSN))
'	myForm.PKICarchk.value=certHex
'	myForm.submit()
End Sub

</script>
</html>
