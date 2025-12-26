<%@ Page Language="VB" AutoEventWireup="false" Debug=true CodeFile="Main.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>績效獎勵金試算系統</title>
</head>
<body>
    <form id="form1" runat="server">
    <table width="1000" border="0" align="center" style="top:100px">
    <tr><td style="height:25px">
        <div id="Layer1" style="position:absolute; z-index:0; top:3px; left:225px ; height:34px; width: 589px">
            <asp:Label ID="LbUserData" runat="server" Font-Bold="True" Height="30px" Width="477px" Font-Size="Large" ForeColor="Black"></asp:Label>
            <asp:Button ID="Button1" runat="server" Height="26px" Text="登出" Width="57px" onclick="UserLogout"/>
        </div>
    </td></tr>
    <tr><td>
    
    <%  
        Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")
        Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        
        '----------系統選擇區塊-----------------------
        Response.Write("<table top='100' width='100%' border='0' align='center' style='top:100px'>")
        Response.Write("<tr><td width='200' height='1'>")
        Response.Write("<td width='200' height='1'>")
        Response.Write("<td width='200' height='1'>")
        Response.Write("<td width='200' height='1'>")
        Response.Write("<td width='200' height='1'></tr>")
        Dim s As Integer = 1
        Dim picName, SystemName As String
        Dim strFunc = "select a.* from FunctionPageDataDotNet a,FunctionData b where a.SystemID=b.SystemID and b.GroupID='" & Trim(UserCookie.Values("GroupRoleID")) & "' and b.Function='1' order by ShowOrder"
        Dim CmdFunc As New Data.OracleClient.OracleCommand(strFunc, conn)
        Dim rdFunc As Data.OracleClient.OracleDataReader = CmdFunc.ExecuteReader()
        If rdFunc.HasRows Then
            While rdFunc.Read()
                If s = 1 Then
                    Response.Write("<tr>")
                End If
                If IsDBNull(rdFunc("ImageLocation")) Then
                    picName = "tmp.jpg"
                Else
                    picName = rdFunc("ImageLocation")
                End If
                SystemName = ""
                Dim strSName = "select * from Code where ID=" & Trim(rdFunc("SystemID"))
                Dim CmdSName As New Data.OracleClient.OracleCommand(strSName, conn)
                Dim rdSName As Data.OracleClient.OracleDataReader = CmdSName.ExecuteReader()
                If rdSName.HasRows Then
                    rdSName.Read()
                    SystemName = Trim(rdSName("Content"))
                End If
                rdSName.Close()
                

                Response.Write("<td width='200' height='190'>")
                Response.Write("<div id='" & rdFunc("SystemID") & "' style='position:absolute; width:160px; height:170px; z-index:1 ;'>")
                Response.Write("<table id='table" & rdFunc("SystemID") & "' width='100%' border='0' align='center' >")
                Response.Write("<tr><td id='td1' align='center'>")
                Response.Write("<a onclick=""OpenSystem('" & rdFunc("URLLocation") & "','" & rdFunc("SystemID") & "');"" onMouseOver=""DivColorChange('table" & rdFunc("SystemID") & "');"" onMouseOut=""DivColorChange2('table" & rdFunc("SystemID") & "');"">")
                Response.Write("<img src='image/" & picName & "' alt='' width='128' height='128' border='0' align='baseline'>")
                Response.Write("<br><font size='4'>")
                Response.Write(SystemName & "</font></a>")
                Response.Write("</td></tr></table>")
                Response.Write("</div></td>")
                
                If s = 5 Then
                    Response.Write("</tr>")
                    s = 1
                Else
                    s = s + 1
                End If

            End While
        End If
        rdFunc.Close()
        Response.Write("</table>")
        
        conn.Close()
    %>
    </td></tr>
    </table>
    </form>
</body>
<script language="JavaScript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
}

function OpenSystem(PageUrl,Sn){
	SCheight=screen.availHeight;
	SCWidth=screen.availWidth;
	UrlStr=PageUrl;
	newWin(UrlStr,Sn,SCWidth,SCheight,0,0,"yes","yes","yes","no");
}
function DivColorChange(DivNo){
	eval(DivNo).border="1";
}
function DivColorChange2(DivNo){
	eval(DivNo).border="0";
}
</script>
</html>
