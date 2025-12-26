<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    LoginCheck()
%>
<script runat="server">
    Public UnitPoint, PointTotal, MoneyTotal, MemMoney, MemPay, UnitMoney As Decimal
    
    '將民國yymmdd轉換為yyyy/mm/dd
    Public Function gOutDT(ByVal iDate)
        Dim DatetTemp As String
        If iDate IsNot DBNull.Value Then
            DatetTemp = DateSerial(Left(iDate, Len(iDate) - 4) + 1911, Mid(iDate, Len(iDate) - 3, 2), Right(iDate, 2))
            gOutDT = DatetTemp
        Else
            gOutDT = ""
        End If
    End Function
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim fMnoth = Month(Now)
        If fMnoth < 10 Then fMnoth = "0" & fMnoth
        Dim fDay = Day(Now)
        If fDay < 10 Then fDay = "0" & fDay
        Dim fname = Year(Now) & fMnoth & fDay & "_作業項目結餘.xls"
        Response.AddHeader("Content-Disposition", "filename=" & Server.UrlPathEncode(fname))
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        Response.ContentType = "application/x-msexcel; charset=MS950"
    End Sub
    
    Sub LoginCheck()
        If (Request.Cookies("UserFunction") IsNot Nothing) Then
            Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
            If Trim(FuncCookie.Values("FuncID")) = "" Then
            Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
            End If
        Else
        Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
        End If
    End Sub
</script>
<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<style type="text/css">
<!--
body {font-family:新細明體;font-size:10pt; FONT-WEIGHT: bold }

.style1 {font-family:新細明體; font-size: 11pt}
-->
</style>
    <title>作業項目結餘款清冊</title>
</head>
<body>
    <form id="form1" runat="server">
    <table width="680" border="0" cellpadding="3" cellspacing="0" align="center">
        <td align="center" colspan="4"><span class="style1"><strong>作業項目結餘</strong></span></td>
	<tr>
        <td align="center" colspan="4"><span class="style1"><strong>統計範圍  <%=Request("Date1")%>  年  <%=Request("Date2")%> 月 </strong></span></td>
    </table>
    <table width="680" border="1" cellpadding="3" cellspacing="0" align="center">
        <td style="width: 100" align="left">作業項目</td>
        <td style="width: 50" align="center">應領</td>
        <td style="width: 50" align="center">實領</td>
        <td style="width: 50" align="center">結餘</td>
        </tr>
        <%
            '取得 Web.config 檔的資料連接設定
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
            '建立 Connection 物件
            Dim conn As New Data.OracleClient.OracleConnection()
            conn.ConnectionString = setting.ConnectionString
            '開啟資料連接
            conn.Open()
            
            Dim sys_City As String
            '要用填單或建檔日統計
            Dim theDateType As String = Trim(Request("DateType"))

	    dim TheAnaDate1 as String=""
            dim TheAnaDate2 as String=""
            TheAnaDate1 = (Trim(Request("Date1")) + 1911) & "/" & Trim(Request("Date2")) & "/1"
            TheAnaDate2 = DateAdd("d", -1, DateAdd("m", 1, (Trim(Request("Date1")) + 1911) & "/" & Trim(Request("Date2")) & "/1"))
 
            '===================列出清冊內容========================
            '---------------所有單位-----------------
            Dim V1, V2,V3,D1,D2 As Decimal
            
            Dim strUnit = "Select nvl(SUM(SHOULDGETMONEY),0) AS SHOULDGETMONEY2,nvl(SUM(REALGETMONEY),0) AS REALGETMONEY2 from REWARDMONTHDATA where UNITID IN (" & request("sUnitID") & ") AND DIRECTORTOGETHER='0' AND YEARMONTH between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

            Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
            Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
            If rdUnit.HasRows Then
                While rdUnit.Read()
                    Response.Write("<tr>")
                    Response.Write("<td width=""50"">共同作業</td>")
                    Response.Write("<td width=""50"">" & (rdUnit("SHOULDGETMONEY2")) & "</td>")
                    
                    V1 = CDbl(rdUnit("SHOULDGETMONEY2"))
                    V2 = CDbl(rdUnit("REALGETMONEY2"))
                    V3 = CDbl(rdUnit("SHOULDGETMONEY2")) - CDbl(rdUnit("REALGETMONEY2"))
                    Response.Write("<td  width=""50"" align=""center"">" & (rdUnit("REALGETMONEY2")) & "</td>")
        	    Response.Write("<td  width=""50"" align=""center"">" & v1-v2 & "</td>")
                    Response.Write("</tr>")
                End While
            End If
            rdUnit.Close()
            
            dim strUnit2 = "Select nvl(SUM(SHOULDGETMONEY),0) AS SHOULDGETMONEY2,nvl(SUM(REALGETMONEY),0) AS REALGETMONEY2 from REWARDMONTHDATA where DIRECTORTOGETHER='1' AND YEARMONTH between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdUnit2 As New Data.OracleClient.OracleCommand(strUnit2, conn)
            Dim rdUnit2 As Data.OracleClient.OracleDataReader = CmdUnit2.ExecuteReader()
            If rdUnit2.HasRows Then
                While rdUnit2.Read()
                    Response.Write("<tr>")
                    Response.Write("<td width=""50"">直接作業</td>")
                    Response.Write("<td width=""50"">" & (rdUnit2("SHOULDGETMONEY2")) & "</td>")
                    V1=cdbl(rdUnit2("SHOULDGETMONEY2"))+V1
                    V2=cdbl(rdUnit2("REALGETMONEY2"))+V2

		    D1=cdbl(rdUnit2("SHOULDGETMONEY2"))
                    D2=cdbl(rdUnit2("REALGETMONEY2"))

                    Response.Write("<td width=""50"" align=""center"">" & (rdUnit2("REALGETMONEY2")) & "</td>")
          	    Response.Write("<td width=""50"" align=""center"">" & D1-D2 & "</td>")
                    Response.Write("</tr>")
                End While

            End If
            rdUnit2.Close()


            conn.Close()
        %>
        <tr>
        <td>合計</td>
        <td align="right"><%=V1%></td>
        <td align="right"><%=V2%></td>
        <td align="right"><%=V1-V2%></td>

    </table>
    </form>
</body>
</html>
