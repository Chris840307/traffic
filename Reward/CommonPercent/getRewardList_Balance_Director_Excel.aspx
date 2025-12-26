<%@ Page Language="VB" %>
<% LoginCheck() %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

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
        Dim fname = Year(Now) & fMnoth & fDay & "_直接人員結餘.xls"
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
    <title>直接人員結餘款清冊</title>
</head>
<body>
    <form id="form1" runat="server">
    <table width="680" border="0" cellpadding="6" cellspacing="0" align="center">
        <td align="center" colspan="6"><%=Request("Date1")%>年 <%=Request("Date2")%>月   直接人員結餘款</td>
    </table>
    <table width="680" border="1" cellpadding="1" cellspacing="0" align="center">
        <td style="width: 40%" align="left">單位名稱</td>
        <td style="width: 20%" align="right">員警代碼</td>
        <td style="width: 20%" align="right">員警姓名</td>
        <td style="width: 20%" align="right">應領</td>
	<td style="width: 20%" align="right">實領</td>
	<td style="width: 20%" align="right">結餘</td>
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
            Dim MoneyTotal, SHOULDGETMONEY, REALGETMONEY As Decimal
            Dim OverFlag As String
            MoneyTotal = 0
            SHOULDGETMONEY = 0
            REALGETMONEY = 0
            Dim strUnit = "Select UNITID,LOGINID,CHNAME,SHOULDGETMONEY,REALGETMONEY from REWARDMONTHDATA where UNITID IN (" & request("sUnitID") & ") AND DIRECTORTOGETHER='1' AND SHOULDGETMONEY-REALGETMONEY>0 AND YEARMONTH between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
            Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
            If rdUnit.HasRows Then
                While rdUnit.Read()
                    Response.Write("<tr>")
                    Response.Write("<td>" & Trim(rdUnit("UNITID"))& "</td>")
      
                    SHOULDGETMONEY = (SHOULDGETMONEY) + cdbl(rdUnit("SHOULDGETMONEY"))
                    REALGETMONEY = (REALGETMONEY) + cdbl(rdUnit("REALGETMONEY"))
                    MoneyTotal = (MoneyTotal) + cdbl(rdUnit("SHOULDGETMONEY"))-cdbl(rdUnit("REALGETMONEY"))
                    Response.Write("<td align=""right"">" & Trim(rdUnit("LOGINID"))& "</td>")
                    Response.Write("<td align=""right"">" & Trim(rdUnit("CHNAME"))& "</td>")
                    Response.Write("<td align=""right"">" & Trim(rdUnit("SHOULDGETMONEY"))& "</td>")
		    Response.Write("<td align=""right"">" & Trim(rdUnit("REALGETMONEY"))& "</td>")
		    Response.Write("<td align=""right"">" & cdbl(rdUnit("SHOULDGETMONEY"))-cdbl(rdUnit("REALGETMONEY")) & "</td>")
                    Response.Write("</tr>")
                End While
            End If
            rdUnit.Close()
            conn.Close()
        %>
        <tr>
        <td>總計</td>
	<td></td>
	<td></td>
        <td align="right"><%=SHOULDGETMONEY%></td>
        <td align="right"><%=REALGETMONEY%></td>
        <td align="right"><%=MoneyTotal%></td>    

    </table>
    </form>
</body>
</html>
