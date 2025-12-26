<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%
    Dim fMnoth = Month(Now)
    If fMnoth < 10 Then fMnoth = "0" & fMnoth
    Dim fDay = Day(Now)
    If fDay < 10 Then fDay = "0" & fDay
    Dim fname = Year(Now) & fMnoth & fDay & ".xls"
    Response.AddHeader("Content-Disposition", "filename=" & fname)
    If Trim(Request("sMemID")) = "" Then
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
    End If
    Response.ContentType = "application/ms-excel"

    'Response.Clear()
    'Response.Buffer = True
    'Response.Charset = "utf-8"
    ''//下面这行很重要， attachment 参数表示作为附件下载，您可以改成 online在线打开 
    ''//filename=FileFlow.xls 指定输出文件的名称，注意其扩展名和指定文件类型相符，可以为：.doc 　　 .xls 　　 .txt 　　.htm　　 
    'Response.AppendHeader("Content-Disposition", "attachment;filename=FileFlow.xls")
    'Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
    ''//Response.ContentType指定文件类型 可以为application/ms-excel 　　 application/ms-word 　　 application/ms-txt 　　 application/ms-html 　　 或其他浏览器可直接支持文档　 
    'Response.ContentType = "application/ms-excel"
    'Me.EnableViewState = False

    Server.ScriptTimeout = 86400
    Response.Flush()
 %>
<script runat="server">
    Dim DBReward, RewardTotal, UnitReward, DBAllReward, UnitScore, UnitScoreTotal As Decimal
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
    
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        LoginCheck()
    End Sub
</script>
<script language="JavaScript">
	window.focus();
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>處理道路交通安全人員獎勵金分配統計表</title>
</head>
<body>
    <form id="form1" runat="server">
    <table border="1">
<% 
    '取得 Web.config 檔的資料連接設定
    Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
    '建立 Connection 物件
    Dim conn As New Data.OracleClient.OracleConnection()
    conn.ConnectionString = setting.ConnectionString
    '開啟資料連接
    conn.Open()
    Dim strUnitName As String = ""
    Dim strUnit = "select * from Apconfigure where id=40"
    Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
    Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
    If rdUnit.HasRows Then
        rdUnit.Read()
        strUnitName = Trim(rdUnit("value"))
    End If
    rdUnit.Close()
    
    Dim YearTmp, MonthTmp, BeginMonth, EndMonth As Integer
    For YearTmp = Trim(Request("Year1")) To Trim(Request("Year2"))
        Response.Write("<tr>")
        Response.Write("<td colspan=""7"" align=""center"" height=""30"">")
        Response.Write(strUnitName & Trim(Request("Year1")) & "年處理道路交通安全人員獎勵金分配統計表")
        Response.Write("</td>")
        Response.Write("</tr>")
        Response.Write("<tr>")
        Response.Write("<td width=""80"" height=""70"" align=""center"">月份</td>")
        Response.Write("<td width=""80"" align=""center"">總金額</td>")
        Response.Write("<td width=""95"" align=""center"">共同作業人員金額</td>")
        Response.Write("<td width=""95"" align=""center"">直接執行人員金額</td>")
        Response.Write("<td width=""95"" align=""center"">外勤人員領取人數</td>")
        Response.Write("<td width=""95"" align=""center"">單月領取最高之金額數(元)</td>")
        Response.Write("<td width=""95"" align=""center"">單月領取最低之金額數(元)</td>")
        Response.Write("</tr>")
        
        If Trim(Request("Year1")) = Trim(Request("Year2")) Then
            BeginMonth = Trim(Request("Month1"))
            EndMonth = Trim(Request("Month2"))
        ElseIf YearTmp = Trim(Request("Year1")) Then
            BeginMonth = Trim(Request("Month1"))
            EndMonth = 12
        ElseIf YearTmp = Trim(Request("Year2")) Then
            BeginMonth = 1
            EndMonth = Trim(Request("Month2"))
        Else
            BeginMonth = 1
            EndMonth = 12
        End If
        
        For MonthTmp = BeginMonth To EndMonth
            
            Dim strReward = "select * from RewardAnalyze where BeginDate between TO_DATE('" & (YearTmp + 1911) & "/" & MonthTmp & "/1','YYYY/MM/DD')"
            strReward = strReward + " and TO_DATE('" & DateAdd("d", -1, DateAdd("m", 1, (YearTmp + 1911) & "/" & MonthTmp & "/1")) & "','YYYY/MM/DD')"
            Dim CmdReward As New Data.OracleClient.OracleCommand(strReward, conn)
            Dim rdReward As Data.OracleClient.OracleDataReader = CmdReward.ExecuteReader()
            If rdReward.HasRows Then
                rdReward.Read()
                    
                Response.Write("<tr>")
                Response.Write("<td height=""40"" align=""center"">" & MonthTmp & "</td>")
                Response.Write("<td align=""center"">" & Trim(rdReward("RewardTotal")) & "</td>")
                Response.Write("<td align=""center"">" & Decimal.Truncate(Trim(rdReward("RewardTotal")) * 0.72) & "</td>")
                Response.Write("<td align=""center"">" & Decimal.Truncate(Trim(rdReward("RewardTotal")) * 0.28) & "</td>")
                Response.Write("<td align=""center"">" & Trim(rdReward("PeopleCount")) & "</td>")
                Response.Write("<td align=""center"">" & Trim(rdReward("MaxMoney")) & "</td>")
                Response.Write("<td align=""center"">" & Trim(rdReward("MinMoney")) & "</td>")
                Response.Write("</tr>")
                    
                
            End If
            rdReward.Close()

        Next
    Next
    
    conn.Close()
%>
    </table>
    </form>
</body>
</html>
