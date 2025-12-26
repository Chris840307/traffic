<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language="vb" runat="server" src="../LoginCheck.vb" />
<%  
    LoginCheck()
    Server.ScriptTimeout = 86400
    Response.Flush()
%>
<object id="factory" style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<script runat="server">
    Dim DBReward, RewardTotal, UnitReward As Decimal
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
</script>
<script language="JavaScript">
	window.focus();
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <style type="text/css">
    <!--
    body {font-family:新細明體;font-size:10pt; FONT-WEIGHT: bold }

    .style1 {font-family:新細明體; font-size: 11pt}
    -->
    </style>
    <style media=print>
    .Noprint{display:none;}
    .PageNext{page-break-after: always;}
    </style>
    <title>共同人員獎勵金核發清冊</title>
</head>
<body>
    <form id="form1" runat="server">
    <table width="680" border="0" cellpadding="3" cellspacing="0" align="center">
        <tr>
        <td align="center" colspan="4"><span class="style1"><strong>共&nbsp;同&nbsp;人&nbsp;員&nbsp;支&nbsp;領&nbsp;獎&nbsp;勵&nbsp;金&nbsp;核&nbsp;發&nbsp;清&nbsp;冊</strong></span></td>
        </tr>
        <tr>
        <td style="width: 15%"></td>
        <td style="width: 35%">單位名稱</td>
        <td style="width: 25%" align="right">金額</td>
        <td style="width: 25%"></td>
        </tr>
    </table>
    <hr size="3" />
    <table width="680" border="0" cellpadding="3" cellspacing="0" align="center">
        <%
            '取得 Web.config 檔的資料連接設定
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
            '建立 Connection 物件
            Dim conn As New Data.OracleClient.OracleConnection()
            conn.ConnectionString = setting.ConnectionString
            '開啟資料連接
            conn.Open()
            
            '取得28%總額
            Dim strReward As String
            strReward = "select * from Apconfigure where ID=48"
            Dim CmdReward As New Data.OracleClient.OracleCommand(strReward, conn)
            Dim rdReward As Data.OracleClient.OracleDataReader = CmdReward.ExecuteReader()
            If rdReward.HasRows Then
                rdReward.Read()
                DBReward = Trim(rdReward("Value"))
            Else
                DBReward = 0
            End If
            rdReward.Close()

            Dim strUnit As String
            strUnit = "select * from CommonShareReward order by ShareGroupID"
            Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
            Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
            If rdUnit.HasRows Then
                While rdUnit.Read()
                    UnitReward = 0
                    
                    '計算這個群組有幾個單位
                    Dim strUnitReward As String
                    strUnitReward = "select count(*) as cnt from CommonShareReward where ShareGroupID=" & Trim(rdUnit("ShareGroupID"))
                    Dim CmdUnitReward As New Data.OracleClient.OracleCommand(strUnitReward, conn)
                    Dim rdUnitReward As Data.OracleClient.OracleDataReader = CmdUnitReward.ExecuteReader()
                    If rdUnitReward.HasRows Then
                        rdUnitReward.Read()
                        If rdUnitReward("cnt") > 1 Then
                            UnitReward = Decimal.Truncate((DBReward * CDec(rdUnit("SharePercent"))) / rdUnitReward("cnt"))
                        Else
                            UnitReward = Decimal.Truncate(DBReward * CDec(rdUnit("SharePercent")))
                        End If
                        
                    End If
                    RewardTotal = RewardTotal + UnitReward
                    
                    rdUnitReward.Close()
                    Response.Write("<tr>")
                    Response.Write("<td style=""width: 15%""></td>")
                    Response.Write("<td style=""width: 35%"">" & rdUnit("CommonShareUnit") & "</td>")
                    Response.Write("<td style=""width: 25%"" align=""right"">" & Format(UnitReward, "##,##0") & "</td>")
                    Response.Write("<td style=""width: 25%""></td>")
                    Response.Write("")
                    Response.Write("</tr>")
                End While
            End If
            rdUnit.Close()
            
            conn.Close()
        %>
    </table>
    <hr size="3" />
    <table width="680" border="0" cellpadding="3" cellspacing="0" align="center">
        <tr>
        <td style="width: 15%"></td>
        <td style="width: 35%">總計</td>
        <td style="width: 25%" align="right"><%=Format(RewardTotal, "##,##0")%></td>
        <td style="width: 25%"></td>
        </tr>
    </table>
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
	printWindow(true,5.08,5.08,5.08,5.08);
</script>
</html>
