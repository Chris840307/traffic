<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%
    Dim fMnoth = Month(Now)
    If fMnoth < 10 Then fMnoth = "0" & fMnoth
    Dim fDay = Day(Now)
    If fDay < 10 Then fDay = "0" & fDay
    Dim fname = Year(Now) & fMnoth & fDay & "_LawScore.xls"
    Response.AddHeader("Content-Disposition", "filename=" & fname)
    Response.ContentType = "application/x-msexcel; charset=MS950"
%>
<script runat="server">
    Public sys_City As String
    
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
        
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        sys_City = ""
        Dim strCity = "select Value from ApConfigure where ID=31"
        Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
        Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
        If rdCity.HasRows Then
            rdCity.Read()
            sys_City = Trim(rdCity("Value"))

        End If
        rdCity.Close()

        conn.Close()
        
    End Sub
</script>
<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>法條配分列表</title>
</head>
<body>
    <table width="100%" border="1" cellpadding="3" cellspacing="0" align="center">
        <tr>
            <td style="width:70px; text-align: center;">法條代碼</td>
            <td style="width:270px; text-align: center;">法條名稱</td>
            <td style="width:65px; text-align: center;">車種</td>
            <td style="width:40px; text-align: center;">版本</td>
            <td style="width:70px; text-align: center;">配分標準</td>
            <td style="width:40px; text-align: center;">攔停配分</td>
            <td style="width:40px; text-align: center;">逕舉配分</td>
<%
    If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
 %>            
            <td style="width:40px; text-align: center;">拖吊配分</td>
<%  
    End If
 %>            
            <td style="width:40px; text-align: center;">A1配分</td>
            <td style="width:40px; text-align: center;">A2配分</td>
            <td style="width:40px; text-align: center;">A3配分</td>
        </tr>
<%  
    Dim strSqlPlua As String
    strSqlPlua = ""
    If Trim(Request("sLawID")) <> "" Then
        Dim ArrayLawID = Split(Trim(Request("sLawID")), ",")
        Dim LawCnt As Integer
        strSqlPlua = strSqlPlua & " ("
        For LawCnt = 0 To UBound(ArrayLawID)
            If LawCnt = 0 Then
                strSqlPlua = strSqlPlua & "b.LawItem like '" & ArrayLawID(LawCnt) & "%'"
            Else
                strSqlPlua = strSqlPlua & " or b.LawItem like '" & ArrayLawID(LawCnt) & "%'"
            End If
        Next
        strSqlPlua = strSqlPlua & ")"
    End If
    If Trim(Request("sLawVer")) <> "" Then
        If strSqlPlua = "" Then
            strSqlPlua = strSqlPlua & " b.LawVersion='" & Trim(Request("sLawVer")) & "'"
        Else
            strSqlPlua = strSqlPlua & " and b.LawVersion='" & Trim(Request("sLawVer")) & "'"
        End If
    End If
    If Trim(Request("sCountyOrNpa")) <> "n" Then
        If strSqlPlua = "" Then
            strSqlPlua = strSqlPlua & " b.CountyOrNpa='" & Trim(Request("sCountyOrNpa")) & "'"
        Else
            strSqlPlua = strSqlPlua & " and b.CountyOrNpa='" & Trim(Request("sCountyOrNpa")) & "'"
        End If
    End If
    
    '取得 Web.config 檔的資料連接設定
    Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
    '建立 Connection 物件
    Dim conn As New Data.OracleClient.OracleConnection()
    conn.ConnectionString = setting.ConnectionString
    '開啟資料連接
    conn.Open()
    Dim LawItem, IllegalRule, CarSimpleID, LawVersion, CountyOrNpa, BillType1Score, BillType2Score, Other1Score, A1Score, A2Score, A3Score As String
    Dim strSQL = "select b.* from LawScore b where " & strSqlPlua & " order by b.LawItem,b.CountyOrNpa"
    Dim CmdLaw As New Data.OracleClient.OracleCommand(strSQL, conn)
    Dim rdLaw As Data.OracleClient.OracleDataReader = CmdLaw.ExecuteReader()
    If rdLaw.HasRows Then
        While rdLaw.Read()
            Response.Write("<tr style=""height:45px"">")
            If IsDBNull(rdLaw("LawItem")) Then  '法條代碼
                LawItem = "&nbsp;"
            Else
                LawItem = Trim(rdLaw("LawItem"))
            End If
            Response.Write("<td style=""text-align: left;"">" & LawItem & "</td>")
            
            Dim strRule = "select IllegalRule from Law where ItemID='" & trim(rdLaw("LawItem")) & "' and Version='" & trim(rdLaw("LawVersion")) & "'"
            Dim CmdRule As New Data.OracleClient.OracleCommand(strRule, conn)
            Dim rdRule As Data.OracleClient.OracleDataReader = CmdRule.ExecuteReader()
            If rdRule.HasRows Then
                rdRule.Read()
                If rdRule("IllegalRule") Is DBNull.Value Then
                    IllegalRule = ""
                Else
                    IllegalRule = Trim(rdRule("IllegalRule"))
                End If
            End If
            rdRule.close()
                     
            Response.Write("<td style=""text-align: left;"">" & IllegalRule & "</td>")
            
            If IsDBNull(rdLaw("CarSimpleID")) Then  '車種
                CarSimpleID = "&nbsp;"
            Else
                If Trim(rdLaw("CarSimpleID")) = "1" Then
                    CarSimpleID = Trim(rdLaw("CarSimpleID")) & "自用汽車"
                ElseIf Trim(rdLaw("CarSimpleID")) = "2" Then
                    CarSimpleID = Trim(rdLaw("CarSimpleID")) & "營業車"
                ElseIf Trim(rdLaw("CarSimpleID")) = "3" Then
                    CarSimpleID = Trim(rdLaw("CarSimpleID")) & "機車"
                ElseIf Trim(rdLaw("CarSimpleID")) = "4" Then
                    CarSimpleID = Trim(rdLaw("CarSimpleID")) & "汽車"
                ElseIf Trim(rdLaw("CarSimpleID")) = "5" Then
                    CarSimpleID = Trim(rdLaw("CarSimpleID")) & "小型車"
                ElseIf Trim(rdLaw("CarSimpleID")) = "6" Then
                    CarSimpleID = Trim(rdLaw("CarSimpleID")) & "大型車"
                ElseIf Trim(rdLaw("CarSimpleID")) = "7" Then
                    CarSimpleID = Trim(rdLaw("CarSimpleID")) & "大客"
                ElseIf Trim(rdLaw("CarSimpleID")) = "8" Then
                    CarSimpleID = Trim(rdLaw("CarSimpleID")) & "營大客"
                Else
                    CarSimpleID = "&nbsp;"
                End If
            End If
            Response.Write("<td style=""text-align: left;"">" & CarSimpleID & "</td>")
            
            If IsDBNull(rdLaw("LawVersion")) Then  '版本
                LawVersion = "&nbsp;"
            Else
                LawVersion = Trim(rdLaw("LawVersion"))
            End If
            Response.Write("<td style=""text-align: left;"">" & LawVersion & "</td>")
            
            If IsDBNull(rdLaw("CountyOrNpa")) Then  '配分標準
                CountyOrNpa = "&nbsp;"
            Else
                If Trim(rdLaw("CountyOrNpa")) = "1" Then
                    CountyOrNpa = "績效"
                Else
                    CountyOrNpa = "獎勵金"
                End If
            End If
            Response.Write("<td style=""text-align: left;"">" & CountyOrNpa & "</td>")
            
            If IsDBNull(rdLaw("BillType1Score")) Then  '攔停配分
                BillType1Score = "&nbsp;"
            Else
                BillType1Score = Trim(rdLaw("BillType1Score"))
            End If
            Response.Write("<td style=""text-align: left;"">" & BillType1Score & "</td>")
            
            If IsDBNull(rdLaw("BillType2Score")) Then  '逕舉配分
                BillType2Score = "&nbsp;"
            Else
                BillType2Score = Trim(rdLaw("BillType2Score"))
            End If
            Response.Write("<td style=""text-align: left;"">" & BillType2Score & "</td>")
            
            If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
                If IsDBNull(rdLaw("Other1")) Then  '拖吊配分
                    Other1Score = "&nbsp;"
                Else
                    Other1Score = Trim(rdLaw("Other1"))
                End If
                Response.Write("<td style=""text-align: left;"">" & Other1Score & "</td>")
                
            End If
            
            If IsDBNull(rdLaw("A1Score")) Then  'A1配分
                A1Score = "&nbsp;"
            Else
                A1Score = Trim(rdLaw("A1Score"))
            End If
            Response.Write("<td style=""text-align: left;"">" & A1Score & "</td>")
            
            If IsDBNull(rdLaw("A2Score")) Then  'A2配分
                A2Score = "&nbsp;"
            Else
                A2Score = Trim(rdLaw("A2Score"))
            End If
            Response.Write("<td style=""text-align: left;"">" & A2Score & "</td>")
            
            If IsDBNull(rdLaw("A3Score")) Then  'A3配分
                A3Score = "&nbsp;"
            Else
                A3Score = Trim(rdLaw("A3Score"))
            End If
            Response.Write("<td style=""text-align: left;"">" & A3Score & "</td>")
            Response.Write("</tr>")
        End While
    End If
    rdLaw.Close()
    
    conn.Close()
%>
    </table>
</body>
</html>
