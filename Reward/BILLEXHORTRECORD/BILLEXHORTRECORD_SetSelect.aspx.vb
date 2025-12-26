
Partial Class BILLEXHORTRECORD_SetSelect
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        conn.Open()
        Dim strUNITLEVELID As String = ""
        Dim strDbChk = "select UNITLEVELID from UNITINFO where UNITID='" & Request.Cookies("RewardUser")("UnitID") & "'"

        Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
        Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
        If rdDbChk.HasRows Then
            rdDbChk.Read()
            strUNITLEVELID = Trim(rdDbChk("UNITLEVELID"))
        End If
        rdDbChk.Close()

        If strUNITLEVELID = "1" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO order by showorder"
        ElseIf strUNITLEVELID = "2" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UNITTYPEID='" & Request.Cookies("RewardUser")("UnitID") & "' or unitid='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        ElseIf strUNITLEVELID = "3" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UnitID='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        Else
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UNITTYPEID='" & Request.Cookies("RewardUser")("UnitID") & "' or unitid='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        End If
        conn.Close()
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        If txtDate1.Text = "" Or txtDate2.Text = "" Then
            Response.Write("<script language=""javascript"">")
            Response.Write("alert(""起迄日期需要輸入"");")
            Response.Write("</script>")
            Exit Sub
        End If

        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=" & Server.UrlPathEncode(Format(Now(), "yyyyMMdd") & "_交通違規勸導績效統計表.xls"))
        Response.Charset = "big5"
        Response.ContentType = "application/vnd.ms-excel"

        Response.Write(GenHtmlTable())

        Response.End()
        
    End Sub

    Function GenHtmlTable()
        Dim sRet
        Dim H As String = "height=50"

        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        conn.Open()

        Dim UnitName As String = ""
        Dim strDbChk = "select UnitName from UNITINFO where UNITID='" & Request.Cookies("RewardUser")("UnitID") & "'"
        'Dim strDbChk = "select UnitName from UNITINFO where UNITID='9C00'"
        Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
        Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
        If rdDbChk.HasRows Then
            rdDbChk.Read()
            UnitName = Trim(rdDbChk("UnitName"))
        End If
        rdDbChk.Close()

        Dim chName As String = ""
        Dim strDbChk2 = "select UnitName from UNITINFO where UNITID='" & Request.Cookies("RewardUser")("MemberID") & "'"
        'Dim strDbChk2 = "select chName from MemberData where MemberID=471"
        Dim CmdDbChk2 As New Data.OracleClient.OracleCommand(strDbChk2, conn)
        Dim rdDbChk2 As Data.OracleClient.OracleDataReader = CmdDbChk2.ExecuteReader()
        If rdDbChk2.HasRows Then
            rdDbChk2.Read()
            chName = Trim(rdDbChk2("chName"))
        End If
        rdDbChk2.Close()
        conn.Close()


        sRet = ""
        sRet = "<table border=0>"
        sRet = sRet & "<td colspan=4 align=""center"">" & UnitName & Year(txtDate1.Text) - 1911 & "年" & Month(txtDate1.Text) & "月執行交通違規勸導績效統計表</td>"
        sRet = sRet & "<tr>"
        'sRet = sRet & "<td colspan=4>填報單位:" & UnitName & "</td>"
        sRet = sRet & "<tr>"
        'sRet = sRet & "<td colspan=4>填表人:" & chName & "</td>"
        sRet = sRet & "</table>"

        sRet = sRet & "<table border=1 fontsize=4>"
        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">勸導內容</td>"
        sRet = sRet & "<td>勸導件數</td>"
        sRet = sRet & "<td>截至本月份勸導件數</td>"
        sRet = sRet & "<td>備    考</td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">1.未帶駕照（經查證領有駕照）。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("1,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("1,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">2.未帶行照（經查證領有行照）</td>"
        sRet = sRet & "<td>" & GetSqlCnt("2,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("2,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">3.號牌污穢(責令當場改正)。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("3,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("3,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">4.亂鳴喇叭(當場勸戒)。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("4,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("4,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">5.超載10%以下。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("5,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("5,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">6.酒測值逾0.02毫克以下。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("6,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("6,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">7.大型車右轉未先駛入外側車道。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("7,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("7,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">8.機車附載人員或物品未依規定者。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("8,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("8,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">9.號誌燈變換，車前輪未進入停止線。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("9,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("9,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">10.號誌燈變換，車前輪未進入機車停等區。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("10,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("10,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">11.在道路堆積、置放、設置或拋擲足以妨礙交通之物。</td>"
        sRet = sRet & "<td>" & GetSqlCnt("11,") & " </td>"
        sRet = sRet & "<td>" & GetSqlCntToNow("11,") & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">12.臨時停車及其他。</td>"
        sRet = sRet & "<td>" & GetSqlCnt12() & " </td>"
        sRet = sRet & "<td>" & GetSqlCnt12ToNow() & " </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td " & H & ">總計</td>"
        sRet = sRet & "<td>=sum(B5:B16) </td>"
        sRet = sRet & "<td>=sum(C5:C16) </td>"
        sRet = sRet & "<td></td>"
        sRet = sRet & "</tr>"

        sRet = sRet & "<tr>"
        sRet = sRet & "<td colspan=4>備註：本表自９５年７月起開始統計，並於次月７日前免備文逕送本局（交通隊）彙辦</td>"
        sRet = sRet & "</tr>"


        sRet = sRet & "</table>"
        GenHtmlTable = sRet
    End Function

    Function GetSqlCnt(ByVal sql)

        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        conn.Open()

        Dim strDbChk = "select count(*) as cnt from BILLEXHORTRECORD where IllegalItemList like '%" & sql & "%'"
        strDbChk = strDbChk & " and IllegalDateTime between to_date('" & txtDate1.Text & " 00:00:00','yyyy/MM/dd HH24:MI:SS') and to_date('" & txtDate2.Text & " 23:59:59','yyyy/MM/dd HH24:MI:SS')"

        If cbxUnit.Checked Then
            strDbChk = strDbChk & GetUnitLeveL()
        End If

        Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
        Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()

        If rdDbChk.HasRows Then
            rdDbChk.Read()
            GetSqlCnt = Trim(rdDbChk("cnt"))
        End If

        rdDbChk.Close()
        conn.Close()
    End Function

    Function GetSqlCntToNow(ByVal sql)

        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        conn.Open()

        Dim strDbChk = "select count(*) as cnt from BILLEXHORTRECORD where IllegalItemList like '%" & sql & "%'"

        If cbxUnit.Checked Then
            strDbChk = strDbChk & GetUnitLeveL()
        End If

        Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
        Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
        If rdDbChk.HasRows Then
            rdDbChk.Read()
            GetSqlCntToNow = Trim(rdDbChk("cnt"))
        End If
        rdDbChk.Close()
        conn.Close()
    End Function

    Function GetSqlCnt12()

        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        conn.Open()

        Dim strDbChk = "select count(*) as cnt from BILLEXHORTRECORD where otherillegalitem is not null"
        strDbChk = strDbChk & " and IllegalDateTime between to_date('" & txtDate1.Text & " 00:00:00','yyyy/MM/dd HH24:MI:SS') and to_date('" & txtDate2.Text & " 23:59:59','yyyy/MM/dd HH24:MI:SS')"

        If cbxUnit.Checked Then
            strDbChk = strDbChk & GetUnitLeveL()
        End If

        Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
        Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
        If rdDbChk.HasRows Then
            rdDbChk.Read()
            GetSqlCnt12 = Trim(rdDbChk("cnt"))
        End If
        rdDbChk.Close()
        conn.Close()
    End Function

    Function GetSqlCnt12ToNow()

        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        conn.Open()

        Dim strDbChk = "select count(*) as cnt from BILLEXHORTRECORD where otherillegalitem is not null"

        If cbxUnit.Checked Then
            strDbChk = strDbChk & GetUnitLeveL()
        End If

        Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
        Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
        If rdDbChk.HasRows Then
            rdDbChk.Read()
            GetSqlCnt12ToNow = Trim(rdDbChk("cnt"))
        End If
        rdDbChk.Close()
        conn.Close()
    End Function

    Protected Sub cbxUnit_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxUnit.CheckedChanged
        If cbxUnit.Checked Then
            DDLUnit.Enabled = True
        Else
            DDLUnit.Enabled = False
        End If
    End Sub
    Function GetUnitLeveL()
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        conn.Open()
        Dim strUNITLEVELID As String = ""
        Dim strDbChk = "select UNITLEVELID from UNITINFO where UNITID='" & DDLUnit.SelectedValue & "'"
        Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
        Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
        If rdDbChk.HasRows Then
            rdDbChk.Read()
            strUNITLEVELID = Trim(rdDbChk("UNITLEVELID"))
        End If
        rdDbChk.Close()

        If strUNITLEVELID = "1" Then
            GetUnitLeveL = ""
        ElseIf strUNITLEVELID = "2" Then
            GetUnitLeveL = " and EXHORTUNITID in (SELECT UNITID FROM UNITINFO where UNITTYPEID='" & DDLUnit.SelectedValue & "' or unitid='" & DDLUnit.SelectedValue & "')"
        ElseIf strUNITLEVELID = "3" Then
            GetUnitLeveL = " and EXHORTUNITID ='" & DDLUnit.SelectedValue & "'"
        End If
        conn.Close()
    End Function
End Class
