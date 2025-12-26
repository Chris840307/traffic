
Partial Class LawScore_LawScore
    Inherits System.Web.UI.Page
    Public str1 As String
    Dim strJavaScript As String
    Public sys_City As String



    '抓現行版本編號
    Protected Sub LawVer_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LawVer.Init
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        Dim strLawVer = "select Value from ApConfigure where ID=3"
        Dim CmdLawVer As New Data.OracleClient.OracleCommand(strLawVer, conn)
        Dim rdLawVer As Data.OracleClient.OracleDataReader = CmdLawVer.ExecuteReader()
        If rdLawVer.HasRows Then
            rdLawVer.Read()
            LawVer.Text = Trim(rdLawVer("Value"))

        End If
        rdLawVer.Close()
        conn.Close()
        btBack.Enabled = False
        btNext.Enabled = False
        'enable excel鍵
        strJavaScript = vbCrLf & "<script language='javascript' type='text/javascript'>"
        strJavaScript += vbCrLf & "form1.btExcel.disabled=true"
        strJavaScript += vbCrLf & "</script>"
        Me.Literal1.Text = strJavaScript
    End Sub
    '查詢
    Sub LawView(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strPlus As String = ""
        If Trim(Request("LawID")) <> "" Then
            Dim ArrayLawID = Split(Trim(Request("LawID")), ",")
            Dim LawCnt As Integer
            strPlus = strPlus & " ("
            For LawCnt = 0 To UBound(ArrayLawID)
                If LawCnt = 0 Then
                    strPlus = strPlus & "b.LawItem like '" & ArrayLawID(LawCnt) & "%'"
                Else
                    strPlus = strPlus & " or b.LawItem like '" & ArrayLawID(LawCnt) & "%'"
                End If
            Next
            strPlus = strPlus & ")"
        End If
        If Trim(Request("LawVer")) <> "" Then
            If strPlus = "" Then
                strPlus = strPlus & " b.LawVersion='" & Trim(Request("LawVer")) & "'"
            Else
                strPlus = strPlus & " and b.LawVersion='" & Trim(Request("LawVer")) & "'"
            End If
        End If
        If Trim(Request("sCountyOrNpa")) <> "n" Then
            If strPlus = "" Then
                strPlus = strPlus & " b.CountyOrNpa='" & Trim(Request("sCountyOrNpa")) & "'"
            Else
                strPlus = strPlus & " and b.CountyOrNpa='" & Trim(Request("sCountyOrNpa")) & "'"
            End If
        End If

        LbPageNum.Text = 0

        Dim PageCount As Integer = Trim(Request("PageCount"))
        Dim PageNum As Integer = 0
        Dim RowStart As Integer = (PageNum * PageCount) + 1
        Dim RowEnd As Integer = ((PageNum + 1) * PageCount)
        Dim strLaw = "select * from (select rownum r,c.* from (select b.* from LawScore b where " & strPlus & " order by LawItem,CountyOrNpa) c) where r between " & RowStart & " and " & RowEnd
        RowView(strLaw, strPlus)
        btBack.Enabled = False
        'Response.Write(strLaw)
    End Sub
    '上一頁
    Sub Page_Back(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strPlus As String = ""
        If Trim(Request("LawID")) <> "" Then
            Dim ArrayLawID = Split(Trim(Request("LawID")), ",")
            Dim LawCnt As Integer
            strPlus = strPlus & " ("
            For LawCnt = 0 To UBound(ArrayLawID)
                If LawCnt = 0 Then
                    strPlus = strPlus & "b.LawItem like '" & ArrayLawID(LawCnt) & "%'"
                Else
                    strPlus = strPlus & " or b.LawItem like '" & ArrayLawID(LawCnt) & "%'"
                End If
            Next
            strPlus = strPlus & ")"
        End If
        If Trim(Request("LawVer")) <> "" Then
            If strPlus = "" Then
                strPlus = strPlus & " b.LawVersion='" & Trim(Request("LawVer")) & "'"
            Else
                strPlus = strPlus & " and b.LawVersion='" & Trim(Request("LawVer")) & "'"
            End If
        End If
        If Trim(Request("sCountyOrNpa")) <> "n" Then
            If strPlus = "" Then
                strPlus = strPlus & " b.CountyOrNpa='" & Trim(Request("sCountyOrNpa")) & "'"
            Else
                strPlus = strPlus & " and b.CountyOrNpa='" & Trim(Request("sCountyOrNpa")) & "'"
            End If
        End If

        Dim PageCount As Integer = Trim(Request("PageCount"))
        Dim PageNum As Integer = CInt(LbPageNum.Text) - 1
        Dim RowStart As Integer = (PageNum * PageCount) + 1
        Dim RowEnd As Integer = ((PageNum + 1) * PageCount)
        Dim strLaw = "select * from (select rownum r,c.* from (select b.* from LawScore b where " & strPlus & " order by LawItem,CountyOrNpa) c) where r between " & RowStart & " and " & RowEnd
        LbPageNum.Text = LbPageNum.Text - 1
        RowView(strLaw, strPlus)

        If Trim(LbPageNum.Text) = 0 Then
            btBack.Enabled = False
        End If
        ' Response.Write(strLaw)
        
    End Sub
    '下一頁
    Sub Page_Next(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strPlus As String = ""
        If Trim(Request("LawID")) <> "" Then
            Dim ArrayLawID = Split(Trim(Request("LawID")), ",")
            Dim LawCnt As Integer
            strPlus = strPlus & " ("
            For LawCnt = 0 To UBound(ArrayLawID)
                If LawCnt = 0 Then
                    strPlus = strPlus & "b.LawItem like '" & ArrayLawID(LawCnt) & "%'"
                Else
                    strPlus = strPlus & " or b.LawItem like '" & ArrayLawID(LawCnt) & "%'"
                End If
            Next
            strPlus = strPlus & ")"
        End If
        If Trim(Request("LawVer")) <> "" Then
            If strPlus = "" Then
                strPlus = strPlus & " b.LawVersion='" & Trim(Request("LawVer")) & "'"
            Else
                strPlus = strPlus & " and b.LawVersion='" & Trim(Request("LawVer")) & "'"
            End If
        End If
        If Trim(Request("sCountyOrNpa")) <> "n" Then
            If strPlus = "" Then
                strPlus = strPlus & " b.CountyOrNpa='" & Trim(Request("sCountyOrNpa")) & "'"
            Else
                strPlus = strPlus & " and b.CountyOrNpa='" & Trim(Request("sCountyOrNpa")) & "'"
            End If
        End If

        Dim PageCount As Integer = Trim(Request("PageCount"))
        Dim PageNum As Integer = CInt(LbPageNum.Text) + 1
        Dim RowStart As Integer = (PageNum * PageCount) + 1
        Dim RowEnd As Integer = ((PageNum + 1) * PageCount)
        Dim strLaw = "select * from (select rownum r,c.* from (select b.* from LawScore b where " & strPlus & " order by LawItem,CountyOrNpa) c) where r between " & RowStart & " and " & RowEnd
        LbPageNum.Text = LbPageNum.Text + 1
        RowView(strLaw, strPlus)

        btBack.Enabled = True
        'Response.Write(strLaw)
    End Sub
    '法條列表
    Sub RowView(ByVal SqlStr, ByVal SqlPlusStr)
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        Dim BillType1, BillType2, A1, A2, A3, OtherType1, strCountyOrNpa As String
        Dim TagOrder As Integer = 0

        '列出法條
        Dim strLaw = SqlStr
        'response.write(strLaw)
        ' response.end()
        Dim CmdLaw As New Data.OracleClient.OracleCommand(strLaw, conn)
        Dim rdLaw As Data.OracleClient.OracleDataReader = CmdLaw.ExecuteReader()
        If rdLaw.HasRows Then
            While rdLaw.Read()
                TagOrder = TagOrder + 1
                str1 = str1 & "<tr>"
                str1 = str1 & "<td style='background-color:#FFFFFF ; height:40px'>" & Trim(rdLaw("LawItem")) & "</td>"
                str1 = str1 & "<td style='background-color:#FFFFFF'>"
                Dim strRule = "select IllegalRule from Law where ItemID='" & trim(rdLaw("LawItem")) & "' and Version='" & trim(rdLaw("LawVersion")) & "'"
                Dim CmdRule As New Data.OracleClient.OracleCommand(strRule, conn)
                Dim rdRule As Data.OracleClient.OracleDataReader = CmdRule.ExecuteReader()
                If rdRule.HasRows Then
                    rdRule.Read()
                    If rdRule("IllegalRule") Is DBNull.Value Then
                        str1 = str1 & ""
                    Else
                        str1 = str1 & Trim(rdRule("IllegalRule"))
                    End If

                End If
                rdRule.close()
                str1 = str1 & "</td>"

                str1 = str1 & "<td style='background-color:#FFFFFF'>"
                If IsDBNull(rdLaw("CarSimpleID")) Then
                    str1 = str1 & "&nbsp;"
                Else
                    If Trim(rdLaw("CarSimpleID")) = "1" Then
                        str1 = str1 & Trim(rdLaw("CarSimpleID")) & "自用汽車"
                    ElseIf Trim(rdLaw("CarSimpleID")) = "2" Then
                        str1 = str1 & Trim(rdLaw("CarSimpleID")) & "營業車"
                    ElseIf Trim(rdLaw("CarSimpleID")) = "3" Then
                        str1 = str1 & Trim(rdLaw("CarSimpleID")) & "機車"
                    ElseIf Trim(rdLaw("CarSimpleID")) = "4" Then
                        str1 = str1 & Trim(rdLaw("CarSimpleID")) & "汽車"
                    ElseIf Trim(rdLaw("CarSimpleID")) = "5" Then
                        str1 = str1 & Trim(rdLaw("CarSimpleID")) & "小型車"
                    ElseIf Trim(rdLaw("CarSimpleID")) = "6" Then
                        str1 = str1 & Trim(rdLaw("CarSimpleID")) & "大型車"
                    ElseIf Trim(rdLaw("CarSimpleID")) = "7" Then
                        str1 = str1 & Trim(rdLaw("CarSimpleID")) & "大客"
                    ElseIf Trim(rdLaw("CarSimpleID")) = "8" Then
                        str1 = str1 & Trim(rdLaw("CarSimpleID")) & "營大客"
                    End If
                End If
                str1 = str1 & "</td>"

                str1 = str1 & "<td style='background-color:#FFFFFF'>"
                If IsDBNull(rdLaw("LawVersion")) Then
                    str1 = str1 & "&nbsp;"
                Else
                    str1 = str1 & Trim(rdLaw("LawVersion"))
                End If
                str1 = str1 & "</td>"

                If IsDBNull(rdLaw("BillType1Score")) Then
                    BillType1 = ""
                Else
                    BillType1 = Trim(rdLaw("BillType1Score"))
                End If
                If IsDBNull(rdLaw("BillType2Score")) Then
                    BillType2 = ""
                Else
                    BillType2 = Trim(rdLaw("BillType2Score"))
                End If
                If IsDBNull(rdLaw("Other1")) Then
                    OtherType1 = ""
                Else
                    OtherType1 = Trim(rdLaw("Other1"))
                End If
                If IsDBNull(rdLaw("A1Score")) Then
                    A1 = ""
                Else
                    A1 = Trim(rdLaw("A1Score"))
                End If
                If IsDBNull(rdLaw("A2Score")) Then
                    A2 = ""
                Else
                    A2 = Trim(rdLaw("A2Score"))
                End If
                If IsDBNull(rdLaw("A3Score")) Then
                    A3 = ""
                Else
                    A3 = Trim(rdLaw("A3Score"))
                End If
                If IsDBNull(rdLaw("CountyOrNpa")) Then
                    strCountyOrNpa = ""
                Else
                    If Trim(rdLaw("CountyOrNpa")) = "0" Then
                        strCountyOrNpa = "獎勵金"
                    Else
                        strCountyOrNpa = "績效"
                    End If
                End If
                str1 = str1 & "<td style='background-color:#FFFFFF' align='center'>"
                str1 = str1 & strCountyOrNpa
                str1 = str1 & "</td>"
                str1 = str1 & "<td style='background-color:#FFFFFF' align='center'>"
                str1 = str1 & "<input type='Text' Name='BillType1" & TagOrder & "' value='" & BillType1 & "' style='width: 35px' maxlength='4' >"
                str1 = str1 & "</td>"
                str1 = str1 & "<td style='background-color:#FFFFFF' align='center'>"
                str1 = str1 & "<input type='Text' Name='BillType2" & TagOrder & "' value='" & BillType2 & "' style='width: 35px' maxlength='4' >"
                str1 = str1 & "</td>"
                If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
                    str1 = str1 & "<td style='background-color:#FFFFFF' align='center'>"
                    str1 = str1 & "<input type='Text' Name='Other1" & TagOrder & "' value='" & OtherType1 & "' style='width: 35px' maxlength='4' >"
                    str1 = str1 & "</td>"
                End If
                str1 = str1 & "<td style='background-color:#FFFFFF' align='center'>"
                str1 = str1 & "<input type='Text' Name='A1" & TagOrder & "' value='" & A1 & "' style='width: 35px' maxlength='4' >"
                str1 = str1 & "</td>"
                str1 = str1 & "<td style='background-color:#FFFFFF' align='center'>"
                str1 = str1 & "<input type='Text' Name='A2" & TagOrder & "' value='" & A2 & "' style='width: 35px' maxlength='4' >"
                str1 = str1 & "</td>"
                str1 = str1 & "<td style='background-color:#FFFFFF' align='center'>"
                str1 = str1 & "<input type='Text' Name='A3" & TagOrder & "' value='" & A3 & "' style='width: 35px' maxlength='4' >"
                str1 = str1 & "</td>"
                str1 = str1 & "<td style='background-color:#FFFFFF' align='center'>"
                str1 = str1 & "<input type='button' value='修改' onclick='ScoreUpdate(" & Trim(rdLaw("LawItem")) & "," & Trim(rdLaw("LawVersion")) & "," & Trim(rdLaw("CountyOrNpa")) & "," & Trim(rdLaw("CarSimpleID")) & "," & TagOrder & ")'>"
                str1 = str1 & "</td>"
                str1 = str1 & "</tr>"

            End While
            'enable excel鍵
            strJavaScript = vbCrLf & "<script language='javascript' type='text/javascript'>"
            strJavaScript += vbCrLf & "form1.btExcel.disabled=false"
            strJavaScript += vbCrLf & "</script>"
            Me.Literal1.Text = strJavaScript
        Else
            'disable excel鍵
            strJavaScript = vbCrLf & "<script language='javascript' type='text/javascript'>"
            strJavaScript += vbCrLf & "form1.btExcel.disabled=true"
            strJavaScript += vbCrLf & "</script>"
            Me.Literal1.Text = strJavaScript

        End If
        rdLaw.Close()

        '計算總頁數
        Dim PageTotal As Integer
        Dim strLawCnt = "select count(*) as CNT from LawScore b where " & SqlPlusStr
        Dim CmdLawCnt As New Data.OracleClient.OracleCommand(strLawCnt, conn)
        Dim rdLawCnt As Data.OracleClient.OracleDataReader = CmdLawCnt.ExecuteReader()
        If rdLawCnt.HasRows Then
            rdLawCnt.Read()
            PageTotal = Decimal.Ceiling((CInt(rdLawCnt("CNT")) / CInt(Request("PageCount"))))
            LbPageNumTotal.Text = (LbPageNum.Text + 1) & " / " & PageTotal
            If LbPageNum.Text + 1 >= PageTotal Then
                btNext.Enabled = False
            Else
                btNext.Enabled = True
            End If
        End If
        rdLawCnt.Close()
        conn.Close()
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
End Class
