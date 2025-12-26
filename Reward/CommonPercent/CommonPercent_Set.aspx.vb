
Partial Class ScoreList_CommonPercent_Set
    Inherits System.Web.UI.Page
    Public strDisable As String

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        Dim strJavaScript As String
        strJavaScript = ""
        If TbUnit.Text = "" Then
            strJavaScript = vbCrLf & "<script language='javascript' type='text/javascript'>"
            strJavaScript += vbCrLf & "alert(""請輸入分配人員/單位!!"");"
            strJavaScript += vbCrLf & "</script>"
            Me.literal1.Text = strJavaScript
        ElseIf TbPercent.Text = "" Then
            strJavaScript = vbCrLf & "<script language='javascript' type='text/javascript'>"
            strJavaScript += vbCrLf & "alert(""請輸入分配比例!!"");"
            strJavaScript += vbCrLf & "</script>"
            Me.literal1.Text = strJavaScript
        Else
            Dim strSQL1, strSQLSn, UnitID As String
            Dim i, GroupID As Integer
            Dim ArrayUnit = Split(TbUnit.Text, ",")
            Me.literal1.Text = ""
            '群組編號
            GroupID = DropDownList1.SelectedValue
            UnitID = DropDownList2.SelectedValue


            For i = 0 To UBound(ArrayUnit)

                '流水號
                Dim ShareSN As Integer = 0
                Dim strSn = "select max(SN) as MaxSN from CommonShareReward"
                Dim CmdSN As New Data.OracleClient.OracleCommand(strSn, conn)
                Dim rdSN As Data.OracleClient.OracleDataReader = CmdSN.ExecuteReader()
                If rdSN.HasRows Then
                    While rdSN.Read()
                        If rdSN("MaxSN") Is DBNull.Value Then
                            ShareSN = 1
                        Else
                            ShareSN = Int(rdSN("MaxSN")) + 1
                        End If
                    End While
                End If


                strSQL1 = "insert into CommonShareReward(SN,CommonShareUnit,ShareGroupID,SharePercent,UnitID)"
                strSQL1 = strSQL1 & " values(" & ShareSN & ",'" & ArrayUnit(i) & "'," & GroupID & "," & TbPercent.Text / 100 & ",'" & UnitID & "')"

                'Response.Write(strSQL1)
                Dim cmdUpd As New Data.OracleClient.OracleCommand()
                cmdUpd.CommandText = strSQL1
                cmdUpd.Connection = conn
                cmdUpd.ExecuteNonQuery()
            Next

        End If

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

    'Sub UpdateGroupPercent(ByVal GroupID As Object)
    '    '取得 Web.config 檔的資料連接設定
    '    Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
    '    '建立 Connection 物件
    '    Dim conn As New Data.OracleClient.OracleConnection()
    '    conn.ConnectionString = setting.ConnectionString
    '    '開啟資料連接
    '    conn.Open()

    '    Dim strSql As String
    '    Dim PercentValue As Integer = 0
    '    If GroupID = 1 Then
    '        PercentValue = TextBox1.Text
    '    ElseIf GroupID = 2 Then
    '        PercentValue = TextBox2.Text
    '    ElseIf GroupID = 3 Then
    '        PercentValue = TextBox3.Text
    '    ElseIf GroupID = 4 Then
    '        PercentValue = TextBox4.Text
    '    End If
    '    strSql = "update CommonShareReward set SharePercent=" & (PercentValue / 100) & " where commonShareUnit='" & GroupID & "' and ShareGroupID=0"
    '    Dim cmdUpd As New Data.OracleClient.OracleCommand()
    '    cmdUpd.CommandText = strSql
    '    cmdUpd.Connection = conn
    '    cmdUpd.ExecuteNonQuery()

    '    conn.Close()

    '    Dim strJavaScript As String
    '    strJavaScript = vbCrLf & "<script language='javascript' type='text/javascript'>"
    '    strJavaScript += vbCrLf & "alert(""修改完成!!"");"
    '    strJavaScript += vbCrLf & "</script>"
    '    Me.literal1.Text = strJavaScript
    'End Sub

 
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoginCheck()
        Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")

        'LEVEL1能看全部單位  LEVEL2只能看自己
        If Trim(UserCookie.Values("UnitLevelID")) = 1 Then
            SqlDataSource1.SelectCommand = "select UnitID,UnitName from UnitInfo where ShowOrder in (0,1) order by UnitID"
            strDisable = ""
        Else
            SqlDataSource1.SelectCommand = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' and ShowOrder in (0,1) order by UnitID"
            If DropDownList1.SelectedValue = "1" Or DropDownList1.SelectedValue = "4" Then
                Button1.Enabled = False
            Else
                Button1.Enabled = True
            End If
            strDisable = "disabled"
        End If

    End Sub
End Class
