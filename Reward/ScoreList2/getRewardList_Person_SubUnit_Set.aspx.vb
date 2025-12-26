
Partial Class ScoreList_getRewardList_Person_SubUnit_Set
    Inherits System.Web.UI.Page
    Public strUnit As String = "select * from UnitInfo order by UnitID"
    Public checkFlag As Char = ""
    Public ErrorCode As String = ""
    Public UnitFlag As String = ""


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoginCheck()
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        Dim AnalyzeUnitID As String = ""
        Dim strUID = "select Value from Apconfigure where ID=49"
        Dim CmdUID As New Data.OracleClient.OracleCommand(strUID, conn)
        Dim rsUID As Data.OracleClient.OracleDataReader = CmdUID.ExecuteReader()
        If rsUID.HasRows Then
            rsUID.Read()
            AnalyzeUnitID = Trim(rsUID("Value"))
        End If
        rsUID.Close()
        conn.Close()

        Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")
        'LEVEL1能看全部單位  LEVEL2能看所屬單位 LEVEL3不能看
        If Trim(UserCookie.Values("UnitLevelID")) = 1 Then
            If Trim(Request("AnalyzeType")) = "0" Then
                strUnit = "select UnitID,UnitName from UnitInfo where ShowOrder in (0,1) order by UnitID"
            ElseIf Trim(Request("AnalyzeType")) = "1" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & AnalyzeUnitID & "'"
                'Else
                '   strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & AnalyzeUnitID & "'"
            End If
        ElseIf Trim(UserCookie.Values("UnitLevelID")) = 2 Then
            If Trim(Request("AnalyzeType")) = "0" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' and ShowOrder=1 order by UnitID"
            ElseIf Trim(Request("AnalyzeType")) = "1" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & AnalyzeUnitID & "' and UnitID='" & Trim(UserCookie.Values("UnitID")) & "'"
            Else
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' and ShowOrder=1 order by UnitID"
            End If
        Else
            If Trim(Request("AnalyzeType")) = "0" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' and ShowOrder=1 order by UnitID"
            ElseIf Trim(Request("AnalyzeType")) = "1" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & AnalyzeUnitID & "' and UnitID='" & Trim(UserCookie.Values("UnitID")) & "'"
            Else
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' and ShowOrder=1"
            End If
        End If

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
End Class
