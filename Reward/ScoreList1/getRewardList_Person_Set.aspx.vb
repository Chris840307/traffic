
Partial Class ScoreList_getRewardList_Person_Set
    Inherits System.Web.UI.Page
    Public strUnit As String = "select * from UnitInfo order by UnitID"
    Public checkFlag As Char = ""
    Public ErrorCode As String = ""
    Public UnitFlag As String = ""
    Public sys_City As String = ""

    Protected Sub MultiView1_OnLoad(ByVal sender As Object, ByVal e As System.EventArgs)
        If RadioButtonList1.SelectedValue = "0" Then
            MultiView1.ActiveViewIndex = "0"
        Else
            MultiView1.ActiveViewIndex = "1"
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
        'LEVEL1能看全部單位  LEVEL2能看所屬單位 LEVEL3只能看自己
        If Trim(UserCookie.Values("UnitLevelID")) = 1 Then
            If Trim(Request("AnalyzeType")) = "0" Then
                strUnit = "select UnitID,UnitName from UnitInfo order by UnitID"
            ElseIf Trim(Request("AnalyzeType")) = "1" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & AnalyzeUnitID & "'"
                'Else
                '   strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & AnalyzeUnitID & "'"
            End If
        ElseIf Trim(UserCookie.Values("UnitLevelID")) = 2 Then
            If Trim(Request("AnalyzeType")) = "0" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' or UnitTypeID='" & Trim(UserCookie.Values("UnitID")) & "' order by UnitID"
            ElseIf Trim(Request("AnalyzeType")) = "1" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & AnalyzeUnitID & "' and UnitID='" & Trim(UserCookie.Values("UnitID")) & "'"
            Else
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' or UnitTypeID='" & Trim(UserCookie.Values("UnitID")) & "' order by UnitID"
            End If
        Else
            If Trim(Request("AnalyzeType")) = "0" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' order by UnitID"
            ElseIf Trim(Request("AnalyzeType")) = "1" Then
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & AnalyzeUnitID & "' and UnitID='" & Trim(UserCookie.Values("UnitID")) & "'"
            Else
                strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "'"
            End If
        End If

        UnitFlag = RadioButtonList1.SelectedValue
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
