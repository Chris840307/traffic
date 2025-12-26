Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoginCheck()
        Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")
        Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        '--------最上面使用者資料區塊---------------
        Dim strUnit = "select UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "'"
        Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
        Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
        If rdUnit.HasRows Then
            rdUnit.Read()
            LbUserData.Text = Trim(rdUnit("UnitName")) & "&nbsp; &nbsp; &nbsp; &nbsp;"
        End If
        rdUnit.Close()
        LbUserData.Text = LbUserData.Text & HttpUtility.UrlDecode(UserCookie.Values("ChName")) & "&nbsp; &nbsp; &nbsp; &nbsp;"

        Dim strFuncName = "select Content from Code where TypeID=10 and ID=" & Trim(UserCookie.Values("GroupRoleID"))
        Dim CmdFName As New Data.OracleClient.OracleCommand(strFuncName, conn)
        Dim rdFName As Data.OracleClient.OracleDataReader = CmdFName.ExecuteReader()
        If rdFName.HasRows Then
            rdFName.Read()
            LbUserData.Text = LbUserData.Text & Trim(rdFName("Content"))
        End If
        rdFName.Close()


        conn.Close()
    End Sub
    Sub UserLogout(ByVal sender As Object, ByVal e As System.EventArgs)
        '------------清除cookie-------------
        '使用者資料
        Response.Cookies("RewardUser")("CreditID") = ""
        Response.Cookies("RewardUser")("Password") = ""
        Response.Cookies("RewardUser")("CreditID") = ""
        Response.Cookies("RewardUser")("LoginID") = ""
        Response.Cookies("RewardUser")("UnitID") = ""
        Response.Cookies("RewardUser")("ChName") = ""
        Response.Cookies("RewardUser")("GroupRoleID") = ""
        Response.Cookies("RewardUser")("UnitLevelID") = ""
        Response.Cookies("RewardUser")("ManagerPower") = ""
        Response.Cookies("RewardUser")("DCIwindowName") = ""
        Response.Cookies("RewardUser")("DoubleCheck") = ""
        '使用者權限
        Response.Cookies("UserFunction")("FuncID") = ""
        '--------------------------------------
        Response.Redirect("Logout.aspx")
    End Sub
    Sub LoginCheck()
        If (Request.Cookies("UserFunction") IsNot Nothing) Then
            Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
            If Trim(FuncCookie.Values("FuncID")) = "" Then
                Response.Redirect("Login.aspx?ErrMsg=1")

            End If
        Else
            Response.Redirect("Login.aspx?ErrMsg=1")
        End If
    End Sub
End Class
