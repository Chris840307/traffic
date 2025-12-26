
Partial Class ScoreList_RewardTotalSet
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()

        '獎勵金總額
        Dim cmdUpd As New Data.OracleClient.OracleCommand()
        Dim strSql = "Update Apconfigure set Value='" & Trim(Request("RewardTotal")) & "' where ID=46"
        cmdUpd.CommandText = strSql
        cmdUpd.Connection = conn
        cmdUpd.ExecuteNonQuery()

        Dim Money28 = Format(Decimal.Round(Request("RewardTotal") * 0.28), "##,##0")
        Dim Money72 = Format(Decimal.Round(Request("RewardTotal") * 0.72), "##,##0")
        '直接人員獎勵金
        Dim aMoney72 = Decimal.Round(Request("RewardTotal") * 0.72)
        Dim cmdUpd2 As New Data.OracleClient.OracleCommand()
        Dim strSql2 = "Update Apconfigure set Value='" & aMoney72 & "' where ID=45"
        cmdUpd2.CommandText = strSql2
        cmdUpd2.Connection = conn
        cmdUpd2.ExecuteNonQuery()

        '共同人員獎勵金
        Dim aMoney28 = Decimal.Round(Request("RewardTotal") * 0.28)
        Dim cmdUpd3 As New Data.OracleClient.OracleCommand()
        Dim strSql3 = "Update Apconfigure set Value='" & aMoney28 & "' where ID=48"
        cmdUpd3.CommandText = strSql3
        cmdUpd3.Connection = conn
        cmdUpd3.ExecuteNonQuery()

        Response.Write(Money28 & "@@" & Money72)
        conn.Close()
    End Sub
End Class
