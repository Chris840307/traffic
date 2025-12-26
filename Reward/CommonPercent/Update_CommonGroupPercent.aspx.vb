
Partial Class CommonPercent_Update_CommonGroupPercent
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()

        '更新
        Dim cmdUpd As New Data.OracleClient.OracleCommand()
        Dim strSql = "Update CommonShareReward set SharePercent=" & Trim(Request("SharePercent")) / 100 & " where ShareGroupID=0 and CommonShareUnit='" & Trim(Request("GroupID")) & "'"
        cmdUpd.CommandText = strSql
        cmdUpd.Connection = conn
        cmdUpd.ExecuteNonQuery()

        Response.Write("修改完成!!")

        conn.Close()
    End Sub
End Class
