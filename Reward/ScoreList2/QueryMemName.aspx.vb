
Partial Class ScoreList_QueryMemName
    Inherits System.Web.UI.Page


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()

        Dim strNameValue = "select chName from MemberData where LoginID='" & Trim(Request("MLoginID")) & "' and RecordStateID<>-1 and AccountStateID<>-1"
        Dim CmdNameValue As New Data.OracleClient.OracleCommand(strNameValue, conn)
        Dim rdNameValue As Data.OracleClient.OracleDataReader = CmdNameValue.ExecuteReader()
        If rdNameValue.HasRows Then
            rdNameValue.Read()
            Response.Write(Trim(rdNameValue("chName")))
        End If
        rdNameValue.Close()

        conn.Close()

    End Sub
End Class
