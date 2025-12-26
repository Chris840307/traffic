
Partial Class LawScore_LawScoreSet
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        Dim sys_City As String
        sys_City = ""
        Dim strCity = "select Value from ApConfigure where ID=31"
        Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
        Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
        If rdCity.HasRows Then
            rdCity.Read()
            sys_City = Trim(rdCity("Value"))

        End If

        Dim cmdUpd As New Data.OracleClient.OracleCommand()
        Dim strSql As String
        If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
            strSql = "Update LawScore set BillType1Score='" & Trim(Request("Type1")) & "'"
            strSql = strSql & ",BillType2Score='" & Trim(Request("Type2")) & "',A1Score='" & Trim(Request("A1")) & "'"
            strSql = strSql & ",A2Score='" & Trim(Request("A2")) & "',A3Score='" & Trim(Request("A3")) & "'"
            strSql = strSql & ",Other1='" & Trim(Request("Other1")) & "'"
            strSql = strSql & " where LawItem='" & Trim(Request("LawID")) & "' and LawVersion='" & Trim(Request("LawVer")) & "'"
            strSql = strSql & " and CountyOrNpa=" & Trim(Request("CorN")) & " and CarSimpleID='" & Trim(Request("CarSimple")) & "'"
        Else
            strSql = "Update LawScore set BillType1Score='" & Trim(Request("Type1")) & "'"
            strSql = strSql & ",BillType2Score='" & Trim(Request("Type2")) & "',A1Score='" & Trim(Request("A1")) & "'"
            strSql = strSql & ",A2Score='" & Trim(Request("A2")) & "',A3Score='" & Trim(Request("A3")) & "'"
            strSql = strSql & " where LawItem='" & Trim(Request("LawID")) & "' and LawVersion='" & Trim(Request("LawVer")) & "'"
            strSql = strSql & " and CountyOrNpa=" & Trim(Request("CorN")) & " and CarSimpleID='" & Trim(Request("CarSimple")) & "'"
        End If

        cmdUpd.CommandText = strSql
        cmdUpd.Connection = conn
        cmdUpd.ExecuteNonQuery()
        Response.Write("儲存成功!")
        conn.Close()
    End Sub
End Class
