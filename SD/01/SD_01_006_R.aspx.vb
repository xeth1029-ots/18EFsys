Partial Class SD_01_006_R
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
#Region "在這裡放置使用者程式碼以初始化網頁"

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then Call sUtl_Search1()

#End Region
    End Sub

    Sub sUtl_Search1()

        Dim dt As DataTable
        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " SELECT c.OCID, d.OrgName, c.ClassCName " & vbCrLf
        sSql &= " ,COUNT(1) NeedCheck " & vbCrLf
        sSql &= " FROM Stud_EnterType2 a " & vbCrLf
        sSql &= " JOIN Stud_EnterTemp2 b on b.esetid =a.esetid " & vbCrLf
        sSql &= " JOIN Class_ClassInfo c ON c.OCID = a.OCID1 AND a.signUpStatus = 0 " & vbCrLf
        sSql &= " JOIN Auth_Relship r ON r.RID = c.RID " & vbCrLf
        sSql &= " JOIN Org_OrgInfo d ON d.OrgID = r.Orgid " & vbCrLf
        sSql &= " WHERE 1=1 " & vbCrLf
        sSql &= " AND c.PlanID = @PlanID" & vbCrLf
        sSql &= " GROUP BY c.OCID, d.OrgName, c.ClassCName" & vbCrLf
        Dim parms As New Hashtable()
        parms.Add("PlanID", sm.UserInfo.PlanID)
        dt = DbAccess.GetDataTable(sSql, objconn, parms)

        Me.PageControler1.Visible = False
        Me.print.Visible = False
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料")
            msg.Text = "查無資料!!"
            Exit Sub
        End If

        'PageControler1.SqlPrimaryKeyDataCreate(Sql, "OCID")
        msg.Text = ""
        Me.PageControler1.Visible = True
        Me.print.Visible = True

        Me.DataGrid1.DataSource = dt
        Me.DataGrid1.DataBind()

        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "OCID"
        PageControler1.ControlerLoad()

    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "SD_01_006_R", "PlanID=" & sm.UserInfo.PlanID)
    End Sub
End Class