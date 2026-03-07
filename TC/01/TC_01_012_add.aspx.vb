Partial Class TC_01_012_add
    Inherits System.Web.UI.Page

    'Dim objconn As OracleConnection
    'Dim objreader As OracleDataReader
    Dim objconn As OracleConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        Me.bt_addrow.Attributes("onclick") = "return bt_addrow_check();"

        If Not Page.IsPostBack Then
            yearlist_add = TIMS.GetSyear(yearlist_add, 2005)
            ClassChar = TIMS.Get_ClassChar(ClassChar)

            If Request("ProcessType") = "Update" Then
                Me.yearlist_add.SelectedValue = Request("PLANYEAR")
                Me.yearlist_add.Enabled = False
                'RIDValue.Value = Request("orgid")
                OrgID.Value = Request("ORGID")
                RIDValue.Value = Request("RIDValue")
                'If OrgID.Value <> "" Then RIDValue.Value = TIMS.Get_RIDforOrgID(OrgID.Value)
                Me.center.Text = Request("ORGNAME")
                Me.ClassChar.SelectedValue = Request("CHARID")

                Me.Org.Disabled = False
                Org.Attributes("onclick") = ""

                'me.center.Te
                'ElseIf Request("ProcessType") = "Insert" Then
            Else
                If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
                    Org.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg.aspx');"
                Else
                    Org.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg1.aspx');"
                End If

                TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
                If HistoryRID.Rows.Count <> 0 Then
                    center.Attributes("onclick") = "showObj('HistoryList2');"
                    center.Style("CURSOR") = "hand"
                End If

            End If

        End If

        If Not Session("_search") Is Nothing Then
            Me.ViewState("_search") = Session("_search")
            Session("_search") = Nothing
        End If
    End Sub

    Private Sub bt_addrow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_addrow.Click
        Dim sql As String
        Dim dr As DataRow
        Dim da As OracleDataAdapter = Nothing
        Dim dt As DataTable
        'Dim i As Integer
        If OrgID.Value = "" Then OrgID.Value = TIMS.Get_OrgID(Me.RIDValue.Value, objconn)

        'Dim trans As OracleTransaction = Nothing
        Dim conn As OracleConnection = DbAccess.GetConnection()
        Dim trans As OracleTransaction = DbAccess.BeginTrans(conn)
        Try
            sql = "DELETE Org_YearChar WHERE OrgID ='" & Me.OrgID.Value & "' and PlanYear ='" & yearlist_add.SelectedItem.Value & "'"
            DbAccess.ExecuteNonQuery(sql, trans)
            DbAccess.CommitTrans(trans)

            trans = DbAccess.BeginTrans(conn)
            sql = " SELECT * FROM Org_YearChar WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, trans)
            dr = dt.NewRow
            dt.Rows.Add(dr)

            dr("OrgID") = OrgID.Value
            dr("PlanYear") = yearlist_add.SelectedItem.Value
            dr("CHARID") = ClassChar.SelectedItem.Value
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now

            DbAccess.UpdateDataTable(dt, da, trans)
            DbAccess.CommitTrans(trans)
            If Request("ProcessType") = "Update" Then
                Common.MessageBox(Me, "修改資料完成!")
            Else
                Common.MessageBox(Me, "新增資料完成!")
            End If
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Throw ex
        End Try

        Session("_search") = Me.ViewState("_search")
        TIMS.Utl_Redirect1(Me, "TC_01_012.aspx?ID=" & Request("ID"))
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Session("_search") = Me.ViewState("_search")
        'Response.Redirect("TC_01_012.aspx?ID=" & Request("ID"))
    End Sub

End Class
