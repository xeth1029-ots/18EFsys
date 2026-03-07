Partial Class TC_01_007_copy
    Inherits AuthBasePage
    Dim rqIDNO As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        rqIDNO = TIMS.sUtl_GetRqValue(Me, "IDNO", rqIDNO)

        If Not IsPostBack Then
            If rqIDNO = "" Then
                Common.RespWrite(Me, "<script>alert('查無資料');window.close();</script>")
                Exit Sub
            End If
            Call create()
        End If

    End Sub

    Sub create()
        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow
        sql = ""
        sql &= " SELECT a.*,dbo.NVL(c.PlanName,'勞動部勞動力發展署') PlanName " & vbCrLf
        sql &= " FROM (SELECT * FROM Teach_TeacherInfo WHERE IDNO='" & rqIDNO & "' " & vbCrLf
        sql &= " and RID IN (SELECT RID FROM Auth_Relship WHERE OrgID='" & sm.UserInfo.OrgID & "')) a " & vbCrLf
        sql &= " LEFT JOIN view_RIDName b ON a.RID=b.RID " & vbCrLf
        sql &= " LEFT JOIN view_LoginPlan c ON b.PlanID=c.PlanID " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            Common.RespWrite(Me, "<script>alert('查無資料');window.close();</script>")
            Exit Sub
        End If

        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")

                Radio1.Value = drv("TechID").ToString
                Radio1.Attributes("onclick") = "ReturnMyValue(this.value);"

                e.Item.Cells(1).Text = drv("PlanName").ToString
        End Select
    End Sub

End Class
