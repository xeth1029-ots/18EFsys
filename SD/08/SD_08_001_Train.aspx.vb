Partial Class SD_08_001_Train
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow
        sql = "SELECT LTCID,LTCID+'：'+LTCName as LTCName FROM Key_LapmTCode ORDER BY LTCID"
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Attributes("onclick") = "ReturnMyValue('" & drv("LTCID") & "','" & drv("LTCName") & "');"
                e.Item.Style("Cursor") = "hand"
                e.Item.Attributes("onmouseover") = "this.style.backgroundColor='#EFEFEF'"
                e.Item.Attributes("onmouseout") = "this.style.backgroundColor=''"
        End Select
    End Sub
End Class
