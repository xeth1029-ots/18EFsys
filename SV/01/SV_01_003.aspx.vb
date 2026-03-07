Partial Class SV_01_003
    Inherits AuthBasePage
    'Dim IptName As String

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
        '檢查Session是否存在 End

        '分頁設定
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            PageControler1.Visible = False
            If Request("IptName") <> "" Then '按回上一頁
                Ipt_Name.Value = Request("IptName")
                Ipt_Name.Value = TIMS.ClearSQM(Ipt_Name.Value)

                'search_ServerClick(sender, e)
                Call dt_search() '查詢
            End If

        End If

    End Sub

    '查詢
    Sub dt_search()
        Ipt_Name.Value = TIMS.ClearSQM(Ipt_Name.Value)
        Dim sql As String = ""
        sql = ""
        sql &= " select SVID, Name"
        sql &= " ,case Avail when 'Y' then '啟用' else '不啟用' end Avail"
        sql &= " from ID_Survey"
        sql &= " where 1=1"
        sql &= " and Avail <> 'N'"
        If Ipt_Name.Value <> "" Then   '搜尋條件
            sql &= " and Name like '%" & Ipt_Name.Value & "%' "
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        msg.Visible = True
        table_F.Visible = True
        Table2.Visible = True
        DataGrid1.Visible = False
        PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            'msg.Visible = True
            'table_F.Visible = True
            'Table2.Visible = True
            DataGrid1.Visible = True
            PageControler1.Visible = True
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Btn_edit As Button = e.Item.FindControl("Btn_edit")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SVID", Convert.ToString(drv("SVID")))
                Btn_edit.CommandArgument = sCmdArg

                '找出標題
                Dim dt As DataTable = TIMS.Get_dtKSK(Convert.ToString(drv("SVID")), objconn)
                If dt.Rows.Count = 0 Then
                    Btn_edit.Enabled = False
                    Btn_edit.ToolTip = "尚未設定【問卷分類標題設定】功能的問卷分類標題!!"
                    Btn_edit.CommandArgument = ""
                End If
        End Select


    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edit"
                Dim sCmdArg As String = e.CommandArgument
                Dim SVID As String = TIMS.GetMyValue(sCmdArg, "SVID")
                Ipt_Name.Value = TIMS.ClearSQM(Ipt_Name.Value)
                Dim sUrl As String = "SV_01_003_Insert.aspx?ID=" & Request("ID") & "&SVID=" & SVID & "&IptName=" & Ipt_Name.Value & ""
                TIMS.Utl_Redirect1(Me, sUrl) '傳SVID 的值及導向新增頁面程式
        End Select

    End Sub

    '查詢
    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Call dt_search() '查詢
    End Sub
End Class
