Partial Class TC_10_MBR
    Inherits AuthBasePage

    Dim dtDist As DataTable = Nothing 'TIMS.Get_DistIDdt(objconn)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'wopen('TC_10_TechID.aspx?TextField=' + nText + '&ValueField=' + nValue, 'sendEXAMINER', 700, 360, 0);
        dtDist = TIMS.Get_DISTIDdt(objconn) 'Dim dtDist As DataTable = TIMS.Get_DistIDdt(objconn)
        If Not IsPostBack Then Create()

    End Sub

    Sub Create()
        sSearch1()
    End Sub


    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim labPUSHDISTID_N As Label = e.Item.FindControl("labPUSHDISTID_N")
                labPUSHDISTID_N.Text = TIMS.Get_PUSHDISTID_NN(dtDist, Convert.ToString(drv("PUSHDISTID")))

                'Radio1.Value = drv("EMSEQ").ToString
                Checkbox1.Value = drv("EMSEQ").ToString
                'Radio1.Visible = False
                Checkbox1.Visible = True

                Checkbox1.Attributes("onclick") = "SelectMBR(this.checked,'" & drv("EMSEQ") & "','" & drv("MBRNAME") & "');"

            Case ListItemType.Footer
                DataGrid1.ShowFooter = False
                If DataGrid1.Items.Count = 0 Then
                    DataGrid1.ShowFooter = True
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(New TableCell)
                    e.Item.Cells(0).ColumnSpan = DataGrid1.Columns.Count
                    e.Item.Cells(0).Text = "查無資料!"
                    e.Item.Cells(0).HorizontalAlign = HorizontalAlign.Center
                End If
        End Select
    End Sub

    Sub sSearch1()
        Dim MBRNAME_lk As String = ""
        MBRNAME.Text = TIMS.ClearSQM(MBRNAME.Text)
        If MBRNAME.Text <> "" Then MBRNAME_lk = String.Format("%{0}%", MBRNAME.Text)

        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT r.EMSEQ" & vbCrLf
        sql &= " ,r.RECRUIT" & vbCrLf
        sql &= " ,CASE r.RECRUIT WHEN 'A' THEN 'A-產業界' WHEN 'B' THEN 'B-學術界' WHEN 'C' THEN 'C-勞工團體代表' END RECRUIT_N" & vbCrLf
        sql &= " ,r.UNITNAME" & vbCrLf
        sql &= " ,r.MBRNAME" & vbCrLf
        sql &= " ,r.JOBTITLE" & vbCrLf
        sql &= " ,r.PUSHDISTID" & vbCrLf
        sql &= " ,r.STOPUSE" & vbCrLf
        sql &= " FROM dbo.OA_EXAMINER r" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        '查詢審查委員	文字框，點選查詢，可從 「啟用中」的全部委員名單做模糊搜尋。搜尋結果列於下方，可供挑選。(如下圖一)
        sql &= " AND ISNULL(r.STOPUSE,'N')!='Y'" & vbCrLf
        If MBRNAME_lk <> "" Then sql &= " AND r.MBRNAME like '" & MBRNAME_lk & "'" & vbCrLf
        'sql &= " ORDER BY RECRUIT,MBRNAME" & vbCrLf
        '順序排序： 1.先依遴聘類別 2.依姓名筆劃 chinese_taiwan_stroke_cs_as_ks_ws CHINESE_TAIWAN_STROKE_CI_AS
        sql &= " ORDER BY r.RECRUIT ,r.MBRNAME COLLATE CHINESE_TAIWAN_STROKE_CI_AS" & vbCrLf
        'sql2 &= " ORDER BY r.RECRUIT ,r.MBRNAME COLLATE Chinese_PRC_Stroke_ci_as " & vbCrLf

        dt = DbAccess.GetDataTable(sql, objconn)
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    '進階搜尋 (查詢)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sSearch1()
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class