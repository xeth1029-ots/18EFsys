Partial Class OB_01_001
    Inherits AuthBasePage

    Const Cst_TenderSDate As Integer = 5

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            ddlyears = TIMS.GetSyear(ddlyears, Year(Now) - 1, Year(Now) + 3, True)
            TPlanID = TIMS.Get_TPlan(TPlanID)
            DataGridTable.Visible = False
        End If

        Call GetKeepSearch()
    End Sub

    Sub KeepSearch()
        Session("_SearchStr") = "Class=OB_01_001"
        Session("_SearchStr") += "&ddlyears=" & ddlyears.SelectedValue
        Session("_SearchStr") += "&txttsn=" & txttsn.Text.Trim
        Session("_SearchStr") += "&TPlanID=" & TPlanID.SelectedValue
        Session("_SearchStr") += "&PlanName=" & PlanName.Text.Trim
        Session("_SearchStr") += "&Sponsor=" & Sponsor.Text.Trim

        If DataGridTable.Visible Then
            Session("_SearchStr") += "&submit=1"
        Else
            Session("_SearchStr") += "&submit=0"
        End If
    End Sub

    Sub GetKeepSearch()
        '執行 Session("_SearchStr") 保留查詢值
        If Not Session("_SearchStr") Is Nothing Then
            Me.ViewState("_SearchStr") = Session("_SearchStr")
            If TIMS.GetMyValue(Me.ViewState("_SearchStr"), "Class") = "OB_01_001" Then
                Common.SetListItem(ddlyears, TIMS.GetMyValue(Me.ViewState("_SearchStr"), "ddlyears"))
                txttsn.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "txttsn")
                Common.SetListItem(TPlanID, TIMS.GetMyValue(Me.ViewState("_SearchStr"), "TPlanID"))
                PlanName.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PlanName")
                Sponsor.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "Sponsor")
                If TIMS.GetMyValue(Me.ViewState("_SearchStr"), "submit") = "1" Then
                    'btnQuery_Click(sender, e)
                    Call search1()
                End If
            End If
            Session("_SearchStr") = Nothing
        End If
    End Sub

    Sub search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        'Optional ByVal sql As String = ""
        Dim sql As String = ""
        sql += " select b.PlanName, a.* " & vbCrLf
        sql += " from OB_Tender a " & vbCrLf
        sql += " JOIN OB_Plan b on a.PlanSN=b.PlanSN " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If

        If ddlyears.SelectedValue <> "" Then
            sql += " AND a.Years='" & ddlyears.SelectedValue & "'" & vbCrLf
        End If
        If TPlanID.SelectedValue <> "" Then
            sql += " AND a.TPlanID='" & TPlanID.SelectedValue & "'" & vbCrLf
        End If
        If PlanName.Text.Trim <> "" Then
            PlanName.Text = PlanName.Text.Trim()
            sql += " AND b.PlanName like '%" & PlanName.Text & "%'" & vbCrLf
        End If
        If TenderCName.Text.Trim <> "" Then
            TenderCName.Text = TenderCName.Text.Trim()
            sql += " AND a.TenderCName like '%" & TenderCName.Text & "%'" & vbCrLf
        End If
        If Sponsor.Text.Trim <> "" Then
            Sponsor.Text = Sponsor.Text.Trim()
            sql += " AND a.Sponsor like '%" & Sponsor.Text & "%'" & vbCrLf
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        Else
            DataGridTable.Visible = False
            msg.Text = "查無資料!!"
        End If

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call search1()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "OB_01_001_add.aspx?ID=" & Request("ID") & "&Action=ADD")
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound

        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(Cst_TenderSDate).Text = Common.FormatDate(drv("TenderSDate"))

                Dim btn_view As Button = e.Item.FindControl("btn_view")
                btn_view.CommandArgument = drv("tsn")

                Dim btn_edit As Button = e.Item.FindControl("btn_edit")
                btn_edit.CommandArgument = drv("tsn")

                Dim btn_del As Button = e.Item.FindControl("btn_del")
                btn_del.CommandArgument = drv("tsn")
                btn_del.Attributes("onclick") = TIMS.cst_confirm_delmsg1

                Dim btn_mem As Button = e.Item.FindControl("btn_mem")
                btn_mem.CommandArgument = drv("tsn")
                'btn_mem.Attributes.Add("onclick", "wopen('OB_01_001_mem.aspx?ID=" & Request("ID") & "&Action=MEM&tsn=" & drv("tsn") & "','工作小組成員','630','550','yes');")

                Dim btn_con As Button = e.Item.FindControl("btn_con")
                btn_con.CommandArgument = drv("tsn")
                'btn_con.Attributes.Add("onclick", "wopen('OB_01_004.aspx?ID=" & Request("ID") & "&Action=con&tsn=" & drv("tsn") & "','投標廠商設定','630','550','yes');")
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        KeepSearch()

        Dim sql As String = ""
        Select Case e.CommandName
            Case "del"
                sql = "delete OB_Tender where tsn='" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功！")
                'Query()
                Call search1()
            Case "edit"
                Call TIMS.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, "OB_01_001_add.aspx?ID=" & Request("ID") & "&Action=EDIT&tsn=" & e.CommandArgument)
            Case "view"
                Call TIMS.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, "OB_01_001_add.aspx?ID=" & Request("ID") & "&Action=VIEW&tsn=" & e.CommandArgument)
            Case "mem" '工作小組成員
                Call TIMS.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, "OB_01_001_mem.aspx?ID=" & Request("ID") & "&Action=MEM&tsn=" & e.CommandArgument)
            Case "con" '投標廠商設定
                Call TIMS.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, "OB_01_004.aspx?ID=" & Request("ID") & "&Action=con&tsn=" & e.CommandArgument)
        End Select
    End Sub
End Class
