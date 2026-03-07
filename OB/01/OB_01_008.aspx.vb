Partial Class OB_01_008
    Inherits AuthBasePage
    Dim sql As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = dg_Sch
        If Not IsPostBack Then
            ddl_years = TIMS.GetSyear(ddl_years, Year(Now) - 1, Year(Now) + 3, True)
            ddl_TPlanID = TIMS.Get_TPlan(ddl_TPlanID)

            'btn_choose.Attributes("onclick") = "return choose_stn();"
            btn_save.Attributes("onclick") = "return check_save();"
            If Not Session("_SearchStr") Is Nothing Then
                Dim MyArray As Array
                Dim MyItem As String
                Dim MyValue As String
                MyArray = Split(Session("_SearchStr"), "&")
                For i As Integer = 0 To MyArray.Length - 1
                    MyItem = Split(MyArray(i), "=")(0)
                    MyValue = Split(MyArray(i), "=")(1)
                    Select Case MyItem
                        Case "PageIndex"
                            PageControler1.PageIndex = MyValue
                    End Select
                Next
                Panel_View.Visible = False
                Session("_SearchStr") = Nothing
            End If
        End If
    End Sub

    Private Sub btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        search()
    End Sub

    Sub search()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch) '顯示列數不正確

        sql = "" & vbCrLf
        sql += " select ot.tsn" & vbCrLf
        sql += " 	,ot.years" & vbCrLf
        sql += " 	,ot.TenderCName" & vbCrLf
        sql += " 	,op.planname" & vbCrLf
        sql += " 	,oti.tisn" & vbCrLf
        sql += " 	,dbo.NVL(ots.num ,0) num" & vbCrLf
        sql += " 	,dbo.NVL(otid.cnt ,0) otidnum" & vbCrLf
        sql += " 	from ob_tender ot " & vbCrLf
        sql += " 	join ob_plan op on op.plansn=ot.PlanSN " & vbCrLf
        sql += "   left join ob_tenderitem  oti on ot.tsn=oti.tsn " & vbCrLf
        sql += " 	left join (select tsn,count(*) num from ob_tscore2 group by tsn) ots on ots.tsn=ot.tsn " & vbCrLf
        sql += " 	left join (select tisn,  count(*) cnt from ob_titemdetail " & vbCrLf
        sql += " 		where sort2 is null group by tisn ) otid on otid.tisn=oti.tisn " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf

        If sm.UserInfo.DistID <> "000" Then
            sql += " and ot.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If

        If Len(ddl_years.SelectedValue) > 0 Then
            sql += " and ot.years=" & ddl_years.SelectedValue & vbCrLf
        End If
        If Len(ddl_TPlanID.SelectedValue) > 0 Then
            sql += " and ot.TPlanID='" & ddl_TPlanID.SelectedValue & "'" & vbCrLf
        End If
        If PlanName.Text.Trim <> "" Then
            PlanName.Text = PlanName.Text.Trim()
            sql += " AND op.PlanName like '%" & PlanName.Text & "%'" & vbCrLf
        End If

        If Len(txt_TenderCName.Text) > 0 Then
            sql += " and ot.TenderCName like '%" & txt_TenderCName.Text & "%'" & vbCrLf
        End If
        If Len(txt_Sponsor.Text) > 0 Then
            sql += " and ot.Sponsor like '%" & txt_Sponsor.Text & "%'"
        End If


        If TIMS.Get_SQLRecordCount(sql) > 0 Then
            msg.Visible = False
            Panel_View.Visible = True
            PageControler1.SqlString = sql
            PageControler1.ControlerLoad()
        Else
            msg.Visible = True
            Panel_View.Visible = False
            PageControler1.Visible = False
        End If
    End Sub

    Private Sub dg_Sch_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound

        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1

            Dim btnadd As Button = e.Item.FindControl("btn_add")
            btnadd.CommandArgument = drv("tsn").ToString + "," + drv("TenderCName")

            Dim btnedit As Button = e.Item.FindControl("btn_edit")
            btnedit.CommandArgument = drv("tsn").ToString + "," + drv("TenderCName")

            Dim btndel As Button = e.Item.FindControl("btn_del")
            btndel.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆全部資料?');"
            btndel.CommandArgument = drv("tsn")

            Dim btnprt As Button = e.Item.FindControl("btn_prt")
            btnprt.Attributes("onclick") = ReportQuery.ReportScript(Me, "OB", "OB_01_008_0", "tsn=" & drv("tsn"))

            Dim viewmsg As Label = e.Item.FindControl("view_msg")

            'Dim dt As DataTable
            'sql = "select a.tsn from ob_tender a join ob_tenderitem b on b.tsn=a.tsn "
            'sql += "join ob_titemdetail c on c.tisn=b.tisn join ob_reviewitem d on d.orsn=c.orsn "
            'sql += "where a.tsn=" & drv("tsn") & " and sort2 is null order by sort1"
            'dt = DbAccess.GetDataTable(sql)

            If drv("otidnum") = 0 Then
                btnadd.Visible = False
                btnedit.Visible = False
                btndel.Visible = False
                btnprt.Visible = False
                viewmsg.Visible = True
            End If

            If drv("num") = 0 Then
                btnedit.Visible = False
                btndel.Visible = False
            End If
        End If
    End Sub

    Private Sub dg_Sch_ItemCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        Select Case e.CommandName
            Case "add"
                Me.ViewState("un") = "add"
                Dim arr() As String
                arr = Split(e.CommandArgument, ",")
                hid_tsn.Value = arr(0)
                lbl_tsn.Text = arr(1)
                Panel_Sch.Visible = False
                Panel_View.Visible = False
                Panel_Add_Edit.Visible = True
                Dim dt As DataTable
                Dim dt_tmp As New DataTable
                Dim dt_sort1 As DataTable
                Dim dt_sort2 As DataTable

                sql = "select b.csn,b.orgname from ob_tcontractor a join ob_contractor b on a.csn=b.csn where a.tsn=" & arr(0)
                dt = DbAccess.GetDataTable(sql)
                With ddl_tcsn
                    .DataSource = dt
                    .DataTextField = "orgname"
                    .DataValueField = "csn"
                    .DataBind()
                    .Items.Insert(0, New ListItem("請選擇", ""))
                End With

                sql = "select a.tsn + ',' + b.tisn + ',' + c.orsn id,"
                sql += "b.tisn,c.orsn,score,d.orname from ob_tender a join ob_tenderitem b on b.tsn=a.tsn "
                sql += "join ob_titemdetail c on c.tisn=b.tisn join ob_reviewitem d on d.orsn=c.orsn "
                sql += "where a.tsn=" & arr(0) & " and sort2 is null order by sort1"
                dt_sort1 = DbAccess.GetDataTable(sql)
                hid_tisn.Value = dt_sort1.Rows(0)("tisn")

                For i As Int16 = 0 To dt_sort1.Rows.Count - 1
                    If i = 0 Then
                        dt_tmp.Columns.Add("id")
                        dt_tmp.Columns.Add("data")
                        dt_tmp.Columns.Add("score")
                        dt_tmp.Columns.Add("dscore")
                        dt_tmp.Columns.Add("opinion")
                    End If
                    Dim dr_tmp As DataRow
                    dr_tmp = dt_tmp.NewRow
                    dt_tmp.Rows.Add(dr_tmp)
                    dr_tmp("id") = dt_sort1.Rows(i)("id")
                    dr_tmp("score") = dt_sort1.Rows(i)("score")
                    dr_tmp("data") = "<STRONG>" + dt_sort1.Rows(i)("orname") + "</STRONG>" + "<BR>"

                    sql = "select orparent,c.orsn,d.orname from ob_tender a join ob_tenderitem b on b.tsn=a.tsn "
                    sql += "join ob_titemdetail c on c.tisn=b.tisn join ob_reviewitem d on d.orsn=c.orsn "
                    sql += "where sort2 is not null and orparent=" & dt_sort1.Rows(i)("orsn") & " and a.tsn=" & arr(0)
                    sql += "order by sort1,sort2"
                    dt_sort2 = DbAccess.GetDataTable(sql)

                    For j As Int16 = 0 To dt_sort2.Rows.Count - 1
                        If j = dt_sort2.Rows.Count - 1 Then
                            dr_tmp("data") += "(" + (j + 1).ToString + ")" + dt_sort2.Rows(j)("orname").ToString
                        Else
                            dr_tmp("data") += "(" + (j + 1).ToString + ")" + dt_sort2.Rows(j)("orname").ToString + "<BR>"
                        End If
                    Next
                    Me.ViewState("tmptable") = dt_tmp
                Next
                dg_item.DataSource = dt_tmp
                dg_item.DataBind()
            Case "edit"
                Me.ViewState("un") = "edit"
                Dim arr() As String
                arr = Split(e.CommandArgument, ",")
                hid_tsn.Value = arr(0)
                lbl_tsn.Text = arr(1)
                search2()

                Panel_Sch.Visible = False
                Panel_View.Visible = False
                Panel_Edit.Visible = True
            Case "del"
                Dim dr As DataRow
                sql = "select otssn from ob_tscore2 where tsn=" & e.CommandArgument
                dr = DbAccess.GetOneRow(sql)
                sql = "delete ob_tscore2 where otssn=" & dr("otssn")
                DbAccess.ExecuteNonQuery(sql)
                sql = "delete ob_tscore2detail where otssn=" & dr("otssn")
                DbAccess.ExecuteNonQuery(sql)
                Common.MessageBox(Me, "資料已刪除!!")
                search()
        End Select
    End Sub

    Sub search2()
        Dim dt As DataTable
        sql = "select otssn,judgenumber,orgname from ob_tscore2 a join ob_contractor b on b.csn=a.csn "
        sql += "where tsn=" & hid_tsn.Value
        dt = DbAccess.GetDataTable(sql)
        dg_edit.DataSource = dt
        dg_edit.DataBind()
    End Sub

    Private Sub dg_item_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_item.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim dt As DataTable = Me.ViewState("tmptable")
            Dim dt_count As String = dt.Rows.Count
            Dim drv As DataRowView = e.Item.DataItem
            Dim txt_dscore As TextBox = e.Item.FindControl("txt_dscore")
            Dim txt_opinion As TextBox = e.Item.FindControl("txt_opinion")
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1
            txt_dscore.Attributes.Add("onblur", "sum(" & dt_count & ");")
            If Me.ViewState("un") = "edit" Then
                txt_dscore.Text = dt.Rows(e.Item.ItemIndex)("dscore")
                If Not IsDBNull(dt.Rows(e.Item.ItemIndex)("opinion")) Then
                    txt_opinion.Text = dt.Rows(e.Item.ItemIndex)("opinion")
                End If
            End If
        End If
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        Dim dt As DataTable
        Dim dt_tmp As DataTable = Me.ViewState("tmptable")
        Dim dr As DataRow
        sql = "select judgenumber from ob_tscore2 where judgenumber='" & txt_judgenumber.Text & "'"
        If Me.ViewState("un") = "edit" Then
            sql += " and otssn<>" & Me.ViewState("edit_otssn")
        End If
        dt = DbAccess.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            Common.MessageBox(Me, "評選編號重複，請查看是否資料已建置!!")
        Else
            If Me.ViewState("un") = "edit" Then
                sql = "delete ob_tscore2 where otssn=" & Me.ViewState("edit_otssn")
                DbAccess.ExecuteNonQuery(sql)
                sql = "delete ob_tscore2detail where otssn=" & Me.ViewState("edit_otssn")
                DbAccess.ExecuteNonQuery(sql)
            End If
            sql = "insert into ob_tscore2(tsn,tisn,csn,judgenumber,tscore,commnet,judgedate,createacct,createtime) "
            sql += "values(" & hid_tsn.Value & "," & hid_tisn.Value & "," & ddl_tcsn.SelectedValue & ",'" & txt_judgenumber.Text
            sql += "'," & txt_score.Text & ",'" & txt_commnet.Text & "',convert(datetime, '" & txt_judgedate.Text & "', 111),'" & sm.UserInfo.UserID
            sql += "',getdate())"
            DbAccess.ExecuteNonQuery(sql)

            sql = "select otssn from ob_tscore2 where tsn=" & hid_tsn.Value & " and tisn=" & hid_tisn.Value
            sql += " and csn=" & ddl_tcsn.SelectedValue & " and judgenumber='" & txt_judgenumber.Text & "'"
            dr = DbAccess.GetOneRow(sql)

            chg_item()
            For i As Int16 = 0 To dt_tmp.Rows.Count - 1
                Dim arr() As String
                arr = Split(dt_tmp.Rows(i)("id"), ",")
                sql = "insert into ob_tscore2detail(otssn,sort1,dscore,opinion,createacct,createtime) values("
                sql += dr("otssn") & "," & arr(2) & "," & dt_tmp.Rows(i)("dscore") & ",'" & dt_tmp.Rows(i)("opinion")
                sql += "','" & sm.UserInfo.UserID & "',getdate())"
                DbAccess.ExecuteNonQuery(sql)
            Next

            If Me.ViewState("un") = "add" Then
                Common.MessageBox(Me, "資料新增成功!")
                btn_lev_Click(sender, e)
                search()
            Else
                Common.MessageBox(Me, "資料修改成功!")
                btn_lev_Click(sender, e)
                search2()
            End If

        End If
    End Sub

    Private Sub btn_lev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev.Click
        Me.ViewState("tmptable") = Nothing
        txt_judgenumber.Text = ""
        txt_judgedate.Text = ""
        txt_score.Text = "0"
        txt_commnet.Text = ""
        If Me.ViewState("un") = "add" Then
            Panel_Add_Edit.Visible = False
            Panel_View.Visible = True
            Panel_Sch.Visible = True
        Else
            Me.ViewState("edit_otssn") = ""
            Panel_Add_Edit.Visible = False
            Panel_Edit.Visible = True
        End If
    End Sub
    '將textbox內容存入dt
    Sub chg_item()
        Dim dt As DataTable = Me.ViewState("tmptable")
        Dim dr As DataRow
        Dim i As Int16 = 0
        For Each item As DataGridItem In dg_item.Items
            dr = dt.Select("id='" & dt.Rows(i)("id") & "'")(0)
            Dim txt_dscore As TextBox = item.FindControl("txt_dscore")
            Dim txt_opinion As TextBox = item.FindControl("txt_opinion")
            i = i + 1
            dr("dscore") = txt_dscore.Text
            dr("opinion") = txt_opinion.Text
        Next
    End Sub

    Private Sub dg_edit_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_edit.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1
            Dim drv As DataRowView = e.Item.DataItem
            Dim btnedit2 As Button = e.Item.FindControl("btn_edit2")
            Dim btndel2 As Button = e.Item.FindControl("btn_del2")
            Dim btnprt2 As Button = e.Item.FindControl("btn_prt2")
            btndel2.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆資料?');"
            btnprt2.Attributes("onclick") = ReportQuery.ReportScript(Me, "OB", "OB_01_008_1", "tsn=" & hid_tsn.Value & "&otssn=" & drv("otssn"))
            btnedit2.CommandArgument = drv("otssn")
            btndel2.CommandArgument = drv("otssn")
        End If
    End Sub

    Private Sub dg_edit_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_edit.ItemCommand
        Select Case e.CommandName
            Case "edit"
                Me.ViewState("edit_otssn") = e.CommandArgument
                Dim dt As DataTable
                Dim dt_count As DataTable
                Dim dt_item As DataTable
                Dim dt_tmp As New DataTable
                Dim dr As DataRow
                sql = "select b.csn,b.orgname from ob_tcontractor a join ob_contractor b on a.csn=b.csn where a.tsn=" & hid_tsn.Value
                dt = DbAccess.GetDataTable(sql)
                With ddl_tcsn
                    .DataSource = dt
                    .DataTextField = "orgname"
                    .DataValueField = "csn"
                    .DataBind()
                    .Items.Insert(0, New ListItem("請選擇", ""))
                    .SelectedValue = ""
                End With

                sql = "select tisn,csn,judgenumber,CONVERT(varchar, judgedate, 111) judgedate,tscore,commnet "
                sql += "from ob_tscore2 where otssn=" & e.CommandArgument
                dr = DbAccess.GetOneRow(sql)
                hid_tisn.Value = dr("tisn").ToString
                txt_judgenumber.Text = dr("judgenumber").ToString
                txt_judgedate.Text = dr("judgedate").ToString
                ddl_tcsn.SelectedValue = dr("csn").ToString
                txt_score.Text = dr("tscore").ToString
                If Not IsDBNull(dr("commnet")) Then
                    txt_commnet.Text = dr("commnet").ToString
                End If

                sql = "select distinct a.tsn + ',' + a.tisn + ',' + b.sort1 id "
                sql += ",tisn,dscore,opinion from ob_tscore2 a join ob_tscore2detail b "
                sql += "on a.otssn=b.otssn where a.tsn=" & hid_tsn.Value & " and a.tisn=" & dr("tisn")
                sql += " and a.csn = " & ddl_tcsn.SelectedValue & " and a.judgenumber='" & txt_judgenumber.Text & "'"
                dt_count = DbAccess.GetDataTable(sql)
                For i As Int16 = 0 To dt_count.Rows.Count - 1
                    If i = 0 Then
                        dt_tmp.Columns.Add("id")
                        dt_tmp.Columns.Add("data")
                        dt_tmp.Columns.Add("score")
                        dt_tmp.Columns.Add("dscore")
                        dt_tmp.Columns.Add("opinion")
                    Else
                        dt_tmp = Me.ViewState("tmptable")
                    End If
                    Dim dr_tmp As DataRow
                    dr_tmp = dt_tmp.NewRow
                    dt_tmp.Rows.Add(dr_tmp)
                    dr_tmp("id") = dt_count.Rows(i)("id")
                    dr_tmp("dscore") = dt_count.Rows(i)("dscore")
                    dr_tmp("opinion") = dt_count.Rows(i)("opinion")

                    sql = "select score,b.sort2,c.orname from ob_tenderitem a join ob_titemdetail b on a.tisn=b.tisn "
                    sql += "join ob_reviewitem c on c.orsn=b.orsn where a.tisn=" & dr("tisn") & " and sort1=" & i + 1
                    sql += " order by sort2"
                    dt_item = DbAccess.GetDataTable(sql)
                    For j As Int16 = 0 To dt_item.Rows.Count - 1
                        If j = 0 Then
                            dr_tmp("data") = "<STRONG>" + dt_item.Rows(j)("orname") + "</STRONG>" + "<BR>"
                            dr_tmp("score") = dt_item.Rows(j)("score")
                        ElseIf j = dt_item.Rows.Count - 1 Then
                            dr_tmp("data") += "(" + (j).ToString + ")" + dt_item.Rows(j)("orname").ToString
                        Else
                            dr_tmp("data") += "(" + (j).ToString + ")" + dt_item.Rows(j)("orname").ToString + "<BR>"
                        End If
                    Next
                    Me.ViewState("tmptable") = dt_tmp
                Next

                dg_item.DataSource = dt_tmp
                dg_item.DataBind()

                Panel_Edit.Visible = False
                Panel_Add_Edit.Visible = True
            Case "del"
                sql = "delete ob_tscore2 where otssn=" & e.CommandArgument
                DbAccess.ExecuteNonQuery(sql)
                sql = "delete ob_tscore2detail where otssn=" & e.CommandArgument
                DbAccess.ExecuteNonQuery(sql)
                Common.MessageBox(Me, "資料已刪除!!")
                search2()
        End Select
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Panel_Edit.Visible = False
        Panel_Sch.Visible = True
        Panel_View.Visible = True
        search()
    End Sub
End Class
