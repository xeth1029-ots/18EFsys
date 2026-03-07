Partial Class OB_01_007
    Inherits AuthBasePage
    Dim re_item As Int16 = 0

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me) '☆
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = dg_Sch
        If Not IsPostBack Then
            'years()
            ddl_years = TIMS.GetSyear(ddl_years, Year(Now) - 1, Year(Now) + 3, True)
            ddl_TPlanID = TIMS.Get_TPlan(ddl_TPlanID)

            'btn_choose.Attributes("onclick") = "return choose_stn();"
            'btn_select.Attributes("onclick") = "return check_select();"

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

    'Sub years()
    '    ddl_years.Items.Add(New ListItem("請選擇", ""))
    '    For i As Int16 = 0 To 4
    '        ddl_years.Items.Add(New ListItem((Year(Now) + i).ToString, (Year(Now) + i).ToString))
    '    Next
    'End Sub

    Private Sub btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        search()
    End Sub

    Sub search()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch) '顯示列數不正確

        Dim sql As String
        'sql = "select otd.tsn,otd.years,kp.PlanName,otd.TenderName,convert(varchar,otd.TenderSDate,111) TenderSDate,"
        'sql += "count(otd.tsn) num from OB_Tender otd join OB_Tcontent ott on ott.tsn=otd.tsn join Key_Plan kp on "
        'sql += "kp.TPlanID=otd.TPlanID where 1=1"
        sql = "" & vbCrLf
        sql += " SELECT ot.tsn" & vbCrLf
        sql += " ,ot.years" & vbCrLf
        sql += " ,ot.TenderCName" & vbCrLf
        sql += " ,CONVERT(varchar, ot.TenderSDate, 111) TenderSDate" & vbCrLf
        sql += " ,op.PlanName" & vbCrLf
        sql += " ,oti.tisn" & vbCrLf
        sql += " ,dbo.NVL(otm.tnum,0) tnum " & vbCrLf
        sql += " ,dbo.NVL(otc.cnum,0) cnum " & vbCrLf
        sql += " from OB_Tender ot " & vbCrLf
        sql += " join ob_Plan op on op.PlanSN=ot.PlanSN" & vbCrLf
        sql += " LEFT join OB_TenderItem oti on oti.tsn=ot.tsn " & vbCrLf
        sql += " LEFT join (select tsn,count(*) tnum from ob_tmember group by tsn) otm on otm.tsn=ot.tsn " & vbCrLf
        sql += " LEFT join (select tsn,count(*) cnum  from OB_Tcontent group by tsn) otc on otc.tsn=ot.tsn " & vbCrLf
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

        'sql += " group by otd.tsn,otd.years,kp.PlanName,otd.TenderName,otd.TenderSDate"& vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            msg.Visible = True
            Panel_View.Visible = False
            PageControler1.Visible = False
            Exit Sub
        End If

        msg.Visible = False
        Panel_View.Visible = True
        'PageControler1.SqlString = sql
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Private Sub dg_Sch_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound

        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1

            Dim btn_add As Button = e.Item.FindControl("btn_add")
            btn_add.CommandArgument = drv("tsn").ToString

            Dim btnedit As Button = e.Item.FindControl("btn_edit")
            btnedit.CommandArgument = drv("tsn").ToString

            Dim btndel As Button = e.Item.FindControl("btn_del")
            btndel.CommandArgument = drv("tsn").ToString
            btndel.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆全部資料?');"

            Dim btnprt As Button = e.Item.FindControl("btn_prt")
            btnprt.Attributes("onclick") = ReportQuery.ReportScript(Me, "OB", "OB_01_007_0", "tsn=" & drv("tsn"))

            btn_add.Enabled = False
            btnedit.Enabled = False
            btndel.Enabled = False
            btnprt.Enabled = False
            If drv("tnum") > 0 And drv("tisn").ToString <> "" Then
                btn_add.Enabled = True
                btnedit.Enabled = True
                btndel.Enabled = True
                btnprt.Enabled = True
            Else
                TIMS.Tooltip(btndel, "未設定評選工作小組或標案評選項目!!")
                TIMS.Tooltip(btnprt, "未設定評選工作小組或標案評選項目!!")

                If drv("tnum") = 0 Then
                    TIMS.Tooltip(btn_add, "未設定評選工作小組!!")
                    TIMS.Tooltip(btnedit, "未設定評選工作小組!!")
                End If

                If drv("tisn").ToString = "" Then
                    TIMS.Tooltip(btn_add, "未設定標案評選項目!!")
                    TIMS.Tooltip(btnedit, "未設定標案評選項目!!")
                End If
            End If

            If drv("cnum") > 0 Then
                btn_add.Visible = False
                TIMS.Tooltip(btn_add, "委外訓練之投標廠商內容已經設定~")

                btnedit.Visible = True
                btndel.Visible = True
            Else
                btn_add.Visible = True
                btnedit.Visible = False
                TIMS.Tooltip(btnedit, "委外訓練之投標廠商內容未設定!!")

                btndel.Visible = False
                TIMS.Tooltip(btndel, "委外訓練之投標廠商內容未設定!!")
            End If
        End If
    End Sub

    Private Sub dg_Sch_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        Dim sql As String
        Select Case e.CommandName
            Case "add", "edit"
                Me.ViewState("un") = "edit" 'e.CommandName

                Panel_Sch.Visible = False
                Panel_View.Visible = False
                Panel_edit.Visible = True

                txt_tsn.Value = e.CommandArgument
                search2(txt_tsn.Value)

            Case "del"
                sql = "delete OB_Tcontent where tsn=" & e.CommandArgument
                DbAccess.ExecuteNonQuery(sql, objconn)

                Common.MessageBox(Me, "資料已刪除!!")
                search()

        End Select
    End Sub

    Sub search2(ByVal TSNVAL As String)
        Dim sql As String
        Dim dt As DataTable

        sql = "" & vbCrLf
        sql += " select a.msn" & vbCrLf
        sql += " ,b.memname " & vbCrLf
        sql += " ,dbo.NVL(c.cnt,0) Ccnt" & vbCrLf
        sql += " ,oti.tisn" & vbCrLf
        sql += " from ob_tmember a" & vbCrLf
        sql += " join ob_member b on b.msn=a.msn " & vbCrLf
        sql += " LEFT join OB_TenderItem oti on oti.tsn=a.tsn " & vbCrLf
        sql += " LEFT join (select msn,tsn, count(*) cnt from ob_tcontent group by msn,tsn) c " & vbCrLf
        sql += " on c.msn=a.msn and c.tsn=a.tsn" & vbCrLf
        sql += " where a.tsn= " & TSNVAL & vbCrLf

        'sql = "select distinct a.msn,b.memname from ob_tmember a join ob_member b on b.msn=a.msn join ob_tcontent c "
        'sql += "on c.msn=a.msn where a.tsn=" & txt_tsn.Value
        dt = DbAccess.GetDataTable(sql, objconn)
        dg_edit.DataSource = dt
        dg_edit.DataBind()
    End Sub

    Private Sub dg_edit_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_edit.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1

            Dim btn_add2 As Button = e.Item.FindControl("btn_add2")
            btn_add2.CommandArgument = drv("msn").ToString

            Dim btnedit2 As Button = e.Item.FindControl("btn_edit2")
            btnedit2.CommandArgument = drv("msn").ToString

            Dim btndel2 As Button = e.Item.FindControl("btn_del2")
            btndel2.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆資料?');"
            btndel2.CommandArgument = drv("msn").ToString

            Dim btnprt2 As Button = e.Item.FindControl("btn_prt2")
            btnprt2.Attributes("onclick") = ReportQuery.ReportScript(Me, "OB", "OB_01_007_1", "tsn=" & txt_tsn.Value & "&msn=" & drv("msn") & "&tisn=" & drv("tisn"))

            If drv("Ccnt") > 0 Then
                btn_add2.Visible = False
                btnedit2.Visible = True
                btndel2.Visible = True
                btnprt2.Visible = True
            Else
                btn_add2.Visible = True
                btnedit2.Visible = False
                btndel2.Visible = False
                btnprt2.Visible = False
            End If
            'Dim sql As String
            'Dim dr As DataRow
            'sql = "select tisn from ob_tenderitem where tsn=" & txt_tsn.Value
            'dr = DbAccess.GetOneRow(sql)
        End If
    End Sub

    Private Sub dg_edit_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_edit.ItemCommand
        Select Case e.CommandName
            Case "add"
                Panel_edit.Visible = False
                Panel_Add_Edit.Visible = True
                txt_msn.Value = e.CommandArgument
                txt_Name.Text = TIMS.Get_OB_Tendere(txt_tsn.Value, "TenderCName")

                Dim sql As String
                Dim dt As DataTable
                Dim dt2 As DataTable
                Dim dt_tmp As New DataTable
                'Dim dt_ORGnum As DataTable '投標廠商筆數
                'Dim dt_ORSNnum As DataTable '評選項目筆數
                'Dim dr As DataRelation

                TIMS.Get_TMember(ddl_Member, txt_tsn.Value, objconn)
                Common.SetListItem(ddl_Member, e.CommandArgument)
                ddl_Member.Enabled = False

                'sql = "select distinct convert(varchar,a.tsn,111)+','+convert(varchar,b.msn,111)+','+convert(varchar,c.csn,"
                'sql += "111) id,orgname,memname from ob_tender a join ob_tmember b on a.tsn=b.tsn join ob_tcontractor c on "
                'sql += "c.tsn=b.tsn join ob_contractor d on d.csn=c.csn join ob_tenderitem e on e.tsn=a.tsn join "
                'sql += "ob_titemdetail f on f.tisn=e.tisn join ob_reviewitem g on g.orsn=f.orsn join ob_member h on h.msn=b.msn "
                'sql += "where a.tsn=" & txt_tsn.Value & " and b.msn=" & ddl_Member.SelectedValue
                'sql += " and g.orlevels=0 group by a.tsn,b.msn,c.csn,orgname,memname"

                sql = "" & vbCrLf
                sql += " select " & vbCrLf
                sql += " CONVERT(varchar, ot.tsn, 111) + ',' + CONVERT(varchar, otm.msn, 111) + ',' + CONVERT(varchar, ot.csn, 111) id" & vbCrLf
                sql += " ,oc.orgname,memname " & vbCrLf
                sql += " ,ot.tsn ,otm.msn ,ot.csn" & vbCrLf
                sql += " from ob_tcontractor ot" & vbCrLf
                sql += " join ob_tmember otm on ot.tsn=otm.tsn" & vbCrLf
                sql += " join ob_member om on otm.msn=om.msn " & vbCrLf
                sql += " join ob_contractor oc on ot.csn=oc.csn " & vbCrLf
                sql += " where 1=1" & vbCrLf
                sql += " and ot.tsn=" & txt_tsn.Value & vbCrLf
                sql += " and otm.msn=" & ddl_Member.SelectedValue & vbCrLf
                sql += " order by ot.tsn ,otm.msn ,ot.csn" & vbCrLf
                dt = DbAccess.GetDataTable(sql, objconn)
                dg_item.DataSource = dt
                dg_item.DataBind()

                'sql = "select distinct f.sort1,f.orsn,g.orname from ob_tender a join ob_tmember b on a.tsn=b.tsn "
                'sql += "join ob_tcontractor c on c.tsn=b.tsn join ob_contractor d on d.csn=c.csn join ob_tenderitem "
                'sql += "e on e.tsn=a.tsn join ob_titemdetail f on f.tisn=e.tisn join ob_reviewitem g on g.orsn=f.orsn "
                'sql += "where a.tsn=" & txt_tsn.Value & " and msn=" & ddl_Member.SelectedValue & " and g.orlevels=0 "
                'sql += "group by f.sort1,f.orsn,g.orname order by f.sort1"

                sql = "" & vbCrLf
                sql += " select  f.tisn " & vbCrLf
                sql += " , f.sort1" & vbCrLf
                sql += " , f.orsn" & vbCrLf
                sql += " , g.orname " & vbCrLf
                sql += " , e.tsn " & vbCrLf
                sql += " from ob_ReviewItem g" & vbCrLf
                sql += " JOIN ob_TItemDetail f on g.orsn=f.orsn" & vbCrLf
                sql += " JOIN ob_tenderitem  e on f.tisn=e.tisn " & vbCrLf
                sql += " where 1=1" & vbCrLf
                sql += " and g.orlevels=0" & vbCrLf
                sql += " and e.tsn=" & txt_tsn.Value & vbCrLf
                sql += " order by f.sort1, f.orsn" & vbCrLf
                dt2 = DbAccess.GetDataTable(sql, objconn)

                For i As Int16 = 1 To dt2.Rows.Count
                    If i = 1 Then
                        txt_ORSN.Value = dt2.Rows(i - 1)("orsn").ToString
                    Else
                        txt_ORSN.Value += "," + dt2.Rows(i - 1)("orsn").ToString
                    End If
                    Dim num As Int16 = 2 + i
                    dg_item.Columns(num).Visible = True
                    dg_item.Columns(num).HeaderText = i.ToString + "、" + dt2.Rows(i - 1)("orname")
                    dg_item.Columns(num).HeaderStyle.Width = Unit.Pixel(480 / dt.Rows.Count)
                Next
                dg_item.DataBind()
                Panel_Item.Visible = True
                For j As Int16 = 0 To dt.Rows.Count - 1
                    If j = 0 Then
                        dt_tmp.Columns.Add("id")
                        dt_tmp.Columns.Add("orgname")
                        dt_tmp.Columns.Add("memname")
                        dt_tmp.Columns.Add("item1")
                        dt_tmp.Columns.Add("item2")
                        dt_tmp.Columns.Add("item3")
                        dt_tmp.Columns.Add("item4")
                        dt_tmp.Columns.Add("item5")
                        dt_tmp.Columns.Add("item6")
                        dt_tmp.Columns.Add("item7")
                        dt_tmp.Columns.Add("item8")
                        dt_tmp.Columns.Add("item9")
                        dt_tmp.Columns.Add("item10")
                        txt_num.Value = dt2.Rows.Count.ToString
                    End If

                    Dim dr_tmp As DataRow
                    dr_tmp = dt_tmp.NewRow
                    dt_tmp.Rows.Add(dr_tmp)
                    dr_tmp("id") = dt.Rows(j)("id")
                    dr_tmp("orgname") = dt.Rows(j)("orgname")
                    dr_tmp("memname") = ddl_Member.SelectedItem
                    Me.ViewState("tmptable") = dt_tmp
                Next

                dg_item.DataSource = dt_tmp
                dg_item.DataBind()
                'btn_save.Visible = True

                Panel_Item.Visible = True
                btn_save.Visible = True
                btn_lev.Visible = True
            Case "edit"
                Panel_edit.Visible = False
                Panel_Add_Edit.Visible = True
                txt_msn.Value = e.CommandArgument
                txt_Name.Text = TIMS.Get_OB_Tendere(txt_tsn.Value, "TenderCName")

                Dim sql As String
                'Dim dt As DataTable
                Dim dt_tmp As New DataTable
                Dim dt_ORGnum As DataTable '投標廠商筆數
                Dim dt_ORSNnum As DataTable '評選項目筆數
                'Dim dr As DataRelation

                'sql = "select otm.msn,omb.memname from OB_TMember otm join OB_Member omb on omb.msn=otm.msn "
                'sql += "where otm.tsn=" & txt_tsn.Value
                'dt = DbAccess.GetDataTable(sql)
                'sql = "select a.tsn,b.tendername,a.msn from ob_tcontent a join ob_tender b on "
                'sql += "b.tsn=a.tsn join ob_member c on c.msn=a.msn join ob_contractor d on d.csn=a.csn join ob_reviewitem "
                'sql += "e on e.orsn=a.sort1 where a.tsn=" & txt_tsn.Value & " order by a.csn,a.sort1"
                'dt = DbAccess.GetDataTable(sql)
                'dt.Rows(0)("tendername")
                'btn_select_Click(Me, e)

                TIMS.Get_TMember(ddl_Member, txt_tsn.Value, objconn)
                Common.SetListItem(ddl_Member, e.CommandArgument)
                ddl_Member.Enabled = False
                'ddl_Member.SelectedValue = e.CommandArgument


                'sql = "select distinct convert(varchar,a.tsn,111)+','+convert(varchar,a.msn,111)+','+convert(varchar,a.csn,111) id,"
                'sql += "a.msn,memname,a.csn,orgname from ob_tcontent a join ob_contractor b on a.csn=b.csn join ob_member c on a.msn=c.msn "
                'sql += "where tsn=" & txt_tsn.Value & " and a.msn=" & txt_msn.Value

                sql = "" & vbCrLf
                sql += " select CONVERT(varchar, a.tsn, 111)" & vbCrLf
                sql += " + ',' + CONVERT(varchar, a.msn, 111)" & vbCrLf
                sql += " + ',' + CONVERT(varchar, a.csn, 111) id" & vbCrLf
                sql += " ,c.memname" & vbCrLf
                sql += " ,a.tsn, a.msn, a.csn, b.orgname " & vbCrLf
                sql += " from (select tsn,msn,csn,count(*) SortNum from ob_tcontent group by tsn,msn,csn) a " & vbCrLf
                sql += " join ob_contractor b on a.csn=b.csn " & vbCrLf
                sql += " join ob_member c on a.msn=c.msn " & vbCrLf
                sql += " where 1=1" & vbCrLf
                sql += " and a.tsn=" & txt_tsn.Value & vbCrLf
                sql += " and a.msn=" & txt_msn.Value & vbCrLf
                sql += " order by a.tsn,a.msn,a.csn" & vbCrLf
                sql += " " & vbCrLf
                dt_ORGnum = DbAccess.GetDataTable(sql, objconn)

                For i As Int16 = 0 To dt_ORGnum.Rows.Count - 1
                    If i = 0 Then
                        dt_tmp.Columns.Add("id")
                        dt_tmp.Columns.Add("orgname")
                        dt_tmp.Columns.Add("memname")
                        dt_tmp.Columns.Add("item1")
                        dt_tmp.Columns.Add("item2")
                        dt_tmp.Columns.Add("item3")
                        dt_tmp.Columns.Add("item4")
                        dt_tmp.Columns.Add("item5")
                        dt_tmp.Columns.Add("item6")
                        dt_tmp.Columns.Add("item7")
                        dt_tmp.Columns.Add("item8")
                        dt_tmp.Columns.Add("item9")
                        dt_tmp.Columns.Add("item10")
                    Else
                        dt_tmp = Me.ViewState("tmptable")
                    End If

                    Dim dr_tmp As DataRow
                    dr_tmp = dt_tmp.NewRow
                    dt_tmp.Rows.Add(dr_tmp)
                    dr_tmp("id") = dt_ORGnum.Rows(i)("id")
                    dr_tmp("orgname") = dt_ORGnum.Rows(i)("orgname")
                    dr_tmp("memname") = dt_ORGnum.Rows(i)("memname")
                    sql = "select distinct a.sort1,b.orname,content from ob_tcontent a join ob_reviewitem b on a.sort1=b.orsn "
                    sql += "where a.msn=" & dt_ORGnum.Rows(i)("msn") & " and a.csn=" & dt_ORGnum.Rows(i)("csn")
                    dt_ORSNnum = DbAccess.GetDataTable(sql)
                    txt_num.Value = dt_ORSNnum.Rows.Count
                    For j As Int16 = 0 To dt_ORSNnum.Rows.Count - 1
                        If j = 0 Then
                            txt_ORSN.Value = dt_ORSNnum.Rows(j)("sort1").ToString
                        Else
                            txt_ORSN.Value += "," + dt_ORSNnum.Rows(j)("sort1").ToString
                        End If
                        Dim num As Int16 = 3 + j
                        dg_item.Columns(num).Visible = True
                        dg_item.Columns(num).HeaderText = (j + 1).ToString + "、" + dt_ORSNnum.Rows(j)("orname")
                        dg_item.Columns(num).HeaderStyle.Width = Unit.Pixel(420 / dt_ORSNnum.Rows.Count)
                        dt_tmp.Rows(i)("item" + (j + 1).ToString) = dt_ORSNnum.Rows(j)("content")
                        Me.ViewState("tmptable") = dt_tmp
                    Next
                Next
                dg_item.DataSource = dt_tmp
                dg_item.DataBind()
                'btn_clear.Enabled = False

                Panel_Item.Visible = True
                btn_save.Visible = True
                btn_lev.Visible = True

            Case "del"
                Dim sql As String = ""
                sql = "" & vbCrLf
                sql += " delete ob_tcontent where 1=1 " & vbCrLf
                sql += " and msn=" & e.CommandArgument & vbCrLf
                sql += " and tsn=" & txt_tsn.Value & vbCrLf
                DbAccess.ExecuteNonQuery(sql, objconn)

                Common.MessageBox(Me, "資料已刪除!!")
                search2(txt_tsn.Value)
        End Select
    End Sub

    'Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
    '    Panel_Sch.Visible = False
    '    Panel_View.Visible = False
    '    Panel_Add_Edit.Visible = True
    '    btn_lev.Visible = True
    '    Me.ViewState("un") = "add"
    'End Sub

    'Private Sub btn_select_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_select.Click
    '    If txt_Name.Text = "" Then
    '        Common.MessageBox(Me, "請選擇【標案名稱】")
    '    Else
    '        Dim sql As String
    '        Dim dt As DataTable
    '        sql = "select otm.msn,omb.memname from OB_TMember otm join OB_Member omb on omb.msn=otm.msn "
    '        sql += "where otm.tsn=" & txt_tsn.Value
    '        dt = DbAccess.GetDataTable(sql)
    '        With ddl_Member
    '            .DataSource = dt
    '            .DataTextField = "memname"
    '            .DataValueField = "msn"
    '            .DataBind()
    '            ' .Items.Insert(0, New ListItem("請選擇", ""))
    '        End With
    '    End If
    '    btn_select.Visible = False
    '    btn_choose.Enabled = False
    '    btn_clear.Visible = True
    'End Sub

    'Private Sub btn_clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ddl_Member.Items.Clear()
    '    ddl_Member.Items.Insert(0, New ListItem("---", ""))
    '    txt_Name.Enabled = True
    '    txt_Name.Text = ""
    '    txt_tsn.Value = ""
    '    btn_select.Visible = True
    '    btn_choose.Enabled = True
    '    btn_clear.Visible = False
    'End Sub

    'Private Sub ddl_Member_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_Member.SelectedIndexChanged
    '    If ddl_Member.SelectedValue <> "" Then
    '        Dim sql As String
    '        Dim dt As DataTable
    '        Dim dt2 As DataTable
    '        Dim dt_tmp As New DataTable
    '        sql = "select otcsn from ob_tcontent where msn=" & ddl_Member.SelectedValue
    '        If Me.ViewState("un") = "edit" Then
    '            sql += " and msn<>" & txt_msn.Value
    '        End If
    '        dt = DbAccess.GetDataTable(sql)
    '        If dt.Rows.Count > 0 Then
    '            Common.MessageBox(Me, "此工作小組資料已建檔")
    '            ddl_Member.SelectedValue = ""
    '        Else
    '            sql = "select distinct convert(varchar,a.tsn,111)+','+convert(varchar,b.msn,111)+','+convert(varchar,c.csn,"
    '            sql += "111) id,orgname,memname from ob_tender a join ob_tmember b on a.tsn=b.tsn join ob_tcontractor c on "
    '            sql += "c.tsn=b.tsn join ob_contractor d on d.csn=c.csn join ob_tenderitem e on e.tsn=a.tsn join "
    '            sql += "ob_titemdetail f on f.tisn=e.tisn join ob_reviewitem g on g.orsn=f.orsn join ob_member h on h.msn=b.msn "
    '            sql += "where a.tsn=" & txt_tsn.Value & " and b.msn=" & ddl_Member.SelectedValue
    '            sql += " and g.orlevels=0 group by a.tsn,b.msn,c.csn,orgname,memname"
    '            dt = DbAccess.GetDataTable(sql)
    '            dg_item.DataSource = dt
    '            dg_item.DataBind()

    '            sql = "select distinct f.sort1,f.orsn,g.orname from ob_tender a join ob_tmember b on a.tsn=b.tsn "
    '            sql += "join ob_tcontractor c on c.tsn=b.tsn join ob_contractor d on d.csn=c.csn join ob_tenderitem "
    '            sql += "e on e.tsn=a.tsn join ob_titemdetail f on f.tisn=e.tisn join ob_reviewitem g on g.orsn=f.orsn "
    '            sql += "where a.tsn=" & txt_tsn.Value & " and msn=" & ddl_Member.SelectedValue & " and g.orlevels=0 "
    '            sql += "group by f.sort1,f.orsn,g.orname order by f.sort1"
    '            dt2 = DbAccess.GetDataTable(sql)

    '            For i As Int16 = 1 To dt2.Rows.Count
    '                If i = 1 Then
    '                    txt_ORSN.Value = dt2.Rows(i - 1)("orsn").ToString
    '                Else
    '                    txt_ORSN.Value += "," + dt2.Rows(i - 1)("orsn").ToString
    '                End If
    '                Dim num As Int16 = 2 + i
    '                dg_item.Columns(num).Visible = True
    '                dg_item.Columns(num).HeaderText = i.ToString + "、" + dt2.Rows(i - 1)("orname")
    '                dg_item.Columns(num).HeaderStyle.Width = Unit.Pixel(480 / dt.Rows.Count)
    '            Next
    '            dg_item.DataBind()
    '            Panel_Item.Visible = True
    '            For j As Int16 = 0 To dt.Rows.Count - 1
    '                If j = 0 Then
    '                    dt_tmp.Columns.Add("id")
    '                    dt_tmp.Columns.Add("orgname")
    '                    dt_tmp.Columns.Add("memname")
    '                    dt_tmp.Columns.Add("item1")
    '                    dt_tmp.Columns.Add("item2")
    '                    dt_tmp.Columns.Add("item3")
    '                    dt_tmp.Columns.Add("item4")
    '                    dt_tmp.Columns.Add("item5")
    '                    dt_tmp.Columns.Add("item6")
    '                    dt_tmp.Columns.Add("item7")
    '                    dt_tmp.Columns.Add("item8")
    '                    dt_tmp.Columns.Add("item9")
    '                    dt_tmp.Columns.Add("item10")
    '                    txt_num.Value = dt2.Rows.Count.ToString
    '                End If

    '                Dim dr_tmp As DataRow
    '                dr_tmp = dt_tmp.NewRow
    '                dt_tmp.Rows.Add(dr_tmp)
    '                dr_tmp("id") = dt.Rows(j)("id")
    '                dr_tmp("orgname") = dt.Rows(j)("orgname")
    '                dr_tmp("memname") = ddl_Member.SelectedItem
    '                Me.ViewState("tmptable") = dt_tmp
    '            Next

    '            dg_item.DataSource = dt_tmp
    '            dg_item.DataBind()
    '            btn_save.Visible = True
    '        End If
    '    Else
    '        Panel_Item.Visible = False
    '    End If
    'End Sub

    Private Sub dg_item_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_item.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim dt As DataTable = Me.ViewState("tmptable")
            If Not dt Is Nothing Then
                'Dim dr As DataRow
                Dim j As Int16 = 1

                Dim txt_item1 As TextBox = e.Item.FindControl("txt_item1")
                Dim txt_item2 As TextBox = e.Item.FindControl("txt_item2")
                Dim txt_item3 As TextBox = e.Item.FindControl("txt_item3")
                Dim txt_item4 As TextBox = e.Item.FindControl("txt_item4")
                Dim txt_item5 As TextBox = e.Item.FindControl("txt_item5")
                Dim txt_item6 As TextBox = e.Item.FindControl("txt_item6")
                Dim txt_item7 As TextBox = e.Item.FindControl("txt_item7")
                Dim txt_item8 As TextBox = e.Item.FindControl("txt_item8")
                Dim txt_item9 As TextBox = e.Item.FindControl("txt_item9")
                Dim txt_item10 As TextBox = e.Item.FindControl("txt_item10")

                If Int(txt_num.Value) > 0 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item1.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 1 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item2.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 2 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item3.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 3 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item4.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 4 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item5.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 5 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item6.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 6 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item7.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 7 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item8.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 8 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item9.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                If Int(txt_num.Value) > 9 Then
                    If Not IsDBNull(dt.Rows(re_item)("item" + j.ToString)) Then
                        txt_item10.Text = dt.Rows(re_item)("item" + j.ToString)
                        j = j + 1
                    End If
                End If
                re_item = re_item + 1
            End If

        End If
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        '將textbox內容存入dt
        chg_item()

        Dim sql As String
        Dim dt As DataTable = Me.ViewState("tmptable")
        Dim arr() As String
        arr = Split(txt_ORSN.Value, ",")
        If Me.ViewState("un") = "edit" Then
            sql = " delete ob_tcontent where 1=1 " & vbCrLf
            sql += " and tsn=" & txt_tsn.Value & vbCrLf
            sql += " and msn=" & txt_msn.Value & vbCrLf

            DbAccess.ExecuteNonQuery(sql)
        End If
        For i As Int16 = 0 To dt.Rows.Count - 1
            Dim arr2() As String
            arr2 = Split(dt.Rows(i)("id"), ",")
            For j As Int16 = 1 To Int(txt_num.Value)
                Dim str As String = "item" + j.ToString
                sql = "insert into ob_tcontent(tsn,msn,csn,sort1,content,createacct,createtime,modifyacct,modifytime)"
                sql += "values(" & txt_tsn.Value & "," & arr2(1) & "," & arr2(2) & "," & arr(j - 1) & ",'" & dt.Rows(i)(str)
                sql += "','" & sm.UserInfo.UserID & "',getdate() "
                sql += ",'" & sm.UserInfo.UserID & "',getdate())"
                DbAccess.ExecuteNonQuery(sql, objconn)
            Next
        Next

        Me.ViewState("tmptable") = Nothing
        If Me.ViewState("un") = "add" Then
            'Dim strScript As String
            'strScript = "<script language=""javascript"">" + vbCrLf
            'strScript += "if (window.confirm('資料新增成功!\n請問是否繼續新增？')){" + vbCrLf
            'strScript += " document.getElementById('ddl_Member').value=''; " + vbCrLf
            'strScript += "} else {;" + vbCrLf
            'strScript += "  location.href='OB_01_007.aspx';" + vbCrLf
            'strScript &= "}" & vbCrLf
            'strScript &= "</script>"
            'Page.RegisterStartupScript("ring", strScript)
            'Panel_Item.Visible = False
            'btn_save.Visible = False
            'ddl_Member.SelectedValue = ""

            Common.MessageBox(Me, "資料新增成功!!")
            'btn_lev_Click(sender, e)
            Call Utl_lev1()
        Else
            Common.MessageBox(Me, "資料修改成功!!")
            'btn_lev_Click(sender, e)
            Call Utl_lev1()
        End If
        'Me.ViewState("tmptable") = Nothing
    End Sub

    '離開
    Sub Utl_lev1()
        Select Case Convert.ToString(Me.ViewState("un"))
            Case "edit"
                Panel_Add_Edit.Visible = False
                Panel_Item.Visible = False
                Panel_edit.Visible = True

                btn_save.Visible = False
                btn_lev.Visible = False

                'btn_select.Visible = True
                'btn_clear.Enabled = True
                'btn_clear.Visible = False
                'btn_choose.Enabled = True

                ddl_Member.Enabled = True
                ddl_Member.Items.Clear()
                ddl_Member.Items.Insert(0, New ListItem("---", ""))

                Me.ViewState("tmptable") = Nothing
                txt_num.Value = ""
                txt_ORSN.Value = ""
                re_item = 0

                search2(txt_tsn.Value)
        End Select
    End Sub

    '離開
    Private Sub btn_lev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev.Click
        Call Utl_lev1()
        'If Me.ViewState("un") = "add" Then
        '    Common.RespWrite(Me, "<script>location.href='OB_01_007.aspx';</script>")
        'ElseIf Me.ViewState("un") = "edit" Then

        '    Panel_Add_Edit.Visible = False
        '    Panel_Item.Visible = False
        '    Panel_edit.Visible = True

        '    btn_save.Visible = False
        '    btn_lev.Visible = False

        '    'btn_select.Visible = True
        '    'btn_clear.Enabled = True
        '    'btn_clear.Visible = False
        '    'btn_choose.Enabled = True

        '    ddl_Member.Enabled = True
        '    ddl_Member.Items.Clear()
        '    ddl_Member.Items.Insert(0, New ListItem("---", ""))

        '    Me.ViewState("tmptable") = Nothing
        '    txt_num.Value = ""
        '    txt_ORSN.Value = ""
        '    re_item = 0
        'End If
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Panel_Sch.Visible = True
        Panel_View.Visible = True
        Panel_edit.Visible = False
        txt_Name.Text = ""
        txt_tsn.Value = ""
        search()
    End Sub

    '將textbox內容存入dt
    Sub chg_item()
        Dim dt As DataTable = Me.ViewState("tmptable")
        Dim dr As DataRow
        Dim i As Int16 = 0
        For Each item As DataGridItem In dg_item.Items
            dr = dt.Select("id='" & dt.Rows(i)("id") & "'")(0)
            If Int(txt_num.Value) > 0 Then
                Dim txt_item1 As TextBox = item.FindControl("txt_item1")
                dr("item1") = txt_item1.Text
            End If
            If Int(txt_num.Value) > 1 Then
                Dim txt_item2 As TextBox = item.FindControl("txt_item2")
                dr("item2") = txt_item2.Text
            End If
            If Int(txt_num.Value) > 2 Then
                Dim txt_item3 As TextBox = item.FindControl("txt_item3")
                dr("item3") = txt_item3.Text
            End If
            If Int(txt_num.Value) > 3 Then
                Dim txt_item4 As TextBox = item.FindControl("txt_item4")
                dr("item4") = txt_item4.Text
            End If
            If Int(txt_num.Value) > 4 Then
                Dim txt_item5 As TextBox = item.FindControl("txt_item5")
                dr("item5") = txt_item5.Text
            End If
            If Int(txt_num.Value) > 5 Then
                Dim txt_item6 As TextBox = item.FindControl("txt_item6")
                dr("item6") = txt_item6.Text
            End If
            If Int(txt_num.Value) > 6 Then
                Dim txt_item7 As TextBox = item.FindControl("txt_item7")
                dr("item7") = txt_item7.Text
            End If
            If Int(txt_num.Value) > 7 Then
                Dim txt_item8 As TextBox = item.FindControl("txt_item8")
                dr("item8") = txt_item8.Text
            End If
            If Int(txt_num.Value) > 8 Then
                Dim txt_item9 As TextBox = item.FindControl("txt_item9")
                dr("item9") = txt_item9.Text
            End If
            If Int(txt_num.Value) > 9 Then
                Dim txt_item10 As TextBox = item.FindControl("txt_item10")
                dr("item10") = txt_item10.Text
            End If
            i = i + 1
        Next
    End Sub
End Class
