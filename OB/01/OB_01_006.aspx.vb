Partial Class OB_01_006
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在 Start
        'If sm.UserInfo.UserID Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');top.location.href='../../MOICA_Login.aspx';</script>")
        '    Response.End()
        'End If
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me) '☆
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = dg_Sch

        If Not IsPostBack Then
            'years()
            ddl_years = TIMS.GetSyear(ddl_years, Year(Now) - 1, Year(Now) + 3, True)
            ORName()
            ddl_TPlanID = TIMS.Get_TPlan(ddl_TPlanID)
            'ddl_TPlanID2 = TIMS.Get_TPlan(ddl_TPlanID2)
            'ddl_TenderName2.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

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
            'If Request("un") = "add" Then
            '    btn_Add_Click(sender, e)
            'End If
            btn_ORItem.Attributes("onclick") = "return check_score();"
            btn_select.Attributes("onclick") = "return check_select();"
        End If
    End Sub

    'Sub years()
    '    ddl_years.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    '    'ddl_years2.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    '    For i As Int16 = 0 To 4
    '        ddl_years.Items.Add(New ListItem((Year(Now) + i).ToString, (Year(Now) + i).ToString))
    '        'ddl_years2.Items.Add(New ListItem((Year(Now) + i).ToString, (Year(Now) + i).ToString))
    '    Next
    'End Sub

    '可用評選項用
    Sub ORName()
        Dim sql As String
        Dim dt As DataTable
        sql = "select ORSN,ORName from OB_ReviewItem where ORLevels=0 and ORAvail='Y' and DistID='"
        sql += sm.UserInfo.DistID & "' "
        sql += "order by ORSN"
        dt = DbAccess.GetDataTable(sql)
        With ddl_ORName
            .DataSource = dt
            .DataTextField = "ORName"
            .DataValueField = "ORSN"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        ddl_ORName.Visible = True
    End Sub

    Private Sub btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        search()
    End Sub

    Sub search()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch) '顯示列數不正確

        Dim sql As String
        'sql = "select ot.tsn,tisn,ot.years,kp.PlanName,ot.TenderName,convert(varchar,ot.TenderSDate,111) TenderSDate "
        'sql += "from OB_Tender ot join OB_TenderItem oti on oti.tsn=ot.tsn join Key_Plan kp on kp.TPlanID="
        'sql += "ot.TPlanID where 1=1"

        sql = "" & vbCrLf
        sql += " SELECT ot.tsn" & vbCrLf
        sql += " 	,ot.years" & vbCrLf
        sql += " 	,ot.TenderCName" & vbCrLf
        sql += " 	,CONVERT(varchar, ot.TenderSDate, 111) TenderSDate" & vbCrLf
        sql += " 	,op.PlanName" & vbCrLf
        sql += " 	,oti.tisn" & vbCrLf
        sql += " from OB_Tender ot " & vbCrLf
        sql += " join ob_Plan op on op.PlanSN=ot.PlanSN" & vbCrLf
        sql += " LEFT join OB_TenderItem oti on oti.tsn=ot.tsn " & vbCrLf
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
            sql += " and ot.Sponsor like '%" & txt_Sponsor.Text & "%'" & vbCrLf
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

    Private Sub dg_Sch_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            'Dim sql As String
            Dim drv As DataRowView = e.Item.DataItem
            'Dim btnview As Button = e.Item.FindControl("btn_view")
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1

            Dim btn_add As Button = e.Item.FindControl("btn_add")
            btn_add.CommandArgument = drv("tsn").ToString

            Dim btnedit As Button = e.Item.FindControl("btn_edit")
            btnedit.CommandArgument = drv("tsn").ToString

            Dim btndel As Button = e.Item.FindControl("btn_del")
            btndel.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆資料?');"
            btndel.CommandArgument = drv("tisn").ToString

            If drv("tisn").ToString <> "" Then
                btn_add.Visible = False
                btnedit.Visible = True
                btndel.Visible = True
            Else
                btn_add.Visible = True
                btnedit.Visible = False
                btndel.Visible = False
            End If
        End If
    End Sub

    Private Sub dg_Sch_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        Dim sql As String
        Dim dt As DataTable
        Dim dt1 As DataTable
        Dim dt2 As DataTable
        Dim dt_item As New DataTable
        dt_item.Columns.Add("id")
        dt_item.Columns.Add("data")
        dt_item.Columns.Add("score")
        Select Case e.CommandName
            Case "add"
                Me.ViewState("un") = e.CommandName '"add"
                Panel_Add_Edit.Visible = True
                Panel_Sch.Visible = False
                Panel_View.Visible = False
                Panel_ListBox.Visible = True


                sql = "" & vbCrLf
                sql += " select ot.TenderCName, ot.years, ot.tplanid, ot.tsn, op.PlanName" & vbCrLf
                sql += " from ob_tender ot" & vbCrLf
                sql += " join ob_Plan op on ot.PlanSN=op.PlanSN" & vbCrLf
                sql += " where 1=1" & vbCrLf
                sql += " and ot.tsn=" & e.CommandArgument & vbCrLf
                'sql = "select ot.years,ot.tplanid,ot.tsn from ob_tender ot where ot.tsn=" & e.CommandArgument

                dt = DbAccess.GetDataTable(sql)
                Me.labYears.Text = dt.Rows(0)("years")
                Me.labPlanName.Text = dt.Rows(0)("PlanName")
                Me.LabTenderCName.Text = dt.Rows(0)("TenderCName")
                Me.ViewState("tsn") = e.CommandArgument
                'ddl_years2.SelectedValue = dt.Rows(0)("years")
                'ddl_TPlanID2.SelectedValue = dt.Rows(0)("tplanid")
                'btn_select_Click(Me, e)

                'ddl_TenderName2.SelectedValue = dt.Rows(0)("tsn")
                ORName() '可用評選項用

            Case "edit"
                Me.ViewState("un") = e.CommandName '"edit"
                Panel_Add_Edit.Visible = True
                Panel_Sch.Visible = False
                Panel_View.Visible = False
                Panel_ListBox.Visible = True

                sql = "" & vbCrLf
                sql += " select ot.TenderCName, ot.years, ot.tplanid, ot.tsn, op.PlanName" & vbCrLf
                sql += " from ob_tender ot" & vbCrLf
                sql += " join ob_Plan op on ot.PlanSN=op.PlanSN" & vbCrLf
                sql += " where 1=1" & vbCrLf
                sql += " and ot.tsn=" & e.CommandArgument & vbCrLf

                'sql = "select ot.years,ot.tplanid,ot.tsn from ob_tender ot join ob_tenderitem oti on oti.tsn"
                'sql += "=ot.tsn join ob_titemdetail otid on otid.tisn=oti.tisn where ot.tsn=" & e.CommandArgument
                dt = DbAccess.GetDataTable(sql)
                Me.labYears.Text = dt.Rows(0)("years")
                Me.labPlanName.Text = dt.Rows(0)("PlanName")
                Me.LabTenderCName.Text = dt.Rows(0)("TenderCName")
                Me.ViewState("tsn") = e.CommandArgument

                'ddl_years2.SelectedValue = dt.Rows(0)("years")
                'ddl_TPlanID2.SelectedValue = dt.Rows(0)("tplanid")
                'btn_select_Click("sender", e)
                'btn_select_Click(Me, e)

                'ddl_TenderName2.SelectedValue = dt.Rows(0)("tsn")
                ORName() '可用評選項用

                'sql = "select distinct sort1 num from ob_tender ot join ob_tenderitem oti on oti.tsn=ot.tsn "
                'sql += "join ob_titemdetail otid on otid.tisn=oti.tisn where ot.tsn=" & e.CommandArgument

                sql = "" & vbCrLf
                sql += " select oti.TIsn, otid.ORSN, otid.Sort1 Num " & vbCrLf
                sql += " from ob_Tender ot " & vbCrLf
                sql += " join ob_Tenderitem oti on oti.tsn=ot.tsn " & vbCrLf
                sql += " join ob_TItemDetail otid on otid.tisn=oti.tisn " & vbCrLf
                sql += " where ot.tsn=" & Me.ViewState("tsn") & vbCrLf
                sql += " and Sort2 is null" & vbCrLf
                sql += " ORDER BY otid.Sort1 NULLS FIRST" & vbCrLf
                dt1 = DbAccess.GetDataTable(sql)
                Me.ViewState("TIsn") = dt1.Rows(0)("TIsn")

                'Dim str As String = ""
                'lbx_Get.Items.Clear()
                LabAction.Text = "新增"
                Me.lbl_num.Text = CInt(dt1.Rows(dt1.Rows.Count - 1)("Num")) + 1
                Me.ViewState("lbl_num") = Me.lbl_num.Text

                '顯示原先的評選項目大項與子項
                For i As Int16 = 0 To dt1.Rows.Count - 1

                    '移除已經新增的項目
                    ddl_ORName.Items.Remove(ddl_ORName.Items.FindByValue(dt1.Rows(i)("ORSN")))

                    'sql = "select oti.tisn,oti.tsn,ori.orsn,ori.orname,otid.sort1,otid.sort2,otid.score from ob_tender ot join "
                    'sql += "ob_tenderitem oti on oti.tsn=ot.tsn join ob_titemdetail otid on otid.tisn=oti.tisn "
                    'sql += "join ob_reviewitem ori on ori.orsn=otid.orsn where ot.tsn=" & e.CommandArgument
                    'sql += " and sort1=" & i + 1 & " order by otid.sort1,otid.sort2"

                    sql = "" & vbCrLf
                    sql += " select oti.TIsn ,oti.tsn ,ri.ORSN ,ri.ORNAME ,otid.Sort1" & vbCrLf
                    sql += " ,otid.Sort2 ,otid.Score " & vbCrLf
                    sql += " from ob_tender ot " & vbCrLf
                    sql += " join ob_Tenderitem oti on oti.tsn=ot.tsn " & vbCrLf
                    sql += " join ob_TItemDetail otid on otid.tisn=oti.tisn " & vbCrLf
                    sql += " join ob_ReviewItem ri on ri.ORSN=otid.ORSN" & vbCrLf
                    sql += " where 1=1" & vbCrLf
                    sql += " AND ot.tsn=" & Me.ViewState("tsn") & vbCrLf
                    sql += " AND otid.sort1=" & dt1.Rows(i)("Num") & vbCrLf
                    sql += " order by otid.sort1 NULLS FIRST,otid.sort2 NULLS FIRST" & vbCrLf
                    dt2 = DbAccess.GetDataTable(sql)

                    Dim dr As DataRow
                    dr = dt_item.NewRow
                    dt_item.Rows.Add(dr)
                    For j As Int16 = 0 To dt2.Rows.Count - 1
                        If j = 0 Then '資料第1筆
                            'Me.ViewState("TIsn") = dt2.Rows(j)("TIsn")

                            'Me.ViewState("edit_tisn") = dt2.Rows(j)("tisn")
                            'Me.ViewState("edit_tsn") = dt2.Rows(j)("tsn")
                            'Me.ViewState("tsn") = dt2.Rows(j)("tsn")

                            dr("id") = dt2.Rows(j)("ORSN").ToString + ","
                            dr("data") = "<STRONG>" + dt2.Rows(j)("ORNAME").ToString + "</STRONG>" + "<BR>"
                            dr("score") = dt2.Rows(j)("Score")

                        ElseIf j = dt2.Rows.Count - 1 Then '資料最後1筆
                            dr("id") += dt2.Rows(j)("ORSN").ToString
                            dr("data") += "(" + dt2.Rows(j)("Sort2").ToString + ")" + dt2.Rows(j)("ORNAME").ToString

                        Else '其他資料筆數
                            dr("id") += dt2.Rows(j)("ORSN").ToString + ","
                            dr("data") += "(" + dt2.Rows(j)("Sort2").ToString + ")" + dt2.Rows(j)("ORNAME").ToString + "<BR>"

                        End If
                    Next
                Next
                Me.ViewState("tmptable") = dt_item
                Me.ViewState("oldtable") = dt_item

                dg_ORItem.DataSource = dt_item
                dg_ORItem.DataBind()

            Case "del"
                sql = "delete OB_TenderItem where tisn=" & e.CommandArgument
                DbAccess.ExecuteNonQuery(sql)

                sql = "delete OB_TItemDetail where tisn=" & e.CommandArgument
                DbAccess.ExecuteNonQuery(sql)

                Common.MessageBox(Me, "資料已刪除!!")

                search()
        End Select
    End Sub

    'Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
    '    Panel_Add_Edit.Visible = True
    '    Panel_Sch.Visible = False
    '    Panel_View.Visible = False
    '    Me.ViewState("un") = "add"
    'End Sub

    'Private Sub btn_select_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_select.Click
    '    Dim sql As String
    '    Dim dt As DataTable
    '    'sql = "select tsn,TenderName from OB_Tender where years='" & ddl_years2.SelectedValue & "' and TPlanID='"
    '    'sql += ddl_TPlanID2.SelectedValue & "' and DistID='" & sm.UserInfo.DistID & "'"

    '    sql = "" & vbCrLf
    '    sql += " select ot.TSN, ot.TenderName" & vbCrLf
    '    sql += " ,ot.years, ot.tplanid, ot.tsn, op.PlanName" & vbCrLf
    '    sql += " from OB_Tender ot" & vbCrLf
    '    sql += " join ob_Plan op on ot.PlanSN=op.PlanSN" & vbCrLf
    '    sql += " where 1=1" & vbCrLf
    '    sql += " and ot.years='" & labYears.Text & "' " & vbCrLf
    '    sql += " and op.PlanName='" & labPlanName.Text & "' " & vbCrLf
    '    sql += " and DistID='" & sm.UserInfo.DistID & "'" & vbCrLf

    '    dt = DbAccess.GetDataTable(sql)
    '    If dt.Rows.Count = 0 Then
    '        Common.MessageBox(Me, "查無資料")
    '    Else
    '        'ddl_years2.Enabled = False
    '        'ddl_TPlanID2.Enabled = False
    '        'btn_select.Visible = False
    '        'btn_clear.Visible = True
    '        'ddl_TenderName2.Enabled = True
    '        'With ddl_TenderName2
    '        '    .DataSource = dt
    '        '    .DataTextField = "TenderNAME"
    '        '    .DataValueField = "tsn"
    '        '    .DataBind()
    '        '    .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    '        'End With
    '        'ddl_TenderName2.SelectedValue = ""
    '    End If
    'End Sub

    'Private Sub btn_clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ddl_years2.SelectedValue = ""
    '    ddl_years2.Enabled = True
    '    ddl_TPlanID2.SelectedValue = ""
    '    ddl_TPlanID2.Enabled = True
    '    btn_select.Visible = True
    '    btn_clear.Visible = False
    '    ddl_TenderName2.SelectedValue = ""
    '    ddl_TenderName2.Enabled = False
    'End Sub

    'Private Sub ddl_TenderName2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_TenderName2.SelectedIndexChanged
    '    Dim sql As String
    '    Dim dt As DataTable
    '    If ddl_TenderName2.SelectedValue <> "" Then
    '        sql = "select tsn from OB_TenderItem where tsn=" & ddl_TenderName2.SelectedValue
    '        If Me.ViewState("un") = "edit" Then
    '            sql += " and tsn <>" & Me.ViewState("edit_tsn")
    '        End If
    '        dt = DbAccess.GetDataTable(sql)
    '        If dt.Rows.Count > 0 Then
    '            Common.MessageBox(Me, "此標案已建檔!!")
    '        Else
    '            Panel_ListBox.Visible = True
    '        End If
    '    End If
    'End Sub

    Private Sub ddl_ORName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_ORName.SelectedIndexChanged
        lbx_Source.Items.Clear() '可選擇子項用
        lbx_Get.Items.Clear() '選擇後子項用
        txt_score.Text = "" '配分清除

        If Not ddl_ORName.SelectedValue = "" Then
            Dim sql As String
            Dim dt As DataTable
            sql = "select ORSN,ORName from OB_ReviewItem where ORAvail='Y' and ORParent="
            sql += ddl_ORName.SelectedValue & " and DistID='" & sm.UserInfo.DistID & "' order by ORSN"
            dt = DbAccess.GetDataTable(sql)

            With lbx_Source
                .DataSource = dt
                .DataTextField = "ORName"
                .DataValueField = "ORSN"
                .DataBind()
            End With

        End If

    End Sub

    '評選項目確定
    Private Sub btn_ORItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ORItem.Click


        If ddl_ORName.SelectedValue = "" Then
            Common.MessageBox(Me, "請選擇設定【評選大項】")
        ElseIf txt_score.Text = "" Then
            Common.MessageBox(Me, "請填寫【評選項目配分】")
        Else
            Dim dt_ORItem As New DataTable
            If Me.ViewState("tmptable") Is Nothing Then '暫存table
                dt_ORItem.Columns.Add("id")
                dt_ORItem.Columns.Add("data")
                dt_ORItem.Columns.Add("score")
            Else
                dt_ORItem = Me.ViewState("tmptable")
            End If

            '目前編輯的id為空
            If Me.ViewState("Item_Edit_id") = "" Then '新增
                Dim arr() As String

                If dt_ORItem.Rows.Count > 0 Then
                    For i As Int16 = 0 To dt_ORItem.Rows.Count - 1
                        arr = Split(dt_ORItem.Rows(i)("id"), ",")
                        If ddl_ORName.SelectedValue = arr(0) Then
                            Common.MessageBox(Me, "【評選項目】重復")
                            Exit Sub
                        End If
                    Next
                End If

                Dim dr As DataRow
                dr = dt_ORItem.NewRow
                dt_ORItem.Rows.Add(dr)
                dr("id") = ddl_ORName.SelectedValue
                dr("data") = "<STRONG>" + ddl_ORName.SelectedItem.Text + "</STRONG>" + "<BR>"

                '重排評選項目子項
                For j As Int16 = 0 To lbx_Get.Items.Count - 1
                    dr("id") += "," + lbx_Get.Items(j).Value
                    If j = lbx_Get.Items.Count - 1 Then
                        dr("data") += "(" & j + 1 & ")" + lbx_Get.Items(j).Text
                    Else
                        dr("data") += "(" & j + 1 & ")" + lbx_Get.Items(j).Text + "<BR>"
                    End If
                Next
                dr("score") = txt_score.Text

                Me.ViewState("lbl_num") = dt_ORItem.Rows.Count + 1
            Else '修改

                '確定評選項目大項(可不用) 
                If dt_ORItem.Rows.Count > 0 Then
                    Dim arr() As String
                    Dim arr_item() As String
                    arr_item = Split(Me.ViewState("Item_Edit_id"), ",")

                    For i As Int16 = 0 To dt_ORItem.Rows.Count - 1
                        arr = Split(dt_ORItem.Rows(i)("id"), ",")
                        If ddl_ORName.SelectedValue = arr(0) And Not (ddl_ORName.SelectedValue = arr_item(0)) Then
                            Common.MessageBox(Me, "【評選項目】重復")
                            Exit Sub
                        End If
                    Next
                End If

                Dim dt As DataTable = Me.ViewState("tmptable")
                Dim dr As DataRow
                dr = dt.Select("id='" & Me.ViewState("Item_Edit_id") & "'")(0)

                dr("id") = ddl_ORName.SelectedValue
                dr("data") = "<STRONG>" + ddl_ORName.SelectedItem.Text + "</STRONG>" + "<BR>"

                '重排評選項目子項
                For j As Int16 = 0 To lbx_Get.Items.Count - 1
                    dr("id") += "," + lbx_Get.Items(j).Value
                    If j = lbx_Get.Items.Count - 1 Then
                        dr("data") += "(" & j + 1 & ")" + lbx_Get.Items(j).Text
                    Else
                        dr("data") += "(" & j + 1 & ")" + lbx_Get.Items(j).Text + "<BR>"
                    End If
                Next

                dr("score") = txt_score.Text

                '評選項目大項(保留 移除已經新增的項目)
                If dt_ORItem.Rows.Count > 0 Then
                    For i As Int16 = 0 To dt_ORItem.Rows.Count - 1
                        '移除已經新增的項目
                        ddl_ORName.Items.Remove(ddl_ORName.Items.FindByValue(Split(dt_ORItem.Rows(i)("id"), ",")(0)))
                    Next
                End If

                Me.ViewState("lbl_num") = dt.Rows.Count + 1
            End If

            Me.ViewState("tmptable") = dt_ORItem
            dg_ORItem.DataSource = dt_ORItem
            dg_ORItem.DataBind()

            'LabAction.Text = ""
            'lbl_num.Text = Convert.ToInt16(lbl_num.Text) + 1

            lbx_Get.Items.Clear()
            lbx_Source.Items.Clear()
            ddl_ORName.Enabled = True
            LabAction.Text = "新增"
            Me.lbl_num.Text = Me.ViewState("lbl_num")

            txt_score.Text = ""
            Me.ViewState("Item_Edit_id") = "" '目前編輯的id清空
        End If
    End Sub

    Private Sub dg_ORItem_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_ORItem.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1

            Dim btndel2 As Button = e.Item.FindControl("btn_del2")
            btndel2.CommandArgument = drv("id")
            btndel2.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆資料?');"

            Dim btnedit2 As Button = e.Item.FindControl("btn_edit2")
            btnedit2.CommandArgument = drv("id")

        End If
    End Sub

    Private Sub dg_ORItem_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_ORItem.ItemCommand
        Dim sql_Source As String
        Dim sql_Get As String

        Dim dt_Source As DataTable
        Dim dt_Get As DataTable

        Dim dt As DataTable = Me.ViewState("tmptable")
        Dim dr As DataRow

        Dim str_ORSN As String
        Dim arr() As String

        dr = dt.Select("id='" & e.CommandArgument & "'")(0)

        str_ORSN = ""
        Select Case e.CommandName
            Case "edit" '修改評選項目子項
                arr = Split(dr("id"), ",")
                For i As Int16 = 1 To arr.Length - 1 '從1開始0不使用。
                    If str_ORSN <> "" Then str_ORSN &= ","
                    str_ORSN &= arr(i)
                Next

                'lbx_Source
                sql_Source = "select ORSN,ORName from OB_ReviewItem where ORAvail='Y' and ORParent=" & arr(0)
                sql_Source += " and DistID='" & sm.UserInfo.DistID & "'"
                If str_ORSN <> "" Then
                    sql_Source += "and ORSN not in(" & str_ORSN & ")"
                End If
                sql_Source += " order by ORSN"
                dt_Source = DbAccess.GetDataTable(sql_Source)
                lbx_Source.Items.Clear()
                With lbx_Source
                    .DataSource = dt_Source
                    .DataTextField = "ORName"
                    .DataValueField = "ORSN"
                    .DataBind()
                End With

                lbx_Get.Items.Clear()
                If str_ORSN <> "" Then
                    sql_Get = "select ORSN,ORName from OB_ReviewItem where ORAvail='Y' and ORParent=" & arr(0)
                    sql_Get += " and DistID='" & sm.UserInfo.DistID & "'"
                    sql_Get += " and ORSN in(" & str_ORSN & ")"
                    dt_Get = DbAccess.GetDataTable(sql_Get)
                    With lbx_Get
                        .DataSource = dt_Get
                        .DataTextField = "ORName"
                        .DataValueField = "ORSN"
                        .DataBind()
                    End With
                End If

                ORName() '可用評選項用
                ddl_ORName.Enabled = False
                Common.SetListItem(ddl_ORName, arr(0))
                'ddl_ORName.SelectedValue = arr(0)

                LabAction.Text = "修改"
                lbl_num.Text = e.Item.Cells(0).Text
                txt_score.Text = Convert.ToString(dr("score"))

                Me.ViewState("Item_Edit_id") = dr("id") '目前編輯的id 
            Case "del"
                '刪除某一評選項目
                dt.Rows.Remove(dr)
                'For i As Int16 = 0 To dt.Rows.Count - 1
                '    dt.Rows(i)("id") = i + 1
                'Next
                If dt.Rows.Count = 0 Then
                    Me.ViewState("tmptable") = Nothing
                Else
                    Me.ViewState("tmptable") = dt
                End If
                dg_ORItem.DataSource = dt
                dg_ORItem.DataBind()

                Dim dt_ORItem As New DataTable
                If Me.ViewState("tmptable") Is Nothing Then '暫存table
                    dt_ORItem.Columns.Add("id")
                    dt_ORItem.Columns.Add("data")
                    dt_ORItem.Columns.Add("score")
                Else
                    dt_ORItem = Me.ViewState("tmptable")
                End If
                ORName() '可用評選項用
                '評選項目大項(保留 移除已經新增的項目)
                If dt_ORItem.Rows.Count > 0 Then
                    For i As Int16 = 0 To dt_ORItem.Rows.Count - 1
                        '移除已經新增的項目
                        ddl_ORName.Items.Remove(ddl_ORName.Items.FindByValue(Split(dt_ORItem.Rows(i)("id"), ",")(0)))
                    Next
                End If

                Me.ViewState("lbl_num") = dt.Rows.Count + 1
                lbx_Get.Items.Clear()
                lbx_Source.Items.Clear()

                ddl_ORName.Enabled = True
                LabAction.Text = "新增"
                Me.lbl_num.Text = Me.ViewState("lbl_num")

                Me.ViewState("Item_Edit_id") = "" '目前編輯的id清空
        End Select
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        'Dim sqlStr As String
        Dim sql As String
        Dim dt As DataTable = Me.ViewState("tmptable")
        Dim dr As DataRow

        'If ddl_TenderName2.SelectedValue = "" Then
        '    Common.MessageBox(Me, "請選擇【標案名稱】")
        'ElseIf dt Is Nothing Then

        If dt Is Nothing Then
            Common.MessageBox(Me, "【評選項目】尚未建立")
        Else
            '新增動作
            If Me.ViewState("un") = "add" Then
                sql = "" & vbCrLf
                sql += " INSERT INTO OB_TenderItem( tsn " & vbCrLf
                sql += " , CreateAcct, CreateTime, ModifyAcct, ModifyTime) " & vbCrLf
                sql += " VALUES( '" & Me.ViewState("tsn") & "' " & vbCrLf
                sql += " , '" & sm.UserInfo.UserID & "', getdate(), '" & sm.UserInfo.UserID & "', getdate()) " & vbCrLf

                'sql = "insert into OB_TenderItem(tsn,CreateAcct,CreateTime,ModifyAcct,ModifyTime) values("
                'sql += ddl_TenderName2.SelectedValue & ",'" & sm.UserInfo.UserID & "','" & FormatDateTime(Now(), 2)
                'sql += " " & FormatDateTime(Now(), 4) & "','" & sm.UserInfo.UserID & "','" & FormatDateTime(Now(), 2)
                'sql += " " & FormatDateTime(Now(), 4) & "')"
                DbAccess.ExecuteNonQuery(sql)

                'sql = "select TIsn from OB_TenderItem where tsn=" & ddl_TenderName2.SelectedValue
                sql = "select TIsn from OB_TenderItem where tsn=" & Me.ViewState("tsn")
                dr = DbAccess.GetOneRow(sql)

                Dim j As Int16 = 1
                For Each item As DataGridItem In dg_ORItem.Items
                    Dim str_id As String = item.Cells(1).Text
                    Dim arr() As String
                    arr = Split(str_id, ",")
                    For i As Int16 = 0 To arr.Length - 1
                        If i = 0 Then '新增大項
                            sql = "insert into OB_TItemDetail(TIsn,ORSN,Sort1,Score,CreateAcct,CreateTime,ModifyAcct,"
                            sql += "ModifyTime) values(" & dr("TIsn") & "," & arr(i) & "," & j & "," & dt.Rows(j - 1)("score") & ",'"
                            sql += sm.UserInfo.UserID & "',getdate() "
                            sql += ",'" & sm.UserInfo.UserID & "',getdate() )"
                            DbAccess.ExecuteNonQuery(sql)
                            j = j + 1
                        Else '新增小項
                            sql = "insert into OB_TItemDetail(TIsn,ORSN,Sort1,Sort2,CreateAcct,CreateTime,ModifyAcct,"
                            sql += "ModifyTime) values(" & dr("TIsn") & "," & arr(i) & "," & j - 1 & "," & i & ",'"
                            sql += sm.UserInfo.UserID & "',getdate() "
                            sql += ",'" & sm.UserInfo.UserID & "',getdate() )"
                            DbAccess.ExecuteNonQuery(sql)
                        End If
                    Next
                Next

                Dim strScript As String
                strScript = "<script language=""javascript"">" + vbCrLf
                strScript += "if (window.confirm('資料新增成功!\n請問是否繼續新增？')){" + vbCrLf
                strScript += "  location.href='OB_01_006.aspx?un=add';" + vbCrLf
                strScript += "} else {;" + vbCrLf
                strScript += "  location.href='OB_01_006.aspx';" + vbCrLf
                strScript &= "}" & vbCrLf
                strScript &= "</script>"
                TIMS.RegisterStartupScript(Me, "ring", strScript)

            Else
                '非新增動作
                'sql = "DELETE ob_tenderitem where tsn=" & Me.ViewState("edit_tsn")
                'DbAccess.ExecuteNonQuery(sql)
                'sql = "DELETE ob_titemdetail where tisn=" & Me.ViewState("edit_tisn")
                'DbAccess.ExecuteNonQuery(sql)

                sql = "DELETE ob_tenderitem where tsn=" & Me.ViewState("tsn")
                DbAccess.ExecuteNonQuery(sql)
                sql = "DELETE ob_titemdetail where tisn=" & Me.ViewState("TIsn")
                DbAccess.ExecuteNonQuery(sql)

                sql = "" & vbCrLf
                sql += " INSERT INTO OB_TenderItem( tsn " & vbCrLf
                sql += " , CreateAcct, CreateTime, ModifyAcct, ModifyTime) " & vbCrLf
                sql += " VALUES( '" & Me.ViewState("tsn") & "' " & vbCrLf
                sql += " , '" & sm.UserInfo.UserID & "', getdate(), '" & sm.UserInfo.UserID & "', getdate()) " & vbCrLf

                'sql = "insert into OB_TenderItem(tsn,CreateAcct,CreateTime,ModifyAcct,ModifyTime) values("
                'sql += ddl_TenderName2.SelectedValue & ",'" & sm.UserInfo.UserID & "','" & FormatDateTime(Now(), 2)
                'sql += " " & FormatDateTime(Now(), 4) & "','" & sm.UserInfo.UserID & "','" & FormatDateTime(Now(), 2)
                'sql += " " & FormatDateTime(Now(), 4) & "')"

                DbAccess.ExecuteNonQuery(sql)

                sql = "select TIsn from OB_TenderItem where tsn=" & Me.ViewState("tsn") 'ddl_TenderName2.SelectedValue
                dr = DbAccess.GetOneRow(sql)
                Dim j As Int16 = 1
                For Each item As DataGridItem In dg_ORItem.Items
                    Dim str_id As String = item.Cells(1).Text
                    Dim arr() As String
                    arr = Split(str_id, ",")
                    For i As Int16 = 0 To arr.Length - 1
                        If i = 0 Then '新增大項
                            sql = "insert into OB_TItemDetail(TIsn,ORSN,Sort1,Score,CreateAcct,CreateTime,ModifyAcct,"
                            sql += "ModifyTime) values(" & dr("TIsn") & "," & arr(i) & "," & j & "," & dt.Rows(j - 1)("score") & ",'"
                            sql += sm.UserInfo.UserID & "',getdate() "
                            sql += ",'" & sm.UserInfo.UserID & "',getdate() ) "
                            DbAccess.ExecuteNonQuery(sql)
                            j = j + 1
                        Else '新增小項
                            If arr(i) <> "" Then
                                sql = "insert into OB_TItemDetail(TIsn,ORSN,Sort1,Sort2,CreateAcct,CreateTime,ModifyAcct,"
                                sql += "ModifyTime) values(" & dr("TIsn") & "," & arr(i) & "," & j - 1 & "," & i & ",'"
                                sql += sm.UserInfo.UserID & "',getdate() "
                                sql += ",'" & sm.UserInfo.UserID & "',getdate() )"
                                DbAccess.ExecuteNonQuery(sql)
                            End If
                        End If
                    Next
                Next
                Common.MessageBox(Me, "資料修改成功!")
                btn_lev_Click(sender, e)
            End If
        End If
    End Sub

    '離開鍵
    Private Sub btn_lev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev.Click
        If Me.ViewState("un") = "add" Then
            Common.RespWrite(Me, "<script>location.href='OB_01_006.aspx';</script>")
        Else
            Panel_Sch.Visible = True
            Panel_View.Visible = True
            Panel_Add_Edit.Visible = False

            'ddl_years2.Enabled = True
            'ddl_years2.SelectedValue = ""
            'ddl_TPlanID2.Enabled = True
            'ddl_TPlanID2.SelectedValue = ""
            'btn_select.Visible = True
            'btn_clear.Visible = False
            'ddl_TenderName2.Enabled = False
            'ddl_TenderName2.SelectedValue = ""

            Panel_ListBox.Visible = False
            Me.ViewState("un") = ""
            Me.ViewState("tmptable") = Nothing
            dg_ORItem.DataSource = Nothing
            dg_ORItem.DataBind()
            lbx_Source.Items.Clear()
            lbx_Get.Items.Clear()
        End If
    End Sub

    '圖形>|
    Private Sub img_Add_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles img_Add.Click
        Dim i As Int16 = 0
        While i <= lbx_Source.Items.Count - 1
            If lbx_Source.Items(i).Selected Then
                lbx_Get.Items.Add(New ListItem(lbx_Source.Items(i).Text, lbx_Source.Items(i).Value))
                lbx_Source.Items.Remove(lbx_Source.Items(i))
            Else
                i += 1
            End If
        End While
    End Sub
    '圖形>>
    Private Sub img_AddAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles img_AddAll.Click
        If (lbx_Source.Items.Count > 0) Then
            Dim lbx_item As New ListItem
            For Each lbx_item In lbx_Source.Items
                lbx_Get.Items.Add(lbx_item)
            Next
        End If
        lbx_Source.Items.Clear()
    End Sub
    '圖形|<
    Private Sub img_Remove_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles img_Remove.Click
        Dim i As Int16 = 0
        While i <= lbx_Get.Items.Count - 1
            If lbx_Get.Items(i).Selected Then
                lbx_Source.Items.Add(New ListItem(lbx_Get.Items(i).Text, lbx_Get.Items(i).Value))
                lbx_Get.Items.Remove(lbx_Get.Items(i))
                sort()
            Else
                i += 1
            End If
        End While
    End Sub
    '圖形<<
    Private Sub img_RemoveAll_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles img_RemoveAll.Click
        If (lbx_Get.Items.Count > 0) Then
            Dim lbx_item As New ListItem
            For Each lbx_item In lbx_Get.Items
                lbx_Source.Items.Add(lbx_item)
            Next
            sort()
        End If
        lbx_Get.Items.Clear()
    End Sub


    'lbx_Source的item排序
    Sub sort()
        Dim ORSN As String = ""
        Dim lbx_item As New ListItem
        For Each lbx_item In lbx_Source.Items
            ORSN = ORSN + lbx_item.Value + ","
        Next
        ORSN = Mid(ORSN, 1, Len(ORSN) - 1)
        Dim sql As String
        Dim dt As DataTable
        sql = "select ORSN,ORName from OB_ReviewItem where ORAvail='Y' and ORParent=" & ddl_ORName.SelectedValue
        sql += "and DistID='" & sm.UserInfo.DistID & "' and ORSN in(" & ORSN & ") order by ORSN"
        dt = DbAccess.GetDataTable(sql)

        With lbx_Source
            .DataSource = dt
            .DataTextField = "ORName"
            .DataValueField = "ORSN"
            .DataBind()
        End With
    End Sub
End Class
