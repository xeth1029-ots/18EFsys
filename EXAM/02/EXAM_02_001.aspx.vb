Partial Class EXAM_02_001
    Inherits AuthBasePage

    Const Cst_Split_flag As String = vbCrLf
    Const Cst_Split_flag2 As String = ","
    Dim rDistID As String = ""
    Dim ID_ExamType As DataTable
    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    Const Cst_errmsg1 As String = "資料異常，查無資料!!"
    'exam_question
    'exam_answer
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '載入選擇頁
        PageControler1.PageDataGrid = dg_Sch
        PageControler2.PageDataGrid = dg_Sch2

        rDistID = sm.UserInfo.DistID

        If Not IsPostBack Then
            ddl_DistID = TIMS.Get_DistID(ddl_DistID)
            ddl_DistID.Items.Remove(ddl_DistID.Items.FindByValue("000"))

            ddl_DistID.Enabled = True
            If sm.UserInfo.DistID <> "000" Then
                ddl_etid = cls_Exam.Get_ExamTypeParent(ddl_etid, rDistID, 1)

                Common.SetListItem(ddl_DistID, sm.UserInfo.DistID)
                ddl_DistID.Enabled = False
            End If

            '依轄區，確認子類別
            Call SET_ddlcETID(rDistID)

            Dim i_sch_Len As Integer = Len(TIMS.ClearSQM(Request("sch")))
            Dim i_aryid_Len As Integer = Len(TIMS.ClearSQM(Request("aryid")))
            Dim i_DistID_Len As Integer = Len(TIMS.ClearSQM(Request("DistID")))
            Dim i_eqid_Len As Integer = Len(TIMS.ClearSQM(Request("eqid")))
            Dim flag_Can_OK As Boolean = True
            If i_sch_Len = 0 Then flag_Can_OK = False
            If i_aryid_Len = 0 Then flag_Can_OK = False
            If i_DistID_Len = 0 Then flag_Can_OK = False
            If i_eqid_Len = 0 Then flag_Can_OK = False

            Dim Rq_DistID As String = TIMS.ClearSQM(Request("DistID"))
            Me.ViewState("aryid") = TIMS.ClearSQM(Request("aryid"))

            If flag_Can_OK Then
                Common.SetListItem(ddl_DistID, Rq_DistID)
                If Me.ViewState("aryid") <> "" Then
                    Dim arr() As String = Split(Me.ViewState("aryid"), "/")
                    If arr.Length > 1 Then
                        Common.SetListItem(ddl_etid, arr(0))
                        Common.SetListItem(ddl_qtype, arr(1))
                    End If
                End If
                '依轄區，確認子類別
                Call SET_ddlcETID(rDistID)
                Call search()
            End If

        End If
    End Sub

    '依轄區，確認子類別
    Sub SET_ddlcETID(ByVal rDistID As String)
        ddl_cETID.Enabled = False
        If ddl_etid.SelectedValue <> "" Then
            ddl_cETID.Enabled = True
            ddl_cETID = cls_Exam.Get_ExamType(ddl_cETID, rDistID, ddl_etid.SelectedValue)
        Else
            ddl_cETID.Items.Clear()
        End If
    End Sub

    '查詢
    Private Sub btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        Call search()
    End Sub

    Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
        Dim sUrl As String = ""
        sUrl = ""
        sUrl &= "EXAM_02_002.aspx?un=add"
        If sm.UserInfo.DistID <> "000" Then
            sUrl &= "&DistID=" & sm.UserInfo.DistID
        Else
            If Me.ddl_DistID.SelectedValue = "" Then
                Common.MessageBox(Me, "請選擇有效轄區!!")
                Exit Sub
            End If
            sUrl &= "&DistID=" & Me.ddl_DistID.SelectedValue
        End If

        TIMS.Utl_Redirect1(Me, sUrl)
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Panel_Sch.Visible = True
        Panel_View.Visible = True
        Panel_View2.Visible = False
        Call search()
    End Sub

    '查詢
    Sub search()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch) '顯示列數不正確

        Dim sql As String = ""
        sql += " select  CONVERT(varchar, eq.etid) + '/' + eq.qtype + '/' + count(eq.eqid) as aryid" & vbCrLf
        sql += "  	,case when p.name is not null then CONVERT(varchar, p.name) + '_' + CONVERT(varchar, ie.name) " & vbCrLf
        sql += "  	when ie.name is not null then CONVERT(varchar, ie.name) " & vbCrLf
        sql += "  	else '不區分'  end as name " & vbCrLf

        sql += " 	,ISNULL(p.name, ISNULL(ie.name,'不區分')) as pName" & vbCrLf
        sql += " 	,case when p.name is null then '不區分' else CONVERT(varchar, ie.name) end as cName" & vbCrLf

        sql += " 	,case eq.qtype " & vbCrLf
        sql += " 		when '1' then '是非題' " & vbCrLf
        sql += " 		when '2' then '選擇題'" & vbCrLf
        sql += " 		when '3' then '複選題' " & vbCrLf
        sql += " 		when '4' then '問答題' " & vbCrLf
        sql += " 	end as qtype" & vbCrLf
        sql += " 	,count(eq.eqid) total " & vbCrLf
        sql += " ,IE.Avail CAvail" & vbCrLf
        sql += " ,P.Avail PAvail" & vbCrLf
        sql += " from " & vbCrLf
        sql += " 	exam_question eq " & vbCrLf
        sql += " 	left join ID_ExamType ie on ie.etid=eq.etid " & vbCrLf
        sql += "  	left join ID_ExamType p on ie.Parent =p.ETID" & vbCrLf
        sql += " where 1=1 " & vbCrLf
        If sm.UserInfo.DistID <> "000" Then '系統管理者可查全部
            sql += " 	and eq.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        Else
            sql += " 	and eq.DistID='" & ddl_DistID.SelectedValue & "'" & vbCrLf
        End If

        If ddl_cETID.SelectedValue = ddl_etid.SelectedValue Then '子類別值，等同父類別
            If ddl_cETID.SelectedValue <> "" Then
                sql += "    AND ( 1!=1 " & vbCrLf
                sql += "    or eq.etid=" & ddl_cETID.SelectedValue & vbCrLf
                sql += "    or ie.Parent=" & ddl_etid.SelectedValue & vbCrLf
                sql += "    ) " & vbCrLf
            End If
        Else
            If ddl_cETID.SelectedValue <> "" Then '子類別值不為空
                sql += "    and eq.etid=" & ddl_cETID.SelectedValue & vbCrLf
            Else
                If ddl_etid.SelectedValue <> "" Then '父類別值不為空
                    sql += "    and eq.etid=" & ddl_etid.SelectedValue & vbCrLf
                End If
            End If
        End If

        If ddl_qtype.SelectedValue <> "0" Then
            sql += "    and eq.qtype=" & ddl_qtype.SelectedValue & vbCrLf
        End If
        sql += " group by " & vbCrLf
        sql += "  	eq.etid,eq.qtype,p.etid,p.name,ie.name " & vbCrLf
        sql += " ,IE.Avail" & vbCrLf
        sql += " ,P.Avail" & vbCrLf

        sql += " order by 2,3 " & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        msg.Visible = True
        Panel_View.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Visible = False
            Panel_View.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub btn_back2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back2.Click
        Call chg_table(0)
    End Sub

    Protected Sub btnSch_Click(sender As Object, e As EventArgs) Handles btnSch.Click
        Call show_dg_Sch2(Me.ViewState("pName"), Me.ViewState("cName"), Me.ViewState("aryid"))
    End Sub

    '查詢 題庫
    Sub show_dg_Sch2(ByVal pName As String, ByVal cName As String, ByVal aryid As String, Optional ByVal sAction As String = "")

        Dim arr() As String = Split(aryid, "/")

        Panel_Sch.Visible = False
        Panel_View.Visible = False
        Panel_View2.Visible = True

        dg_Sch2.Visible = False
        msg2.Visible = True

        lbl_title1.Visible = True
        lbl_title2.Visible = True
        lbl_title3.Visible = True
        lbl_petidname.Visible = True
        lbl_cetidname.Visible = True

        lbl_qtype.Visible = True
        lbl_math.Visible = True

        Me.ViewState("pName") = pName
        Me.ViewState("cName") = cName
        Me.ViewState("aryid") = aryid

        lbl_petidname.Text = pName 'Exam.Get_ExamTypeName(arr(0)) '不區分"
        lbl_cetidname.Text = cName 'Exam.Get_ExamTypeName(arr(0)) '不區分"
        lbl_qtype.Text = cls_Exam.Get_qtypeName(arr(1))
        If sAction = "del" Then
            arr(2) = arr(2) - 1
            lbl_math.Text = arr(2)
        Else
            lbl_math.Text = arr(2)
        End If
        Me.ViewState("aryid") = arr(0) & "/" & arr(1) & "/" & arr(2)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select " & vbCrLf
        sql += " eq.eqid" & vbCrLf
        sql += " ,eq.question" & vbCrLf
        sql += " ,eq.qtype" & vbCrLf
        sql += " ,case when eq.StopUse ='Y' then '是' else '' end StopUse" & vbCrLf
        sql += " ,ISNULL(CASE WHEN ea.total=1 THEN '－' else CONVERT(varchar, ea.total) end ,'X') total" & vbCrLf
        sql += " from exam_question eq " & vbCrLf
        sql += " left join (" & vbCrLf
        sql += " select eqid ,count(1) total" & vbCrLf
        sql += " from exam_answer " & vbCrLf
        sql += " group by eqid) ea on ea.eqid=eq.eqid " & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and eq.etid='" & arr(0) & "'" & vbCrLf
        sql += " and eq.qtype='" & arr(1) & "'" & vbCrLf

        If sm.UserInfo.DistID <> "000" Then '系統管理者可查全部
            sql += " and eq.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        Else
            sql += " and eq.DistID='" & ddl_DistID.SelectedValue & "'" & vbCrLf
        End If

        If txtQuestion.Text <> "" Then sql += " and eq.question like '%" & txtQuestion.Text & "%' "

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        dg_Sch2.Visible = False
        msg2.Visible = True
        If dt.Rows.Count > 0 Then
            dg_Sch2.Visible = True
            msg2.Visible = False

            PageControler2.PageDataTable = dt
            PageControler2.ControlerLoad()
        End If

    End Sub

    Private Sub dg_Sch_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        'e.CommandArgument
        Dim aryid As String = ""
        Dim pName As String = ""
        Dim cName As String = ""

        If e.CommandName = "view" Then
            txtQuestion.Text = ""

            aryid = TIMS.GetMyValue(e.CommandArgument, "aryid")
            pName = TIMS.GetMyValue(e.CommandArgument, "pName")
            cName = TIMS.GetMyValue(e.CommandArgument, "cName")

            show_dg_Sch2(pName, cName, aryid)
        End If
    End Sub

    Private Sub dg_Sch_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnview As Button = e.Item.FindControl("btn_view")
                'e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                btnview.CommandArgument = ""
                btnview.CommandArgument += "aryid=" & drv("aryid")
                btnview.CommandArgument += "&pName=" & drv("pName")
                btnview.CommandArgument += "&cName=" & drv("cName")

                If Convert.ToString(drv("PAvail")) <> "" AndAlso Convert.ToString(drv("CAvail")) = "1" Then
                    If Convert.ToString(drv("PAvail")) <> Convert.ToString(drv("CAvail")) Then
                        e.Item.Cells(1).Text = "<font color='Red'>" & Convert.ToString(drv("pName")) & "(異常停用)</font>"
                        TIMS.Tooltip(e.Item.Cells(1), "異常停用!!")
                    Else
                        e.Item.Cells(1).Text = "<font color='Silver'>" & Convert.ToString(drv("pName")) & "</font>"
                    End If
                End If

        End Select

    End Sub

    Private Sub dg_Sch2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch2.ItemCommand
        Select Case e.CommandName
            Case "view" '檢視
                Panel_View2.Visible = False
                'Dim sql As String
                'Dim dr As DataRow
                'sql = " select ie.name,eq.question,eq.qtype,eq.stype,count(ea.eqid)total,eq.eqid from exam_question eq join id_examtype"
                'sql += " ie on ie.etid=eq.etid join exam_answer ea on ea.eqid=eq.eqid where 1=1 and eq.eqid=" & e.CommandArgument
                'sql += " group by ie.name,eq.question,eq.qtype,eq.stype,eq.eqid"
                Dim dr As DataRow = GetQuestion(e.CommandArgument)
                If Not dr Is Nothing Then

                    lbl_vpetidname.Text = dr("pname")
                    lbl_vcetidname.Text = dr("cname")
                    lbl_vquestion.Text = dr("question")
                    lbl_vqtype.Text = cls_Exam.Get_qtypeName(dr("qtype"))

                    Select Case Convert.ToString(dr("qtype"))
                        Case "1"
                            Call chg_table(1)
                            show_exam_answer(e.CommandArgument, 1)
                        Case "4"
                            Call chg_table(6)
                            show_exam_answer(e.CommandArgument, 6)
                        Case Else
                            Select Case Convert.ToString(dr("total"))
                                Case "2"
                                    Call chg_table(2)
                                    show_exam_answer(e.CommandArgument, 2)
                                Case "3"
                                    Call chg_table(3)
                                    show_exam_answer(e.CommandArgument, 3)
                                Case "4"
                                    Call chg_table(4)
                                    show_exam_answer(e.CommandArgument, 4)
                                Case "5"
                                    Call chg_table(5)
                                    show_exam_answer(e.CommandArgument, 5)
                                Case Else
                                    Common.MessageBox(Me, Cst_errmsg1)
                                    Exit Sub
                            End Select
                    End Select
                Else
                    Common.MessageBox(Me, Cst_errmsg1)
                    Exit Sub
                End If

            Case "edit" '修改
                ''
                'Common.RespWrite(Me, "<script>location.href='../02/EXAM_02_002.aspx?un=edit&eqid=" & e.CommandArgument & "&aryid=" & Me.ViewState("aryid") & "'</script>")

                Dim sUrl As String = ""
                sUrl = ""
                sUrl &= "EXAM_02_002.aspx?un=edit"
                sUrl &= "&eqid=" & e.CommandArgument
                sUrl &= "&aryid=" & Me.ViewState("aryid")
                If sm.UserInfo.DistID <> "000" Then
                    sUrl &= "&DistID=" & sm.UserInfo.DistID
                Else
                    If Me.ddl_DistID.SelectedValue = "" Then
                        Common.MessageBox(Me, "請選擇有效轄區!!")
                        Exit Sub
                    End If
                    sUrl &= "&DistID=" & Me.ddl_DistID.SelectedValue
                End If
                TIMS.Utl_Redirect1(Me, sUrl)

            Case "del" '刪除
                Dim iEQID As Integer = Val(TIMS.ClearSQM(e.CommandArgument))

                Dim sql As String
                sql = "delete exam_question where 1=1 and eqid=" & iEQID 'e.CommandArgument
                DbAccess.GetDataTable(sql, objconn)

                sql = "delete exam_answer where 1=1 and eqid=" & iEQID 'e.CommandArgument
                DbAccess.GetDataTable(sql, objconn)
                Page.RegisterStartupScript("del", "<script>alert('刪除成功!');</script>")

                show_dg_Sch2(TIMS.ClearSQM(Me.ViewState("pName")), TIMS.ClearSQM(Me.ViewState("cName")), TIMS.ClearSQM(Me.ViewState("aryid")), "del")
                'Dim dr As DataRow = GetQuestion(e.CommandArgument)
                'search()
        End Select
    End Sub

    Private Sub dg_Sch2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch2.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            Dim btnview2 As Button = e.Item.FindControl("btn_view2")
            Dim btnedit2 As Button = e.Item.FindControl("btn_edit2")
            Dim btndel2 As Button = e.Item.FindControl("btn_del2")

            'e.Item.Cells(0).Text = e.Item.ItemIndex + 1
            e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

            If Len(e.Item.Cells(1).Text) > 22 Then
                e.Item.Cells(1).Text = Left((e.Item.Cells(1).Text), 45) & "..."
            End If

            btndel2.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆資料?');"

            btnview2.CommandArgument = drv("eqid")
            btnedit2.CommandArgument = drv("eqid")
            btndel2.CommandArgument = drv("eqid")
        End If
    End Sub

    '查詢該題庫
    Function GetQuestion(ByVal EQID As String) As DataRow
        Dim iEQID As Integer = Val(TIMS.ClearSQM(EQID))
        Dim Sql As String
        Sql = "" & vbCrLf
        Sql += " select case " & vbCrLf
        Sql += "          when p.name is not null then CONVERT(varchar, p.name) + '_' + CONVERT(varchar, ie.name) " & vbCrLf
        Sql += "          when ie.name is not null then CONVERT(varchar, ie.name) " & vbCrLf
        Sql += "          else '不區分'  end as name" & vbCrLf
        Sql += "   	,eq.question ,eq.qtype  ,count(ea.eqid) total ,eq.eqid " & vbCrLf
        Sql += "  	,ISNULL(p.name, ISNULL(ie.name,'不區分')) as pName" & vbCrLf
        Sql += "  	,case when p.name is null then '不區分' else CONVERT(varchar, ie.name) end as cName" & vbCrLf
        Sql += " from exam_question eq " & vbCrLf
        Sql += "   left join id_examtype ie on ie.etid=eq.etid " & vbCrLf
        Sql += "   left join ID_ExamType p on ie.Parent =p.ETID" & vbCrLf
        Sql += "   left join exam_answer ea on ea.eqid=eq.eqid " & vbCrLf
        Sql += " where 1=1 " & vbCrLf
        Sql += " 	and eq.eqid=" & EQID & vbCrLf
        Sql += " group by " & vbCrLf
        Sql += " 	p.name,ie.name,eq.question,eq.qtype ,eq.eqid" & vbCrLf
        Return DbAccess.GetOneRow(Sql, objconn)
    End Function

    '依參數顯示
    Sub chg_table(ByVal parament As Integer)
        'parament 0:不顯示 1:是非 2~5 項目數 6:問答
        Select Case parament
            Case 0 '不顯示
                btn_back2.Visible = False
                Panel_show.Visible = False
                tab_select1.Visible = False
                tab_select2.Visible = False
                tab_select3.Visible = False
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = False
                Panel_View2.Visible = True
            Case 1 '是非題
                btn_back2.Visible = True
                Panel_show.Visible = True
                tab_select1.Visible = True
                tab_select2.Visible = False
                tab_select3.Visible = False
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 2 '項目數二
                btn_back2.Visible = True
                Panel_show.Visible = True
                tab_select1.Visible = False
                tab_select2.Visible = True
                tab_select3.Visible = False
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 3 '項目數三
                btn_back2.Visible = True
                Panel_show.Visible = True
                tab_select1.Visible = False
                tab_select2.Visible = True
                tab_select3.Visible = True
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 4 '項目數四
                btn_back2.Visible = True
                Panel_show.Visible = True
                tab_select1.Visible = False
                tab_select2.Visible = True
                tab_select3.Visible = True
                tab_select4.Visible = True
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 5 '項目數五
                btn_back2.Visible = True
                Panel_show.Visible = True
                tab_select1.Visible = False
                tab_select2.Visible = True
                tab_select3.Visible = True
                tab_select4.Visible = True
                tab_select5.Visible = True
                tab_select6.Visible = False
            Case 6 '問答題
                btn_back2.Visible = True
                Panel_show.Visible = True
                tab_select1.Visible = False
                tab_select2.Visible = False
                tab_select3.Visible = False
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = True
        End Select
    End Sub

    '顯示詳細內容
    Sub show_exam_answer(ByVal iEQID As Int16, ByVal iParament As Int16)
        'parament 1:是非 2~5 項目數 6:問答
        Dim sql As String = ""
        Dim str As String = ""
        Dim arr() As String
        sql = "select answer,isans from exam_answer where 1=1 and eqid= " & iEQID
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, Cst_errmsg1)
            Exit Sub
        End If

        Select Case iParament
            Case 1 'parament   1:是非 2~5 項目數 6:問答
                dr = DbAccess.GetOneRow(sql, objconn)
                If dr("isans") = "Y" Then
                    lbl_ans1.Text = "○"
                Else
                    lbl_ans1.Text = "╳"
                End If
            Case 6 'parament   1:是非 2~5 項目數 6:問答
                dr = DbAccess.GetOneRow(sql, objconn)
                lbl_ans4.Text = dr("answer")
            Case Else
                Const Cst_checkFlag1 As String = "V"

                'parament   1:是非 2~5 項目數 6:問答
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
                For Each dr In dt.Rows
                    str += dr("answer").ToString & Cst_Split_flag
                    str += dr("isans").ToString & Cst_Split_flag
                Next
                arr = Split(str, Cst_Split_flag)
                chg_type("lbl")
                If iParament >= 2 Then
                    lbl_ans2_1.Text = arr(0)
                    lbl_ans2_2.Text = arr(2)
                    lbl_chk2_1.Text = ""
                    If arr(1) = "Y" Then
                        lbl_chk2_1.Text = Cst_checkFlag1
                    End If
                    lbl_chk2_2.Text = ""
                    If arr(3) = "Y" Then
                        lbl_chk2_2.Text = Cst_checkFlag1
                    End If
                End If
                If iParament >= 3 Then
                    lbl_ans2_3.Text = arr(4)
                    lbl_chk2_3.Text = ""
                    If arr(5) = "Y" Then
                        lbl_chk2_3.Text = Cst_checkFlag1
                    End If
                End If
                If iParament >= 4 Then
                    lbl_ans2_4.Text = arr(6)
                    lbl_chk2_4.Text = ""
                    If arr(7) = "Y" Then
                        lbl_chk2_4.Text = Cst_checkFlag1
                    End If
                End If
                If iParament = 5 Then
                    lbl_ans2_5.Text = arr(8)
                    lbl_chk2_5.Text = ""
                    If arr(9) = "Y" Then
                        lbl_chk2_5.Text = Cst_checkFlag1
                    End If
                End If
        End Select
    End Sub

    Sub chg_type(ByVal str As String)
        Select Case str
            Case "lbl"
                lbl_ans2_1.Visible = True
                lbl_ans2_2.Visible = True
                lbl_ans2_3.Visible = True
                lbl_ans2_4.Visible = True
                lbl_ans2_5.Visible = True
                lkb_ans2_1.Visible = False
                lkb_ans2_2.Visible = False
                lkb_ans2_3.Visible = False
                lkb_ans2_4.Visible = False
                lkb_ans2_5.Visible = False
            Case "lkb"
                lkb_ans2_1.Visible = True
                lkb_ans2_2.Visible = True
                lkb_ans2_3.Visible = True
                lkb_ans2_4.Visible = True
                lkb_ans2_5.Visible = True
                lbl_ans2_1.Visible = False
                lbl_ans2_2.Visible = False
                lbl_ans2_3.Visible = False
                lbl_ans2_4.Visible = False
                lbl_ans2_5.Visible = False

        End Select
    End Sub

    Private Sub ddl_etid_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_etid.SelectedIndexChanged
        If sm.UserInfo.DistID <> "000" Then
            rDistID = sm.UserInfo.DistID
        Else
            If Me.ddl_DistID.SelectedValue = "" Then
                Common.MessageBox(Me, "請選擇有效轄區!!")
                Exit Sub
            End If
            rDistID = Me.ddl_DistID.SelectedValue
        End If

        Call SET_ddlcETID(rDistID)
    End Sub

    Private Sub ddl_DistID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_DistID.SelectedIndexChanged
        If sm.UserInfo.DistID <> "000" Then
            rDistID = sm.UserInfo.DistID
        Else
            If Me.ddl_DistID.SelectedValue = "" Then
                Common.MessageBox(Me, "請選擇有效轄區!!")
                Exit Sub
            End If
            rDistID = Me.ddl_DistID.SelectedValue
        End If

        ddl_etid = cls_Exam.Get_ExamTypeParent(ddl_etid, rDistID, 1)
    End Sub

    '檢查(匯入Excel)資料
    Function CheckImportData(ByVal col_dr As Array) As String
        Dim Reason As String = ""

        Const Cst_ETID As Integer = 0
        Const Cst_Question As Integer = 1
        Const Cst_QTYPE As Integer = 2
        Const Cst_ANSWER As Integer = 3
        Const Cst_ISANS As Integer = 4

        Dim ok_ISANS_cnt As Integer = 0
        Dim ok_ANSWER_cnt As Integer = 0
        Dim ok_ANSWER_flag As Boolean = False
        'Const Cst_SplitFlag1 As String = ","

        'ETID	
        'Question問題(題目)	
        'QTYPE型態(1=是非題;2=選擇;3=複選;4=問答;)	
        'ANSWER答案內容(請用半型逗點分隔)(2=選擇;3=複選;)
        'ISANS正確答案(請用半型逗點分隔)(2=選擇;3=複選;)(1=是非題;4=問答;)

        If col_dr.Length < 3 Then
            Reason += "欄位對應有誤<BR>"
        Else
            If Convert.ToString(col_dr(Cst_ETID)) <> "" Then
                If ID_ExamType.Select("ETID='" & Convert.ToString(col_dr(Cst_ETID)) & "'").Length = 0 Then
                    Reason += "類別ID 超過系統範圍，請確認<BR>"
                End If
            Else
                Reason += "類別ID不可為空<BR>"
            End If
            If Convert.ToString(col_dr(Cst_Question)) <> "" Then
            Else
                Reason += "問題(題目) 不可為空<BR>"
            End If
            If Convert.ToString(col_dr(Cst_QTYPE)) <> "" Then
                Select Case Convert.ToString(col_dr(Cst_QTYPE))
                    Case "1", "2", "3", "4"
                    Case Else
                        Reason += "題目型態 只可是 1.2.3.4 (1=是非題;2=選擇;3=複選;4=問答;)<BR>"
                End Select
            Else
                Reason += "題目型態 不可為空<BR>"
            End If

            If Reason = "" Then
                '1=是非題;2=選擇;3=複選;4=問答;
                Select Case Convert.ToString(col_dr(Cst_QTYPE))
                    Case "1"  '1=是非題;2=選擇;3=複選;4=問答;
                        If Convert.ToString(col_dr(Cst_ISANS)) <> "" Then
                            Select Case LCase(Convert.ToString(col_dr(Cst_ISANS)))
                                Case "o", "x"
                                Case Else
                                    Reason += "正確答案 只可以是 o或x<BR>"
                            End Select
                        Else
                            Reason += "[是非題] 正確答案 不可為空<BR>"
                        End If
                    Case "2" '1=是非題;2=選擇;3=複選;4=問答;
                        ok_ANSWER_cnt = 0
                        ok_ANSWER_flag = False
                        If Convert.ToString(col_dr(Cst_ANSWER)) <> "" Then
                            '查看是否有分隔符號
                            If Convert.ToString(col_dr(Cst_ANSWER)).IndexOf(Cst_Split_flag2) = -1 Then
                                Reason += "(選擇) 答案內容 無分隔符號<BR>"
                            Else
                                Dim sANSWER2 As String() = Convert.ToString(col_dr(Cst_ANSWER)).Split(Cst_Split_flag2)
                                For i As Integer = 0 To sANSWER2.Length - 1
                                    If Convert.ToString(sANSWER2(i)).Trim = "" Then
                                        Reason += "(選擇) 答案內容 沒有輸入資料<BR>"
                                        Exit For
                                    End If
                                Next
                                If Reason = "" Then
                                    ok_ANSWER_cnt = sANSWER2.Length()
                                    ok_ANSWER_flag = True
                                End If
                                If ok_ANSWER_cnt > 5 Then
                                    Reason += "(選擇) 答案內容 不可大於 5個項目 (系統限制)<BR>"
                                End If
                            End If
                        Else
                            Reason += "(選擇) 答案內容 不可為空<BR>"
                        End If

                        If Convert.ToString(col_dr(Cst_ISANS)) <> "" Then
                            '查看是否有分隔符號
                            ok_ISANS_cnt = 0
                            If Convert.ToString(col_dr(Cst_ISANS)).IndexOf(Cst_Split_flag2) > -1 Then
                                Reason += "(選擇) 正確答案 有分隔符號<BR>"
                            Else
                                Try
                                    If Not IsNumeric(Convert.ToString(col_dr(Cst_ISANS))) Then
                                        Reason += "(選擇) 正確答案 應該為小於" & ok_ANSWER_cnt & "的數字(1~" & ok_ANSWER_cnt & ")<BR>"
                                    Else
                                        If col_dr(Cst_ISANS) > ok_ANSWER_cnt Then
                                            Reason += "(選擇) 正確答案 應該為小於" & ok_ANSWER_cnt & "的數字(1~" & ok_ANSWER_cnt & ")<BR>"
                                        End If
                                    End If
                                Catch ex As Exception
                                    Reason += "(選擇) 正確答案 應該為小於" & ok_ANSWER_cnt & "的數字(1~" & ok_ANSWER_cnt & ")<BR>"
                                End Try

                            End If
                        Else
                            Reason += "[選擇題] 正確答案 不可為空<BR>"
                        End If

                    Case "3" '1=是非題;2=選擇;3=複選;4=問答;
                        If Convert.ToString(col_dr(Cst_ANSWER)) <> "" Then
                            '查看是否有分隔符號
                            ok_ANSWER_cnt = 0
                            ok_ANSWER_flag = False
                            If Convert.ToString(col_dr(Cst_ANSWER)).IndexOf(Cst_Split_flag2) = -1 Then
                                Reason += "(複選) 答案內容 無分隔符號<BR>"
                            Else
                                Dim sANSWER2 As String() = Convert.ToString(col_dr(Cst_ANSWER)).Split(Cst_Split_flag2)
                                For i As Integer = 0 To sANSWER2.Length - 1
                                    If Convert.ToString(sANSWER2(i)).Trim = "" Then
                                        Reason += "(複選) 答案內容 沒有輸入資料<BR>"
                                        Exit For
                                    End If
                                Next
                                If Reason = "" Then
                                    ok_ANSWER_cnt = sANSWER2.Length()
                                    ok_ANSWER_flag = True
                                End If
                                If ok_ANSWER_cnt > 5 Then
                                    Reason += "(複選) 答案內容 不可大於 5個項目 (系統限制)<BR>"
                                End If
                            End If
                        Else
                            Reason += "(複選) 答案內容 不可為空<BR>"
                        End If

                        If Convert.ToString(col_dr(Cst_ISANS)) <> "" Then
                            '查看是否有分隔符號
                            ok_ISANS_cnt = 0
                            If Convert.ToString(col_dr(Cst_ISANS)).IndexOf(Cst_Split_flag2) = -1 Then
                                Reason += "(複選) 正確答案 無分隔符號<BR>"
                            Else
                                Dim sISANS2 As String() = Convert.ToString(col_dr(Cst_ISANS)).Split(Cst_Split_flag2)
                                For i As Integer = 0 To sISANS2.Length - 1
                                    If Convert.ToString(sISANS2(i)).Trim = "" Then
                                        Reason += "(複選) 正確答案 沒有輸入資料<BR>"
                                        Exit For
                                    Else
                                        Try
                                            If Not IsNumeric(sISANS2(i)) Then
                                                Reason += "(複選) 正確答案 應該為小於" & ok_ANSWER_cnt & "的數字(1~" & ok_ANSWER_cnt & ")<BR>"
                                            Else
                                                If sISANS2(i) > ok_ANSWER_cnt Then
                                                    Reason += "(複選) 正確答案 應該為小於" & ok_ANSWER_cnt & "的數字(1~" & ok_ANSWER_cnt & ")<BR>"
                                                End If
                                            End If
                                        Catch ex As Exception
                                            Reason += "(複選) 正確答案 應該為小於" & ok_ANSWER_cnt & "的數字(1~" & ok_ANSWER_cnt & ")<BR>"
                                        End Try
                                        If Reason <> "" Then
                                            Exit For
                                        End If
                                    End If
                                Next
                                If Reason = "" Then
                                    ok_ISANS_cnt = sISANS2.Length()
                                End If
                                If ok_ISANS_cnt > 4 Then
                                    Reason += "(複選) 正確答案數量 不可大於 4個項目 (系統限制)<BR>"
                                End If
                                If ok_ISANS_cnt >= ok_ANSWER_cnt Then
                                    Reason += "(複選) 正確答案數量 不可大於等於 答案內容數量<BR>"
                                End If

                            End If
                        Else
                            Reason += "[複選題] 正確答案 不可為空<BR>"
                        End If

                    Case "4" '1=是非題;2=選擇;3=複選;4=問答;
                        If Convert.ToString(col_dr(Cst_ISANS)) <> "" Then
                        Else
                            Reason += "[問答題] 正確答案 不可為空<BR>"
                        End If

                End Select
            End If
        End If
        Return Reason
    End Function

    '匯入Excel
    Sub Insert_DataTableXLS(ByVal dt_xls As DataTable)
        'Const Cst_Spage As String = "cp_08_013"
        Const WrongPageUrl As String = "../01/EXAM_01_001_Wrong.aspx"

        Const Cst_ETID As Integer = 0
        Const Cst_Question As Integer = 1
        Const Cst_QTYPE As Integer = 2
        Const Cst_ANSWER As Integer = 3
        Const Cst_ISANS As Integer = 4
        'ETID	
        'Question問題(題目)	
        'QTYPE型態(1=是非題;2=選擇;3=複選;4=問答;)	
        'ANSWER答案內容(請用半型逗點分隔)(2=選擇;3=複選;)
        'ISANS正確答案(請用半型逗點分隔)(2=選擇;3=複選;)(1=是非題;4=問答;)

        '建立sql的津貼連線
        Call TIMS.OpenDbConn(objconn)
        '建立範圍table 
        Dim sql As String = ""
        sql = "select * from ID_ExamType where distid ='" & sm.UserInfo.DistID & "'"
        ID_ExamType = DbAccess.GetDataTable(sql, objconn)


        '檢查重覆資料 '題目內容  '轄區分署(轄區中心)代碼
        Dim strSql As String = "
SELECT COUNT(1) CNT FROM Exam_Question
where Question=@Question and DistID=@DistID and QType=@QType
"
        Dim SEL_COUNT As New SqlCommand(strSql, objconn)

        Dim strSql_into As String = "
INSERT INTO Exam_Question(EQID,ETID,Question,QType,DistID,ModifyAcct,ModifyDate)
VALUES (@EQID,@ETID,@Question,@QType,@DistID,@ModifyAcct,GETDATE())
"
        Dim insert_cmd As SqlCommand = New SqlCommand(strSql_into, objconn)

        Dim strSql_into2 As String = "
INSERT INTO Exam_Answer(EAID,EQID,Answer,IsAns)
VALUES (@EAID,@EQID,@Answer,@IsAns)
"
        Dim insert_cmd2 As New SqlCommand(strSql_into2, objconn)

        '開使處理要匯入的資料
        Dim RowIndex As Integer = 1 '改為Excel位置由1開始再加1等於2
        Dim Reason As String = "" '儲存錯誤的原因
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        Dim drWrong As DataRow
        '建立錯誤資料格式Table
        dtWrong.Columns.Add(New DataColumn("Index")) '序號
        dtWrong.Columns.Add(New DataColumn("Reason")) '問題

        For Each dr As DataRow In dt_xls.Rows
            RowIndex += 1 '改為Excel位置由1開始再加1等於2
            Dim colArray As Array = dr.ItemArray
            Reason = CheckImportData(colArray)

            If Reason <> "" Then
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)

                drWrong("Index") = RowIndex '改為Excel位置由1開始再加1等於2
                drWrong("Reason") = Reason
            Else
                '匯入資料
                Dim iNEW_EQID As Integer = 0
                Dim iSqlCount As Integer = 0
                iSqlCount = 0
                SEL_COUNT.Parameters.Clear()
                SEL_COUNT.Parameters.Add("Question", SqlDbType.NVarChar).Value = Convert.ToString(dr(Cst_Question))
                SEL_COUNT.Parameters.Add("DistID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
                SEL_COUNT.Parameters.Add("QType", SqlDbType.VarChar).Value = Convert.ToString(dr(Cst_QTYPE))  'QType
                iSqlCount = SEL_COUNT.ExecuteScalar()
                If iSqlCount > 0 Then
                    '資料重複
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)

                    drWrong("Index") = RowIndex
                    drWrong("Reason") = "(" & Convert.ToString(dr(Cst_Question)) & ")題目內容 重複!!"
                Else
                    'insert
                    insert_cmd.Parameters.Clear()
                    iNEW_EQID = DbAccess.GetNewId(objconn, "EXAM_QUESTION_EQID_SEQ,EXAM_QUESTION,EQID") 'EQID
                    insert_cmd.Parameters.Add("EQID", SqlDbType.Int).Value = iNEW_EQID 'EQID
                    insert_cmd.Parameters.Add("ETID", SqlDbType.VarChar).Value = Convert.ToString(dr(Cst_ETID)) 'ETID
                    insert_cmd.Parameters.Add("Question", SqlDbType.NVarChar).Value = Convert.ToString(dr(Cst_Question)) ' 
                    insert_cmd.Parameters.Add("QType", SqlDbType.VarChar).Value = Convert.ToString(dr(Cst_QTYPE)) 'QType
                    insert_cmd.Parameters.Add("DistID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
                    insert_cmd.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    insert_cmd.ExecuteNonQuery()

                    'strSql += " SELECT Exam_Question_EQID_SEQ.currval  " & vbCrLf
                    'insert_cmd = New SqlCommand(strSql, objconn)
                    'insert_cmd.Parameters.Clear()
                    'NEW_EQID = Convert.ToString(insert_cmd.ExecuteScalar())

                    'QTYPE型態(1=是非題;2=選擇;3=複選;4=問答;)	
                    Select Case Convert.ToString(dr(Cst_QTYPE))
                        Case "1" 'QTYPE型態(1=是非題;2=選擇;3=複選;4=問答;)	
                            Dim iNEW_EAID As Integer = 0
                            iNEW_EAID = DbAccess.GetNewId(objconn, "EXAM_ANSWER_EAID_SEQ,EXAM_ANSWER,EAID") 'EAID
                            insert_cmd2.Parameters.Clear()
                            insert_cmd2.Parameters.Add("EAID", SqlDbType.Int).Value = iNEW_EAID 'EAID
                            insert_cmd2.Parameters.Add("EQID", SqlDbType.VarChar).Value = iNEW_EQID
                            insert_cmd2.Parameters.Add("Answer", SqlDbType.NVarChar).Value = ""
                            Select Case LCase(Convert.ToString(dr(Cst_ISANS)))
                                Case "o"
                                    insert_cmd2.Parameters.Add("IsAns", SqlDbType.VarChar).Value = "Y"
                                Case Else
                                    insert_cmd2.Parameters.Add("IsAns", SqlDbType.VarChar).Value = "N"
                            End Select
                            insert_cmd2.ExecuteNonQuery()
                        Case "2", "3"
                            Convert.ToString(dr(Cst_ANSWER)).Split(Cst_Split_flag2)
                            Dim sISANS As String = "," & Convert.ToString(dr(Cst_ISANS)) & ","
                            Dim sANSWER2 As String() = Convert.ToString(dr(Cst_ANSWER)).Split(Cst_Split_flag2)
                            For i As Integer = 0 To sANSWER2.Length - 1
                                Dim NEW_EAID As Integer = 0
                                NEW_EAID = DbAccess.GetNewId(objconn, "EXAM_ANSWER_EAID_SEQ,EXAM_ANSWER,EAID") 'EAID
                                If Convert.ToString(sANSWER2(i)).Trim <> "" Then
                                    insert_cmd2.Parameters.Clear()
                                    insert_cmd2.Parameters.Add("EAID", SqlDbType.Int).Value = NEW_EAID 'EAID
                                    insert_cmd2.Parameters.Add("EQID", SqlDbType.VarChar).Value = iNEW_EQID
                                    insert_cmd2.Parameters.Add("Answer", SqlDbType.NVarChar).Value = Convert.ToString(sANSWER2(i)).Trim
                                    If sISANS.IndexOf("," & CStr(i + 1) & ",") > -1 Then
                                        insert_cmd2.Parameters.Add("IsAns", SqlDbType.VarChar).Value = "Y"
                                    Else
                                        insert_cmd2.Parameters.Add("IsAns", SqlDbType.VarChar).Value = "N"
                                    End If
                                    insert_cmd2.ExecuteNonQuery()
                                End If
                            Next
                        Case "4" 'QTYPE型態(1=是非題;2=選擇;3=複選;4=問答;)	
                            Dim iNEW_EAID As Integer = 0
                            iNEW_EAID = DbAccess.GetNewId(objconn, "EXAM_ANSWER_EAID_SEQ,EXAM_ANSWER,EAID") 'EAID
                            insert_cmd2.Parameters.Clear()
                            insert_cmd2.Parameters.Add("EAID", SqlDbType.Int).Value = iNEW_EAID 'EAID
                            insert_cmd2.Parameters.Add("EQID", SqlDbType.VarChar).Value = iNEW_EQID
                            insert_cmd2.Parameters.Add("Answer", SqlDbType.NVarChar).Value = Convert.ToString(dr(Cst_ISANS))
                            insert_cmd2.Parameters.Add("IsAns", SqlDbType.VarChar).Value = "Y"
                            insert_cmd2.ExecuteNonQuery()
                    End Select
                End If

            End If
        Next
        'Call TIMS.CloseDbConn(tConn)

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        Dim explain2 As String = ""
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

        explain2 = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

        If dtWrong.Rows.Count = 0 Then
            If Reason = "" Then
                Common.MessageBox(Me, explain)
            Else
                Common.MessageBox(Me, explain & Reason)
            End If
        Else
            Session("MyWrongTable") = dtWrong '塞入session 
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('" & WrongPageUrl & "','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If
    End Sub

    '匯入Excel(按鈕)
    Private Sub BtnImport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnImport1.Click
        Session("MyWrongTable") = Nothing
        Dim Upload_Path As String = "~/EXAM/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Const Cst_Filetype As String = "xls" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, Cst_Filetype) Then Return
        'Dim MyFile As System.IO.File
        Dim MyFileName As String
        Dim MyFileType As String
        Dim flag As String

        If File1.Value <> "" Then
            If File1.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                Exit Sub
            Else
                '取出檔案名稱
                MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then
                    Common.MessageBox(Me, "檔案類型錯誤!")
                    Exit Sub
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If LCase(MyFileType) = LCase(Cst_Filetype) Then
                        flag = ","
                    Else
                        Common.MessageBox(Me, "檔案類型錯誤，必須為" & UCase(Cst_Filetype) & "檔!")
                        Exit Sub
                    End If
                End If
            End If
            '檢查檔案格式與大小----------   End

            Dim dt_xls As DataTable
            Dim Errmag As String = ""
            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            Dim filePath1 As String = Server.MapPath(Upload_Path & MyFileName)
            '上傳檔案
            File1.PostedFile.SaveAs(filePath1)
            '取得內容
            dt_xls = TIMS.GetDataTable_XlsFile(filePath1, "", Errmag, "問題(題目)")
            '刪除檔案'IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)) 
            TIMS.MyFileDelete(filePath1)

            If Errmag <> "" Then
                Common.MessageBox(Me, Errmag)
                Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝")
                Exit Sub
            End If

            Call Insert_DataTableXLS(dt_xls)
        Else
            '沒有檔案名稱
            Common.MessageBox(Me, "請選擇匯入檔案的路徑!!")
        End If
    End Sub
End Class
