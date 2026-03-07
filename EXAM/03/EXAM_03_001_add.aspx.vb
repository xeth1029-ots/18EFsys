Partial Class EXAM_03_001_add
    Inherits AuthBasePage

    Const Cst_id As Integer = 0
    Const Cst_ecid As Integer = 1
    Const Cst_etid As Integer = 2
    Const Cst_petid As Integer = 3
    Const Cst_pName As Integer = 4
    Const Cst_cName As Integer = 5
    Const Cst_qtype As Integer = 6
    Const Cst_qtype_name As Integer = 7
    Const Cst_num As Integer = 8
    Const Cst_one_score As Integer = 9
    Const Cst_score As Integer = 10
    Const Cst_total_num As Integer = 11

    Dim strMsg As String = ""

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            tab_add.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            Create_TIME() '時間格式代入

            Select Case Request("un")
                Case "add" '新增狀況與繼續新增狀況
                    CreateAdd()
                Case "edit"
                    Exam_Edit_Cmd(Request("OCID"), Request("IsOnline"))
            End Select

        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button5.Attributes("onclick") = "choose_class();"

        btn_dgcrt.Attributes("onclick") = "return dg_create();"
        btn_save.Attributes("onclick") = "return check_asave();"
    End Sub

#Region "NOUSE"
    'Function Get_ID_ExamTYPE(ByVal obj As ListControl, ByVal DistID As String) As ListControl
    '    '甄試類別檔 ddl_CETID
    '    Dim sql As String = ""
    '    Dim dt As DataTable = Nothing
    '    sql = "select etid,name from id_examtype where avail=1 "
    '    If DistID <> "000" Then '系統管理者可查全部
    '        sql += " and DistID='" & DistID & "'"
    '    End If
    '    dt = DbAccess.GetDataTable(sql)

    '    With obj
    '        .DataSource = dt
    '        .DataTextField = "name"
    '        .DataValueField = "etid"
    '        .DataBind()
    '        .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, "0"))
    '    End With

    '    Return obj
    'End Function

    'Sub SetQType(ByVal obj As ListControl, ByVal IsOnline As String)
    '    '設定題目類型
    '    obj.Items.Clear()
    '    With obj
    '        .Items.Insert(0, New ListItem("==請選擇==", "0"))
    '        .Items.Insert(1, New ListItem("是非題", "1"))
    '        .Items.Insert(2, New ListItem("選擇題", "2"))
    '        .Items.Insert(3, New ListItem("複選題", "3"))
    '        If IsOnline = "N" Then
    '            'SetQType(ddl_qtype, dr("isonline"))
    '            '非線上考試，可考問答題
    '            .Items.Insert(4, New ListItem("問答題", "4"))
    '        End If
    '    End With
    'End Sub
#End Region

    Sub Exam_Edit_Cmd(ByVal OCID As String, ByVal IsOnline As String)
        Dim arr() As String
        Dim dt As DataTable
        Dim sql As String = ""
        ''甄試類別檔
        ddl_pETID = cls_Exam.Get_ExamTypeParent(ddl_pETID, sm.UserInfo.DistID, 1)
        SET_ddlcETID()

        sql = "" & vbCrLf
        sql += " select " & vbCrLf
        sql += "  ec.ecid,ec.ocid" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.classcname,cc.cycltype) classcname" & vbCrLf
        sql += "  ,ec.sequence id,ec.qtype" & vbCrLf
        sql += "  ,case CONVERT(varchar, ec.qtype) " & vbCrLf
        sql += "   when '1' then '是非題' " & vbCrLf
        sql += "   when '2' then '選擇題' " & vbCrLf
        sql += "   when '3' then '複選題' " & vbCrLf
        sql += "   when '4' then '問答題' end as qtype_name" & vbCrLf
        sql += "   ,dbo.NVL(p.etid, dbo.NVL(ie.etid,0)) as petid" & vbCrLf
        sql += "   ,ec.etid" & vbCrLf
        sql += "   ,ie.name as etid_name" & vbCrLf
        sql += "   ,dbo.NVL(p.name, dbo.NVL(ie.name,'不區分')) as pName" & vbCrLf
        sql += "   ,case when p.name is null then '不區分' else CONVERT(varchar, ie.name) end as cName" & vbCrLf
        sql += "  	,ec.num" & vbCrLf
        sql += "  	,ec.score/ec.num as One_Score" & vbCrLf
        sql += "  	,ec.score" & vbCrLf
        sql += "  	,ec.isonline" & vbCrLf
        sql += "  	,ec.sorttype" & vbCrLf
        sql += "  	,CONVERT(varchar, cc.examDATE, 111) examDATE" & vbCrLf
        sql += "  	,convert(varchar, ec.examsdate, 108) examSTIME" & vbCrLf
        sql += "  	,convert(varchar, ec.examedate, 108) examETIME" & vbCrLf
        sql += "  	,ec.examtime" & vbCrLf
        sql += "  	, cc.tmid" & vbCrLf
        sql += "  	, '[' + kt.JobID + ']' + kt.JobName JobName" & vbCrLf
        sql += "  	, '[' + kt.TrainID + ']' + kt.TrainName TrainName" & vbCrLf
        'sql += "  	,(select count(eqid) from exam_question where qtype=ec.qtype) total_num " & vbCrLf
        sql += "  	,q2.total_num " & vbCrLf
        sql += " from " & vbCrLf
        sql += "  	exam_classdata ec " & vbCrLf
        sql += "  	join class_classinfo cc on cc.ocid=ec.ocid " & vbCrLf
        sql += "  	join Key_TrainType kt on kt.tmid=cc.tmid " & vbCrLf
        sql += "  	join id_examTYPE ie on ie.etid=ec.etid " & vbCrLf
        sql += " 	left join ID_ExamType p on ie.Parent =p.ETID" & vbCrLf
        sql += " 	left join (select qtype,count(*) total_num from exam_question group by qtype) q2 on q2.qtype =ec.qtype" & vbCrLf
        sql += " where 1=1" & vbCrLf

        sql += " 	AND ec.ocid=" & OCID & " " & vbCrLf
        sql += " 	AND ec.isonline='" & IsOnline & "' " & vbCrLf

        sql += " order by " & vbCrLf
        sql += "  	ec.sequence" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow
            dr = dt.Rows(0)

            '設定題目類型
            cls_Exam.SetQTypeName(ddl_qtype, dr("isonline"))
            If Not dr("isonline") = "N" Then
                '線上考試，可設定時間
                ddl_shour.Enabled = True
                ddl_sminute.Enabled = True
                ddl_ehour.Enabled = True
                ddl_eminute.Enabled = True

                arr = Split(dr("examSTIME"), ":")
                Common.SetListItem(ddl_shour, arr(0))
                Common.SetListItem(ddl_sminute, arr(1))
                arr = Split(dr("examETIME"), ":")
                Common.SetListItem(ddl_ehour, arr(0))
                Common.SetListItem(ddl_eminute, arr(1))
            End If

            If Not IsDBNull(dr("examdate")) Then
                txt_examdate1.Text = dr("examdate")
            Else
                txt_examdate1.Text = "未填寫"
            End If
            Common.SetListItem(rbl_isonline, dr("isonline"))
            'rbl_isonline.SelectedValue = dr("isonline")
            If Convert.ToString(dr("sorttype")) <> "" Then
                Common.SetListItem(rblSortType, dr("sorttype"))
            End If
            txt_examtime.Text = dr("examtime")
            If Not IsDBNull(dr("trainname")) Then
                TMID1.Text = dr("trainname")
            Else
                TMID1.Text = dr("jobname")
            End If
            OCID1.Text = dr("classcname")
            OCIDValue1.Value = dr("ocid")
            TMIDValue1.Value = dr("TMID")

            Me.ViewState("oldtable") = dt
            Me.ViewState("tmptable") = dt

            dg_view.DataSource = dt
            dg_view.DataBind()

            '檢查題數是否大於題庫數 傳入  Me.ViewState("tmptable") 傳出  Me.ViewState("flag")
            'check_num()
            '檢查題目類型是否重複控制 & 總合計 設定  lbl_score.Visible lbl_total.Visible
            check_dgview(dt)

            Button8.Visible = False '不可選擇轄區
            Button5.Visible = False '不可選擇班級
            rbl_isonline.Enabled = False '不可修改
            TIMS.Tooltip(rbl_isonline, "不可修改", True)

            tab_add.Visible = True 'panel_add_edit
            Me.ViewState("un") = "edit" '目前功能為 edit
        End If

    End Sub

    Sub CreateAdd()
        tab_add.Visible = True
        '設定題目類型
        Call cls_Exam.SetQTypeName(ddl_qtype, rbl_isonline.SelectedValue)
        ''甄試類別檔
        ddl_pETID = cls_Exam.Get_ExamTypeParent(ddl_pETID, sm.UserInfo.DistID, 1)
        SET_ddlcETID()

        txt_num.Text = ""
        txt_score.Text = ""
        'Me.ViewState("flag") = 0 '
        Me.ViewState("un") = "add" '目前功能為 add
    End Sub

    Private Sub rbl_isonline_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbl_isonline.SelectedIndexChanged
        cls_Exam.SetQTypeName(ddl_qtype, rbl_isonline.SelectedValue)
        If rbl_isonline.SelectedValue = "N" Then
            'txt_examdate1.Text = ""
            ddl_shour.SelectedValue = "0"
            ddl_sminute.SelectedValue = "0"
            ddl_ehour.SelectedValue = "0"
            ddl_eminute.SelectedValue = "0"

            ddl_shour.Enabled = False
            ddl_sminute.Enabled = False
            ddl_ehour.Enabled = False
            ddl_eminute.Enabled = False
        Else
            Common.SetListItem(ddl_shour, "00")
            Common.SetListItem(ddl_sminute, "00")
            Common.SetListItem(ddl_ehour, "23")
            Common.SetListItem(ddl_eminute, "59")

            ddl_shour.Enabled = True
            ddl_sminute.Enabled = True
            ddl_ehour.Enabled = True
            ddl_eminute.Enabled = True
        End If
    End Sub

    '輸入
    Private Sub btn_dgcrt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_dgcrt.Click
        Dim dt As New DataTable
        'Dim dr_num As DataRow
        Dim dr As DataRow
        Dim i As Int16
        Dim arr() As String
        'Dim sql As String

        If Me.ViewState("tmptable") Is Nothing Then
            dt.Columns.Add("id")
            dt.Columns.Add("ecid")

            dt.Columns.Add("petid")
            dt.Columns.Add("etid")
            dt.Columns.Add("pName")
            dt.Columns.Add("cName")

            dt.Columns.Add("qtype")
            dt.Columns.Add("qtype_name")
            dt.Columns.Add("num")
            dt.Columns.Add("one_score")
            dt.Columns.Add("score")
            dt.Columns.Add("total_num") '題庫數
        Else
            dt = Me.ViewState("tmptable")
        End If

        '判斷是否有該題目類型
        If check_question(Int(txt_num.Text), ddl_cETID.SelectedValue, _
                        ddl_qtype.SelectedValue, _
                        sm.UserInfo.DistID, _
                        ddl_pETID.SelectedValue, _
                        strMsg) = "" Then
            'Common.MessageBox(Me, "(意外狀況) 題組類別或題目類型不足，請檢視此資料正確性!")
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If

        If hid_chkedit.Value = "" Then '新增dr

            'start 檢查題目類型是否重複 
            arr = Split(hid_qtype.Value, ",")

            For i = 0 To arr.Length - 2
                If arr(i) = ddl_qtype.SelectedValue And arr(i + 1) = ddl_cETID.SelectedValue Then
                    Common.MessageBox(Me, "此題目類型重複!")
                    Exit Sub
                End If
            Next
            'end 檢查題目類型是否重複

            'start 將資料代入dr
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("id") = dt.Rows.Count '順序
            dr("qtype") = ddl_qtype.SelectedValue 'qtype
            dr("qtype_name") = ddl_qtype.SelectedItem.Text

            'Exam.Get_qtypeName(dr("qtype"))
            'Select Case ddl_qtype.SelectedValue '題目類別
            '    Case "1"
            '        dr("qtype_name") = "是非題"
            '    Case "2"
            '        dr("qtype_name") = "選擇題"
            '    Case "3"
            '        dr("qtype_name") = "複選題"
            '    Case "4"
            '        dr("qtype_name") = "問答題"
            'End Select

            dr("petid") = ddl_pETID.SelectedValue 'petid
            dr("etid") = ddl_cETID.SelectedValue 'etid
            dr("pName") = ddl_pETID.SelectedItem.Text '題目類型
            dr("cName") = ddl_cETID.SelectedItem.Text '題目子類型

            dr("num") = Int(txt_num.Text) '題數
            dr("one_score") = Int(Mid(Int(txt_score.Text) / Int(txt_num.Text), 1, 5)) '每題配分
            dr("score") = Int(txt_score.Text) '總配分

            'If ddl_cETID.SelectedValue <> ddl_pETID.SelectedValue Then
            '    sql = "" & vbCrLf
            '    sql += " select count(eq.eqid) total_num " & vbCrLf
            '    sql += " from exam_question eq" & vbCrLf
            '    sql += " join id_examtype ie on ie.etid=eq.etid " & vbCrLf
            '    sql += " where eq.qtype='" & ddl_qtype.SelectedValue & "'" & vbCrLf
            '    sql += " and ie.etid=" & ddl_cETID.SelectedValue & "" & vbCrLf
            'Else
            '    sql = "" & vbCrLf
            '    sql += " select count(eq.eqid) total_num " & vbCrLf
            '    sql += " from exam_question eq" & vbCrLf
            '    sql += " join id_examtype ie on ie.etid=eq.etid " & vbCrLf
            '    sql += " where eq.qtype='" & ddl_qtype.SelectedValue & "'" & vbCrLf
            '    sql += " and (ie.etid=" & ddl_cETID.SelectedValue & " or ie.parent=" & ddl_pETID.SelectedValue & ")" & vbCrLf
            'End If
            'dr_num = DbAccess.GetOneRow(sql)
            'dr("total_num") = dr_num("total_num") '題庫數
            dr("total_num") = show_total_num(ddl_cETID.SelectedValue, ddl_pETID.SelectedValue, ddl_qtype.SelectedValue, sm.UserInfo.DistID)    '題庫數
            'end 將資料代入dr

        Else '修改dr

            'start 檢查題目類型是否重複 
            arr = Split(hid_qtype.Value, ",")
            For i = 0 To arr.Length - 2
                If arr(i) = ddl_qtype.SelectedValue And arr(i + 1) = ddl_cETID.SelectedValue And dt.Rows(hid_chkedit.Value)("qtype") <> arr(i) Then
                    Common.MessageBox(Me, "此題目類型重複!")
                    Exit Sub
                End If
            Next
            'end 檢查題目類型是否重複

            'start 取代原本dr

            dt.Rows(hid_chkedit.Value)("qtype") = ddl_qtype.SelectedValue 'qtype
            dt.Rows(hid_chkedit.Value)("qtype_name") = ddl_qtype.SelectedItem.Text
            dt.Rows(hid_chkedit.Value)("petid") = ddl_pETID.SelectedValue 'petid
            dt.Rows(hid_chkedit.Value)("etid") = ddl_cETID.SelectedValue 'etid
            dt.Rows(hid_chkedit.Value)("pName") = ddl_pETID.SelectedItem.Text '題目類型
            dt.Rows(hid_chkedit.Value)("cName") = ddl_cETID.SelectedItem.Text '題目子類型

            'dt.Rows(hid_chkedit.Value)("etid") = Int(ddl_cETID.SelectedValue)
            'dt.Rows(hid_chkedit.Value)("etid_name") = ddl_cETID.SelectedItem.Text
            'dt.Rows(hid_chkedit.Value)("qtype") = ddl_qtype.SelectedValue
            'Select Case ddl_qtype.SelectedValue
            '    Case "1"
            '        dt.Rows(hid_chkedit.Value)("qtype_name") = "是非題"
            '    Case "2"
            '        dt.Rows(hid_chkedit.Value)("qtype_name") = "選擇題"
            '    Case "3"
            '        dt.Rows(hid_chkedit.Value)("qtype_name") = "複選題"
            '    Case "4"
            '        dt.Rows(hid_chkedit.Value)("qtype_name") = "問答題"
            'End Select
            dt.Rows(hid_chkedit.Value)("num") = Int(txt_num.Text) '題數
            dt.Rows(hid_chkedit.Value)("one_score") = Int(Mid(Int(txt_score.Text) / Int(txt_num.Text), 1, 5)) '每題配分
            dt.Rows(hid_chkedit.Value)("score") = Int(txt_score.Text) '總配分

            'sql = "select count(eqid) total_num from exam_question where etid=" & ddl_cETID.SelectedValue & " and qtype="
            'sql += ddl_qtype.SelectedValue
            'dr_num = DbAccess.GetOneRow(sql)
            'dt.Rows(hid_chkedit.Value)("total_num") = dr_num("total_num") '題庫數
            dt.Rows(hid_chkedit.Value)("total_num") = show_total_num(ddl_cETID.SelectedValue, ddl_pETID.SelectedValue, ddl_qtype.SelectedValue, sm.UserInfo.DistID)    '題庫數
            hid_chkedit.Value = ""
            'end 取代原本dr
        End If

        Me.ViewState("tmptable") = dt
        dg_view.DataSource = dt
        dg_view.DataBind()

        '檢查題目類型是否重複控制 & 總合計 設定  lbl_score.Visible lbl_total.Visible
        check_dgview(dt)

        'Me.ViewState("flag") = 1 '原為 0 設為 1 表示有輸入資料
        '檢查題數是否大於題庫數 傳入  Me.ViewState("tmptable") 傳出  正常@True 異常@False 
        If Not Check_Num(dt) Then
            'Me.ViewState("flag") = -1
            Common.RespWrite(Me, "<script>alert('【題庫數】小於設定【題數】\n請修改【題數】或【新增題庫】內容')</script>")
        End If
        '確認目前配分可取小數第一位 傳出值 正常@True 異常@False
        If Not Check_Score(dt) Then
            'Me.ViewState("flag") = -2
            Common.RespWrite(Me, "<script>alert('【每題配分】不為整數\n請修改【題數】或【總配分】內容')</script>")
        End If

        Common.SetListItem(ddl_qtype, "0")
        'Common.SetListItem(ddl_cETID, "0")
        ''甄試類別檔
        ddl_pETID = cls_Exam.Get_ExamTypeParent(ddl_pETID, sm.UserInfo.DistID, 1)
        SET_ddlcETID()

        txt_num.Text = ""
        txt_score.Text = ""
    End Sub

    Private Sub dg_view_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_view.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            Dim btnedit As Button = e.Item.FindControl("btn_aedit")
            Dim btndel As Button = e.Item.FindControl("btn_adel")

            btndel.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(Cst_id).Text & " 筆資料?');"
            btnedit.CommandArgument = drv("id")
            btndel.CommandArgument = drv("id")
            'If Int(drv("num")) > Int(drv("total_num")) Then
            '    e.Item.Cells(Cst_num).ForeColor = Color.Red
            'Else
            '    e.Item.Cells(Cst_num).ForeColor = Color.Black
            'End If
            Dim one_score_value As Double = Int(drv("score")) / Int(drv("num"))

            If Int(drv("score")) / Int(drv("num")) * 10 Mod 10 = 0 Then
                e.Item.Cells(Cst_one_score).ForeColor = Color.Black
            Else
                e.Item.Cells(Cst_one_score).Text = String.Format(one_score_value.ToString("00.00"))
                e.Item.Cells(Cst_one_score).ForeColor = Color.Red
                TIMS.Tooltip(e.Item.Cells(Cst_one_score), "(意外狀況) 題目分配分數不為整數 ，請檢視此資料正確性並重新建立!")
            End If

            '判斷是否有該題目類型
            If check_question(drv("num"), drv("etid"), drv("qtype"), sm.UserInfo.DistID, _
                            drv("petid"), strMsg) = "" Then
                ''Common.MessageBox(Me, "(意外狀況) 題組類別或題目類型不足，請檢視此資料正確性!")
                'Common.MessageBox(Me, strMsg)
                'Exit Sub
                e.Item.Cells(Cst_pName).ForeColor = Color.Red
                e.Item.Cells(Cst_cName).ForeColor = Color.Red
                e.Item.Cells(Cst_qtype_name).ForeColor = Color.Red
                TIMS.Tooltip(e.Item, "(意外狀況) 題組類別或題目類型不足，請檢視此資料正確性並重新建立!")
            End If

            If drv("etid") = drv("petid") Then
                e.Item.Cells(Cst_cName).Text = show_question_CName(drv("petid"), drv("qtype"), sm.UserInfo.DistID, "cname")
                e.Item.Cells(Cst_total_num).Text = show_question_CName(drv("petid"), drv("qtype"), sm.UserInfo.DistID, "totalnum")
            Else
                e.Item.Cells(Cst_total_num).Text = show_total_num(drv("etid"), drv("petid"), drv("qtype"), sm.UserInfo.DistID)     '題庫數
                If Int(drv("num")) > Int(e.Item.Cells(Cst_total_num).Text) Then
                    e.Item.Cells(Cst_num).ForeColor = Color.Red
                    TIMS.Tooltip(e.Item.Cells(Cst_num), "(意外狀況) 題目數量不足，請檢視此資料正確性並重新建立!")
                Else
                    e.Item.Cells(Cst_num).ForeColor = Color.Black
                End If
            End If


            'If check_question(drv("num"), drv("etid"), drv("qtype"), sm.UserInfo.DistID) = "" Then
            '    e.Item.Cells(Cst_pName).ForeColor = Color.Red
            '    e.Item.Cells(Cst_cName).ForeColor = Color.Red
            '    e.Item.Cells(Cst_qtype_name).ForeColor = Color.Red
            '    TIMS.Tooltip(e.Item, "(意外狀況) 題組類別或題目類型不足，請檢視此資料正確性並重新建立!")
            'End If
        End If
    End Sub

    Private Sub dg_view_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_view.ItemCommand
        Dim dt As DataTable
        Dim dr As DataRow = Nothing
        Dim i As Int16

        Const cst_errMsg1 As String = "(意外狀況) 傳遞參數有誤，請檢視資料正確性並重新操作!!"
        If Me.ViewState("tmptable") Is Nothing Then
            Common.MessageBox(Me, cst_errMsg1)
            Exit Sub
        End If
        dt = Me.ViewState("tmptable")
        If dt Is Nothing Then
            Common.MessageBox(Me, cst_errMsg1)
            Exit Sub
        End If
        If dt.Select("id=" & e.CommandArgument).Length > 0 Then
            dr = dt.Select("id=" & e.CommandArgument)(0)
        End If
        If dr Is Nothing Then
            Common.MessageBox(Me, cst_errMsg1)
            Exit Sub
        End If

        Select Case e.CommandName
            Case "edit"
                'ddl_qtype.SelectedValue = dr("qtype")
                'ddl_CETID.SelectedValue = dr("etid")
                Common.SetListItem(ddl_qtype, dr("qtype"))

                Common.SetListItem(ddl_pETID, dr("petid"))
                SET_ddlcETID(dr("etid"))

                txt_num.Text = dr("num")
                txt_score.Text = dr("score")
                hid_chkedit.Value = e.CommandArgument - 1
            Case "del"
                dt.Rows.Remove(dr)
                For i = 0 To dt.Rows.Count - 1
                    dt.Rows(i)("id") = i + 1
                Next
                Me.ViewState("tmptable") = dt
                dg_view.DataSource = dt
                dg_view.DataBind()

                '檢查題數是否大於題庫數 傳入  Me.ViewState("tmptable") 傳出  正常@True 異常@False 
                If Not Check_Num(dt) Then
                    'Me.ViewState("flag") = -1
                End If

                '檢查題目類型是否重複控制 & 總合計 設定  lbl_score.Visible lbl_total.Visible
                check_dgview(dt)
        End Select
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        Dim dt As DataTable = Me.ViewState("tmptable")
        If dg_view.Items.Count = 0 Then
            Common.MessageBox(Me, "題組類別未設定無法儲存!!")
            Exit Sub
        End If
        If Me.ViewState("tmptable") Is Nothing Then
            Common.MessageBox(Me, "題組類別未設定無法儲存!!")
            Exit Sub
        End If

        'Me.ViewState("flag") = 1 '原為 0 設為 1 表示有輸入資料
        '檢查題數是否大於題庫數 傳入  Me.ViewState("tmptable") 傳出  正常@True 異常@False 
        If Not Check_Num(dt) Then
            'Me.ViewState("flag") = -1
            Common.RespWrite(Me, "<script>alert('【題庫數】小於設定【題數】\n請修改【題數】或【新增題庫】內容')</script>")
        End If
        '確認目前配分可取小數第一位 傳出值 正常@True 異常@False
        If Not Check_Score(dt) Then
            'Me.ViewState("flag") = -2
            Common.RespWrite(Me, "<script>alert('【每題配分】不為整數\n請修改【題數】或【總配分】內容')</script>")
        End If
        'If Me.ViewState("flag") = -1 Or Me.ViewState("flag") = -2 Then
        '    Common.MessageBox(Me, "【試卷內容】填寫內容有誤!")
        '    Exit Sub
        'End If

        Dim sql As String = ""
        Dim sql_check As String = ""
        If Not IsDate(txt_examdate1.Text) Then
            txt_examdate1.Text = Common.FormatDate(Now)
        End If
        Me.ViewState("examSTIME") = txt_examdate1.Text & " " & ddl_shour.SelectedValue & ":" & ddl_sminute.SelectedValue
        Me.ViewState("examETIME") = txt_examdate1.Text & " " & ddl_ehour.SelectedValue & ":" & ddl_eminute.SelectedValue

        Select Case Me.ViewState("un")
            Case "add" '新增資料
                '1.check exam_classdata
                sql_check = ""
                sql_check += "select distinct ocid from exam_classdata where ocid=" & OCIDValue1.Value
                sql_check += " and isonline='" & rbl_isonline.SelectedValue & "'"
                If TIMS.Get_SQLRecordCount(sql_check) > 0 Then '判斷資料是否重複
                    Common.MessageBox(Me, "此班級題庫重複!")
                    Exit Sub
                End If

            Case "edit" '修改資料
                'sql_check = ""
                'sql_check += "select distinct ocid from exam_classdata where ocid=" & OCIDValue1.Value
                'sql_check += " and isonline='" & rbl_isonline.SelectedValue & "'"
                'If TIMS.Get_SQLRecordCount(sql_check) > 0 Then '判斷資料是否重複
                '    Common.MessageBox(Me, "此班級題庫重複!")
                '    Exit Sub
                'End If

                '1.delete exam_classquestion
                sql = " delete exam_classquestion "
                sql += " where ecid in (select ecid from exam_classdata "
                sql += "    where ocid=" & OCIDValue1.Value & " "
                sql += "    and isonline='" & rbl_isonline.SelectedValue & "'"
                sql += " )"
                DbAccess.ExecuteNonQuery(sql, objconn)

                '2.delete exam_classdata
                sql = " delete exam_classdata "
                sql += " where 1=1 "
                sql += " and ocid=" & OCIDValue1.Value & " "
                sql += " and isonline='" & rbl_isonline.SelectedValue & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)

        End Select

        Select Case Me.ViewState("un")
            Case "add", "edit"  '新增資料'修改資料
                'start 將dg_view資料寫入DB
                'Dim dt As DataTable = Me.ViewState("tmptable")
                Dim dr As DataRow
                For Each item As DataGridItem In dg_view.Items
                    Dim str_id As String = item.Cells(Cst_id).Text
                    Dim int_etid As Int16 = Int(item.Cells(Cst_etid).Text)
                    Dim int_petid As Int16 = Int(item.Cells(Cst_petid).Text)
                    Dim str_qtype As String = item.Cells(Cst_qtype).Text
                    Dim int_num As Int16 = Int(item.Cells(Cst_num).Text)
                    Dim int_one_score As Int16 = Int(item.Cells(Cst_one_score).Text)
                    Dim int_score As Int16 = Int(item.Cells(Cst_score).Text)

                    'Exam_ClassData甄試班級檔
                    sql = "insert into exam_classdata(etid "
                    sql += " ,ocid,sequence,qtype,num,score,isonline,avail,examtime"
                    sql += " ,examsdate ,examedate"
                    sql += " ,sortType"
                    sql += " ,modifyacct,modifydate) "
                    sql += " values(" & int_etid
                    sql += " ," & OCIDValue1.Value
                    sql += " ,'" & str_id & "'," & str_qtype & "," & int_num & "," & int_score
                    sql += " ,'" & rbl_isonline.SelectedValue & "','0', " & Int(txt_examtime.Text) & " "
                    If rbl_isonline.SelectedValue = "N" Then
                        sql += " ,convert(datetime, '" & txt_examdate1.Text & "', 111)"
                        sql += " , NULL "
                    Else
                        sql += " ,'" & Me.ViewState("examSTIME") & "'"
                        sql += " ,'" & Me.ViewState("examETIME") & "'"
                    End If
                    sql += " ,'" & rblSortType.SelectedValue & "' "
                    sql += " ,'" & sm.UserInfo.UserID & "',getdate()) "
                    DbAccess.ExecuteNonQuery(sql, objconn)

                    sql = "select ecid ecid from exam_classdata where etid=" & int_etid
                    sql += " and ocid=" & OCIDValue1.Value & " and isonline='" & rbl_isonline.SelectedValue
                    sql += "' and sequence='" & str_id & "'"
                    dr = DbAccess.GetOneRow(sql, objconn)

                    '產生亂數題庫eqid 傳出 strnum
                    'Exam_ClassQuestion甄試班級題庫檔
                    Me.ViewState("num") = check_question(int_num, int_etid, str_qtype, sm.UserInfo.DistID, int_petid)
                    Dim arr() As String
                    arr = Split(Me.ViewState("num"), ",")
                    For i As Int16 = 0 To int_num - 1
                        sql = " insert into exam_classquestion(ocid,ecid,eqid,score,modifyacct,modifydate) values("
                        sql += OCIDValue1.Value & "," & dr("ecid") & "," & arr(i) & "," & int_one_score & ",'sys'"
                        sql += " ,getdate())"
                        DbAccess.ExecuteNonQuery(sql, objconn)
                    Next
                    Me.ViewState("num") = ""
                Next
                'end 將dg_view資料寫入DB

                Select Case rblSortType.SelectedValue
                    Case "2" '2:亂數
                        sql = "" & vbCrLf
                        sql += " select *" & vbCrLf
                        sql += " from exam_classquestion " & vbCrLf
                        sql += " where 1=1" & vbCrLf
                        sql += " and ocid=" & OCIDValue1.Value & vbCrLf
                        sql += " ORDER BY dbms_random.value " & vbCrLf
                        Dim da As SqlDataAdapter = Nothing
                        Dim dt2 As DataTable
                        dt2 = DbAccess.GetDataTable(sql, da, objconn)
                        Dim i_sort As Integer = 0
                        For Each dr2 As DataRow In dt2.Rows
                            i_sort += 1
                            dr2("sort") = i_sort
                        Next
                        DbAccess.UpdateDataTable(dt2, da)
                    Case Else '1:依順序
                        sql = "" & vbCrLf
                        sql += " select *" & vbCrLf
                        sql += " from exam_classquestion " & vbCrLf
                        sql += " where 1=1" & vbCrLf
                        sql += " and ocid=" & OCIDValue1.Value & vbCrLf
                        Dim da As SqlDataAdapter = Nothing
                        Dim dt2 As DataTable
                        dt2 = DbAccess.GetDataTable(sql, da, objconn)
                        Dim i_sort As Integer = 0
                        For Each dr2 As DataRow In dt2.Rows
                            i_sort += 1
                            dr2("sort") = i_sort
                        Next
                        DbAccess.UpdateDataTable(dt2, da)
                End Select
        End Select

        Select Case Me.ViewState("un")
            Case "add" '新增資料
                Dim strScript As String
                strScript = "<script language=""javascript"">" + vbCrLf
                strScript += "if (window.confirm('資料新增成功!\n請問是否繼續新增？')){" + vbCrLf
                strScript += "  location.href='exam_03_001.aspx?un=add';" + vbCrLf
                strScript += "} else {;" + vbCrLf
                strScript += "  location.href='exam_03_001.aspx';" + vbCrLf
                strScript &= "}" & vbCrLf
                strScript &= "</script>"
                Page.RegisterStartupScript("ring", strScript)

            Case "edit" '修改資料
                Common.MessageBox(Me, "資料修改成功!")
                btn_lev_Click(sender, e)

        End Select
    End Sub

    Private Sub btn_lev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev.Click
        'tab_add.Visible = False
        'lbl_total.Visible = False
        'lbl_score.Visible = False
        'rbl_isonline.Enabled = True
        ''rbl_isonline.SelectedValue = "N"
        'Common.SetListItem(rbl_isonline, "N")
        'Me.ViewState("flag") = 0
        'Me.ViewState("tmptable") = Nothing
        'dg_view.DataSource = Nothing
        'dg_view.DataBind()
        ''ddl_qtype.SelectedValue = "0"
        'Common.SetListItem(ddl_qtype, "0")
        'txt_num.Text = ""
        'txt_score.Text = ""
        'txt_examtime.Text = ""
        'hid_qtype.Value = ""
        'txt_examdate1.Text = ""
        'ddl_shour.Enabled = False
        'ddl_ehour.Enabled = False
        'ddl_sminute.Enabled = False
        'ddl_eminute.Enabled = False

        ''ddl_shour.SelectedValue = "0"
        ''ddl_sminute.SelectedValue = "0"
        ''ddl_ehour.SelectedValue = "0"
        ''ddl_eminute.SelectedValue = "0"
        'Common.SetListItem(ddl_shour, "00")
        'Common.SetListItem(ddl_sminute, "00")
        'Common.SetListItem(ddl_ehour, "23")
        'Common.SetListItem(ddl_eminute, "59")

        'search()
        ''If Me.ViewState("un") = "add" Then
        '    Common.RespWrite(Me, "<script>location.href='exam_03_001.aspx';</script>")
        ''End If
        TIMS.Utl_Redirect1(Me, "exam_03_001.aspx")
    End Sub

    '時間格式代入
    Sub Create_TIME()
        For i As Int16 = 0 To 23
            Dim str As String = i.ToString
            If Len(str) = 1 Then
                str = "0" + str
            End If
            ddl_shour.Items.Insert(str, New ListItem(str, str))
            ddl_ehour.Items.Insert(str, New ListItem(str, str))
        Next
        ddl_shour.Items.Insert(0, New ListItem("--請選擇--", "0"))
        ddl_ehour.Items.Insert(0, New ListItem("--請選擇--", "0"))

        For i As Int16 = 0 To 59
            Dim str As String = i.ToString
            If Len(str) = 1 Then
                str = "0" + str
            End If
            ddl_sminute.Items.Insert(str, New ListItem(str, str))
            ddl_eminute.Items.Insert(str, New ListItem(str, str))
        Next
        ddl_sminute.Items.Insert(0, New ListItem("--請選擇--", "0"))
        ddl_eminute.Items.Insert(0, New ListItem("--請選擇--", "0"))
    End Sub

    '檢查題目類型是否重複控制 & 總合計 設定  lbl_score.Visible lbl_total.Visible
    Sub check_dgview(ByVal dt As DataTable)
        Dim i As Int16
        hid_qtype.Value = ""
        lbl_score.Text = "0"
        If Not dt Is Nothing Then
            For i = 0 To dt.Rows.Count - 1
                hid_qtype.Value += dt.Rows(i)("qtype").ToString + "," + dt.Rows(i)("etid").ToString + ","
                lbl_score.Text = Int(lbl_score.Text) + Int(dt.Rows(i)("score"))
            Next
        End If
        If lbl_score.Text = "0" Then
            lbl_score.Visible = False
            lbl_total.Visible = False
        Else
            lbl_score.Visible = True
            lbl_total.Visible = True
        End If
    End Sub

    '檢查題數是否大於題庫數 傳入  Me.ViewState("tmptable") 傳出  正常@True 異常@False 
    Function Check_Num(ByRef dt As DataTable) As Boolean
        '檢查題數是否大於題庫數 傳入  Me.ViewState("tmptable") 傳出  Me.ViewState("flag")
        Dim flag As Boolean = True
        'Dim i As Int16 = 0
        'Dim dt As DataTable = Me.ViewState("tmptable")
        If Not dt Is Nothing Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Int(dt.Rows(i)("num")) > Int(dt.Rows(i)("total_num")) Then
                    flag = False
                    Exit For
                End If
            Next
        End If
        Return flag
    End Function

    '配分可取小數第一位
    Function Check_Score(ByRef dt As DataTable) As Boolean
        '確認目前配分可取小數第一位 傳出值 正常@True 異常@False
        Dim flag As Boolean = True
        'Dim i As Int16 = 0
        'Dim dt As DataTable = Me.ViewState("tmptable")
        If Not dt Is Nothing Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Int(dt.Rows(i)("Score")) / Int(dt.Rows(i)("Num")) * 10 Mod 5 <> 0 Then
                    flag = False
                    Exit For
                End If
            Next
        End If
        Return flag
    End Function

    '目前題庫數
    Function show_total_num(ByVal int_etid As Int16, ByVal int_Petid As Int16, _
                            ByVal str_qtype As String, _
                            ByVal DistID As String) As Integer

        Dim sql As String = ""
        Dim dr_num As DataRow

        If int_etid <> int_Petid Then
            sql = "" & vbCrLf
            sql += " select count(eq.eqid) total_num " & vbCrLf
            sql += " from exam_question eq   " & vbCrLf
            sql += " join id_examtype ie   on ie.etid=eq.etid " & vbCrLf
            sql += " where eq.qtype='" & str_qtype & "'" & vbCrLf
            sql += " and ie.etid=" & int_etid & "" & vbCrLf
            sql += " and eq.DistID='" & DistID & "'" & vbCrLf
        Else
            sql = "" & vbCrLf
            sql += " select count(eq.eqid) total_num " & vbCrLf
            sql += " from exam_question eq   " & vbCrLf
            sql += " join id_examtype ie   on ie.etid=eq.etid " & vbCrLf
            sql += " where eq.qtype='" & str_qtype & "'" & vbCrLf
            sql += " and (ie.etid=" & int_etid & " or ie.parent=" & int_Petid & ")" & vbCrLf
            sql += " and eq.DistID='" & DistID & "'" & vbCrLf
        End If
        dr_num = DbAccess.GetOneRow(sql, objconn)

        Return dr_num("total_num")
    End Function


    '取得同父層之顯示內容
    Function show_question_CName(ByVal int_Petid As Integer, _
                                ByVal str_qtype As String, _
                                ByVal DistID As String, _
                                ByVal SType As String) As String
        Dim rst As String = ""
        Dim sql As String = ""
        Dim dt_Number As DataTable
        rst = ""

        sql = "" & vbCrLf
        sql += " select dbo.NVL(ie.parent,ie.etid) pETID,ie.etid cETID, ie.Name cName,count(*) totalnum " & vbCrLf
        sql += " from exam_question eq   " & vbCrLf
        sql += " join id_examtype ie   on ie.etid=eq.etid " & vbCrLf
        sql += " where 1=1 " & vbCrLf
        'sql += " and eq.StopUse IS NULL " & vbCrLf
        sql += " and (ie.etid=" & int_Petid & " or ie.parent=" & int_Petid & ")" & vbCrLf
        sql += " and eq.qtype='" & str_qtype & "' and eq.DistID='" & DistID & "'" & vbCrLf
        sql += " group by dbo.NVL(ie.parent,ie.etid)  ,ie.etid , ie.Name" & vbCrLf
        sql += " " & vbCrLf
        dt_Number = DbAccess.GetDataTable(sql, objconn)

        For int_Number As Integer = 0 To dt_Number.Rows.Count - 1
            Dim dr As DataRow = dt_Number.Rows(int_Number)

            Select Case SType
                Case "cname"
                    If dr("pETID").ToString <> dr("cETID").ToString Then
                        If rst <> "" Then rst += ","
                        rst += dr(SType).ToString
                    Else
                        If rst <> "" Then rst += ","
                        rst += "不區分"
                    End If
                Case "totalnum"
                    If dr("pETID").ToString <> dr("cETID").ToString Then
                        If rst <> "" Then rst += ","
                        rst += dr("cname").ToString & "(" & dr(SType).ToString & ")"
                    Else
                        If rst <> "" Then rst += ","
                        rst += "不區分(" & dr(SType).ToString & ")"
                    End If
            End Select
        Next

        Return rst
    End Function

    '產生亂數題庫eqid 傳出 strnum
    Function check_question(ByVal int_num As Int16, _
                            ByVal int_etid As Int16, _
                            ByVal str_qtype As String, _
                            ByVal DistID As String, _
                            ByVal int_Petid As Int16, _
                            Optional ByRef Errmsg As String = "") As String

        ''產生亂數題庫eqid 傳出 strnum
        'Function check_question(ByVal int_num As Int16, _
        '                        ByVal int_etid As Int16, _
        '                        ByVal str_qtype As String, _
        '                        ByVal DistID As String) As String

        '    Dim int_Petid As Int16 = 0
        '    Dim Errmsg As String = ""

        'int_num:傳入數量
        Const Cst_errormsg1 As String = "(意外狀況) 題組類別或題目類型不足，請檢視此資料正確性並重新建立!"
        Const Cst_errormsg2 As String = "每種子類別需出題數量((總出題數)/(子類別數量))   小於等於 零"
        Dim strnum As String = ""
        Dim dt As DataTable
        Dim sql As String

        Dim dt_Number As DataTable
        Dim number_of_each_question As Integer = 0 '每題所需分配數量

        Errmsg = ""
        strnum = ""

        If int_Petid = int_etid Then
            sql = "" & vbCrLf
            sql += " select dbo.NVL(ie.parent,ie.etid) pETID,ie.etid cETID ,count(*) Cnt " & vbCrLf
            sql += " from exam_question eq   " & vbCrLf
            sql += " join id_examtype ie   on ie.etid=eq.etid " & vbCrLf
            sql += " where 1=1 " & vbCrLf
            sql += " and eq.StopUse IS NULL " & vbCrLf '不啟用者不出題
            sql += " and (ie.etid=" & int_etid & " or ie.parent=" & int_Petid & ")" & vbCrLf
            sql += " and eq.qtype='" & str_qtype & "' and eq.DistID='" & DistID & "'" & vbCrLf
            sql += " group by dbo.NVL(ie.parent,ie.etid)  ,ie.etid " & vbCrLf
            sql += " " & vbCrLf
            dt_Number = DbAccess.GetDataTable(sql, objconn)

            If dt_Number.Rows.Count > 0 Then
                'Fix(9)
                number_of_each_question = Fix(int_num / dt_Number.Rows.Count) '每種子類別需出題數量=(總出題數)/(子類別數量)
                If number_of_each_question <= 0 Then
                    Errmsg = Cst_errormsg2
                End If
                For int_Number As Integer = 0 To dt_Number.Rows.Count - 1
                    Dim dr As DataRow = dt_Number.Rows(int_Number)
                    If number_of_each_question > dr("Cnt") Then
                        '每種子類別需出題數量 大於 子類別數量
                        strnum = ""
                        Errmsg = Cst_errormsg1
                        Exit For
                    Else
                        '此類別須出題數量為 (number_of_each_question)
                        sql = "select eqid from exam_question where etid=" & dr("cETID") & " and qtype='" & str_qtype & "' and DistID='" & DistID & "'"
                        sql += " and StopUse IS NULL " '不啟用者不出題
                        dt = DbAccess.GetDataTable(sql, objconn)
                        Dim num(dt.Rows.Count - 1) As Int16
                        For i As Integer = 0 To number_of_each_question - 1
                            num(i) = Int(TIMS.Rnd1X() * dt.Rows.Count) + 1 '取1亂數放到num(i)
                            For j As Integer = 0 To i - 1
                                If dt.Rows(num(i) - 1)("eqid") = dt.Rows(num(j) - 1)("eqid") Then
                                    i = i - 1 '有相同值，本次亂數 得 重新取得
                                    Exit For
                                End If
                            Next j
                        Next i
                        For i As Integer = 0 To number_of_each_question - 1
                            If strnum = "" Then
                                strnum = dt.Rows(num(i) - 1)("eqid").ToString
                            Else
                                strnum += "," & dt.Rows(num(i) - 1)("eqid").ToString
                            End If
                        Next
                    End If
                Next

                If strnum <> "" Then
                    '判斷取得數量是否足夠，不足得再取資料
                    Dim arystrnum() As String = strnum.Split(",")
                    If arystrnum.Length < int_num Then
                        '此類別須出題數量為 (number_of_each_question)
                        sql = "" & vbCrLf
                        sql += " select dbo.NVL(ie.parent,ie.etid) pETID,ie.etid cETID ,eq.eqid " & vbCrLf
                        sql += " from exam_question eq   " & vbCrLf
                        sql += " join id_examtype ie   on ie.etid=eq.etid " & vbCrLf
                        sql += " where 1=1 " & vbCrLf
                        sql += " and eq.StopUse IS NULL " '不啟用者不出題
                        sql += " and (ie.etid=" & int_etid & " or ie.parent=" & int_Petid & ")" & vbCrLf
                        sql += " and eq.qtype='" & str_qtype & "' and eq.DistID='" & DistID & "'" & vbCrLf
                        dt = DbAccess.GetDataTable(sql, objconn)
                        Dim num(dt.Rows.Count - 1) As Int16

                        number_of_each_question = int_num - arystrnum.Length
                        For i As Integer = 0 To number_of_each_question - 1
                            num(i) = Int(TIMS.Rnd1X() * dt.Rows.Count) + 1
                            If strnum.IndexOf(dt.Rows(num(i) - 1)("eqid")) > -1 Then
                                i = i - 1 '有相同值，本次亂數 得 重新取得
                            Else
                                If strnum = "" Then
                                    strnum = dt.Rows(num(i) - 1)("eqid").ToString
                                Else
                                    strnum += "," & dt.Rows(num(i) - 1)("eqid").ToString
                                End If
                            End If

                        Next
                    End If

                End If

            Else
                strnum = ""
                Errmsg = Cst_errormsg1
            End If

        Else
            sql = "select eqid from exam_question   where etid=" & int_etid & " and qtype='" & str_qtype & "' and DistID='" & DistID & "'"
            sql += " and StopUse IS NULL " '不啟用者不出題
            dt = DbAccess.GetDataTable(sql, objconn)
            '有資料 且沒有超過題庫數量
            If dt.Rows.Count > 0 And dt.Rows.Count >= int_num Then
                Dim num(dt.Rows.Count - 1) As Int16 '宣告資料陣列數量
                For i As Integer = 0 To int_num - 1
                    num(i) = Int(TIMS.Rnd1X() * dt.Rows.Count) + 1 '取1亂數放到num(i)

                    For j As Integer = 0 To i - 1
                        If dt.Rows(num(i) - 1)("eqid") = dt.Rows(num(j) - 1)("eqid") Then
                            i = i - 1 '有相同值，本次亂數 得 重新取得
                            Exit For
                        End If
                    Next j
                Next i
                For i As Integer = 0 To int_num - 1
                    If strnum = "" Then
                        strnum = dt.Rows(num(i) - 1)("eqid").ToString
                    Else
                        strnum += "," & dt.Rows(num(i) - 1)("eqid").ToString
                    End If
                Next
            End If
        End If

        If strnum = "" AndAlso Errmsg = "" Then
            Errmsg = Cst_errormsg1
        End If

        Return strnum
    End Function

    'Function SET_ddlcETID()
    '    ddl_cETID.Enabled = False
    '    If ddl_petid.SelectedValue <> "" Then
    '        ddl_cETID.Enabled = True
    '        ddl_cETID = Exam.Get_ExamType(ddl_cETID, sm.UserInfo.DistID, ddl_petid.SelectedValue)
    '    Else
    '        ddl_cETID.Items.Clear()
    '    End If
    'End Function

    Sub SET_ddlcETID(Optional ByVal cETID As Integer = -1)
        ddl_cETID.Enabled = False
        If ddl_pETID.SelectedValue <> "" Then
            ddl_cETID.Enabled = True
            ddl_cETID = cls_Exam.Get_ExamType(ddl_cETID, sm.UserInfo.DistID, ddl_pETID.SelectedValue)
            If cETID <> -1 Then
                Common.SetListItem(ddl_cETID, cETID)
            End If
        Else
            ddl_cETID.Items.Clear()
        End If
    End Sub

    Private Sub ddl_pETID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_pETID.SelectedIndexChanged
        SET_ddlcETID()
    End Sub
End Class
