Partial Class EXAM_02_002
    Inherits AuthBasePage

    Const Cst_Split_flag As String = vbCrLf
    Const Cst_errmsg1 As String = "資料異常，查無資料!!"
    Const Cst_errmsg2 As String = "資料異常，請重新執行查詢作業!!"
    Dim rq_DistID As String = ""
    Dim rq_un As String = "" 'TIMS.ClearSQM(Request("un"))
    Dim rq_eqid As String = "" 'TIMS.ClearSQM(Request("eqid"))
    Dim rq_sch As String = "" 'TIMS.ClearSQM(Request("sch"))
    Dim rq_aryid As String = ""

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

        rq_DistID = TIMS.ClearSQM(Request("DistID"))
        rq_un = TIMS.ClearSQM(Request("un"))
        rq_eqid = TIMS.ClearSQM(Request("eqid"))
        rq_sch = TIMS.ClearSQM(Request("sch"))
        rq_aryid = TIMS.ClearSQM(Request("aryid"))

        If sm.UserInfo.DistID <> "000" Then rq_DistID = sm.UserInfo.DistID

        If Not IsPostBack Then
            '載入題組類別
            '執行修改題目
            If rq_un = "edit" Then
                '顯示題目選項
                '----------start----------
                Dim sql As String
                Dim dr As DataRow
                sql = "" & vbCrLf
                sql += " select " & vbCrLf
                sql += " 	eq.question,eq.etid,dbo.NVL(ie.Parent,eq.etid) petid,eq.qtype " & vbCrLf
                sql += " 	,eq.StopUse" & vbCrLf
                sql += " 	,ea.total " & vbCrLf
                sql += " from exam_question eq " & vbCrLf
                sql += " left join ID_ExamType ie on ie.ETID =eq.etid" & vbCrLf
                sql += " join (SELECT eqid,count(1) total " & vbCrLf
                sql += " 		FROM exam_answer   " & vbCrLf
                sql += " 		group by eqid) ea on ea.eqid=eq.eqid " & vbCrLf
                sql += " where 1=1 " & vbCrLf
                sql += " and eq.eqid=" & rq_eqid & vbCrLf

                Try
                    dr = DbAccess.GetOneRow(sql, objconn)
                    If dr Is Nothing Then
                        Common.MessageBox(Me, Cst_errmsg1)
                        Exit Sub
                    End If
                Catch ex As Exception
                    Common.MessageBox(Me, Cst_errmsg2)
                    Exit Sub
                End Try

                ddl_etid = cls_Exam.Get_ExamTypeParent(ddl_etid, rq_DistID, 1)

                Common.SetListItem(ddl_etid, dr("petid").ToString)
                SET_ddlcETID(dr("etid").ToString)

                txt_question.Text = Convert.ToString(dr("question"))
                ddl_qtype.SelectedValue = dr("qtype").ToString
                If dr("total") >= 2 Then
                    ddl_select2.SelectedValue = dr("total").ToString
                End If
                '不啟用
                chkStopUse.Checked = False
                If Convert.ToString(dr("StopUse")) = "Y" Then
                    chkStopUse.Checked = True
                End If
                btn_rst.Enabled = False
                TIMS.Tooltip(btn_rst, "不可重新設定")

                ddl_qtype.Enabled = False
                btn_check.Visible = False
                btn_rst.Visible = True
                '----------end----------

                '顯示解答
                '----------start----------
                Select Case dr("qtype")
                    Case 1
                        chg_table(1)
                        edit_search(1, 1, 0)
                    Case 4
                        chg_table(6)
                        edit_search(4, 1, 0)
                    Case Else
                        Select Case dr("total")
                            Case 2
                                chg_table(2)
                                edit_search(2, 1, 2)
                            Case 3
                                chg_table(3)
                                edit_search(2, 1, 3)
                            Case 4
                                chg_table(4)
                                edit_search(2, 1, 4)
                            Case 5
                                chg_table(5)
                                edit_search(2, 1, 5)
                        End Select
                End Select
                '----------end----------
                Me.ViewState("flag") = 1 '可儲存控制開始
            Else

                ddl_etid = cls_Exam.Get_ExamTypeParent(ddl_etid, rq_DistID, 1)

                SET_ddlcETID()

                '執行新增題目-將所有選項還原
                '----------start----------
                ddl_qtype.Enabled = True
                btn_check.Visible = True
                btn_rst.Visible = False
                chg_table(0)
                '----------end----------
            End If
            btn_save.Attributes("onclick") = "return check_save();"
            btn_check.Attributes("onclick") = "return check_select();"
        End If
    End Sub

    Sub SET_ddlcETID(Optional ByVal cETID As Integer = -1)
        ddl_cETID.Enabled = False
        If ddl_etid.SelectedValue <> "" Then
            ddl_cETID.Enabled = True

            ddl_cETID = cls_Exam.Get_ExamType(ddl_cETID, rq_DistID, ddl_etid.SelectedValue)

            If cETID <> -1 Then
                Common.SetListItem(ddl_cETID, cETID)
            End If
        Else
            ddl_cETID.Items.Clear()
        End If
    End Sub

#Region "NOUSE"

    ''取出鍵詞-上層類別代碼
    'Public Shared Function Get_ExamType(ByVal obj As ListControl, ByVal DistID As String) As ListControl
    '    Dim sql As String
    '    sql = "" & vbCrLf
    '    sql += " select c.etid" & vbCrLf
    '    sql += " 	,case when p.name is not null then p.name+'_'+c.name " & vbCrLf
    '    sql += " 	else c.name end as Name" & vbCrLf
    '    sql += " FROM ID_ExamType c" & vbCrLf
    '    sql += " LEFT join (SELECT * FROM ID_ExamType where avail=1 ) p on c.Parent =p.ETID" & vbCrLf
    '    sql += " where 1=1" & vbCrLf

    '    If DistID <> "000" Then '系統管理者可查全部
    '        sql += "and C.DistID = '" & DistID & "' and C.avail=1 " & vbCrLf
    '    Else
    '        sql += "and C.avail=1" & vbCrLf
    '    End If
    '    sql += " ORDER BY 2" & vbCrLf
    '    Dim dt As DataTable = DbAccess.GetDataTable(sql)
    '    With obj
    '        .DataSource = dt
    '        .DataTextField = "Name"
    '        .DataValueField = "ETID"
    '        .DataBind()
    '        If TypeOf obj Is DropDownList Then
    '            .Items.Insert(0, New ListItem("不區分", "0"))
    '        End If
    '    End With
    '    Return obj
    'End Function
#End Region

    Private Sub btn_check_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_check.Click
        '題目類別控制
        '----------start----------
        btn_rst.Visible = True
        btn_check.Visible = False
        ddl_qtype.Enabled = False
        ddl_select2.Enabled = False
        '----------end----------
        If ddl_qtype.SelectedValue = "1" Or ddl_qtype.SelectedValue = "4" Then
            Me.ViewState("check") = ddl_qtype.SelectedValue + "/" + "0" + "/" + ddl_select2.SelectedValue
        Else
            Me.ViewState("check") = ddl_qtype.SelectedValue + "/" + "1" + "/" + ddl_select2.SelectedValue
        End If


        '顯示table及內容
        '----------start----------
        Select Case ddl_qtype.SelectedValue
            Case "1"
                chg_table(1)
            Case "4"
                chg_table(6)
        End Select
        Select Case ddl_select2.SelectedValue
            Case "2"
                chg_table(2)
            Case "3"
                chg_table(3)
            Case "4"
                chg_table(4)
            Case "5"
                chg_table(5)
        End Select
        '----------end----------
        Me.ViewState("flag") = 1 '可儲存控制開始
    End Sub

    Private Sub btn_rst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_rst.Click
        btn_rst.Visible = False
        btn_check.Visible = True
        ddl_qtype.Enabled = True
        ddl_qtype.SelectedValue = "0"
        ddl_select2.SelectedValue = "0"
        chg_table(0)
        Me.ViewState("flag") = 0 '可儲存控制關閉
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If ddl_cETID.SelectedValue = "" Then
            Errmsg += "題組類別 未選擇" & vbCrLf
        End If

        txt_question.Text = txt_question.Text.Trim
        If txt_question.Text = "" Then
            Errmsg += "題目名稱 未輸入" & vbCrLf
        End If

        If ddl_qtype.SelectedValue <> "" Then
            Select Case ddl_qtype.SelectedValue
                Case "1", "2", "3", "4"
                Case Else
                    Errmsg += "題目類型 超過系統範圍(1=是非題;2=選擇;3=複選;4=問答;)" & vbCrLf
            End Select
        Else
            Errmsg += "題目類型 未選擇" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Dim stxt_question As String = txt_question.Text.Replace("'", "''")
        Dim i As Int16
        Dim sql As String
        Dim sql2 As String
        Dim sql_check As String
        Dim dr As DataRow
        Dim strScript As String
        sql_check = "select * from exam_question where 1=1"
        sql_check += " and etid=" & ddl_cETID.SelectedValue
        sql_check += " and question='" & stxt_question & "'"
        sql_check += " and qtype=" & ddl_qtype.SelectedValue

        If rq_un = "add" Then '判斷動作為-新增資料
            If Me.ViewState("flag") = 0 Then '判斷-可儲存控制
                Common.MessageBox(Me, "未點選【確認】填寫選項內容!")
                Exit Sub
            Else
                If TIMS.Get_SQLRecordCount(sql_check, objconn) > 0 Then '判斷資料是否重複
                    Common.MessageBox(Me, "此題目名稱重複!")
                    Exit Sub
                Else
                    '執行新增資料
                    '----------start----------
                    sql = "insert into exam_question(etid,question,qtype,distid,StopUse,modifyacct,modifydate)"
                    sql += "values(" & ddl_cETID.SelectedValue & ",'" & stxt_question & "'," & ddl_qtype.SelectedValue
                    sql += ",'" & rq_DistID & "'"
                    If Me.chkStopUse.Checked Then
                        sql += ",'Y'"
                    Else
                        sql += ",NULL"
                    End If
                    sql += ",'" & sm.UserInfo.UserID & "',getdate())"
                    'sql += FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4) & "')"
                    DbAccess.ExecuteNonQuery(sql, objconn)

                    sql = ""
                    sql += " select * from exam_question "
                    sql += " where 1=1 "
                    sql += " and etid=" & ddl_cETID.SelectedValue
                    sql += " and question='" & stxt_question & "'"
                    sql += " and qtype=" & ddl_qtype.SelectedValue
                    sql += " and distid='" & rq_DistID & "'"
                    dr = DbAccess.GetOneRow(sql, objconn)

                    Me.ViewState("eqid") = dr("eqid")
                    sql2 = "insert into exam_answer(eqid,answer,isans) values(" & dr("eqid") & ",'"
                    Select Case Me.ViewState("check")
                        Case "1/0/0"
                            sql = sql2 + "','" & rdo_ans1.SelectedValue & "')"
                            DbAccess.ExecuteNonQuery(sql, objconn)
                        Case "2/1/2", "3/1/2"  '單選-選項數有2筆時儲存
                            For i = 1 To 2
                                If i = 1 Then
                                    sql = sql2 + txt_ans2_1.Text.Replace("'", "''") & "','" & chk_isans("1") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                Else
                                    sql = sql2 + txt_ans2_2.Text.Replace("'", "''") & "','" & chk_isans("2") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                End If
                            Next
                        Case "2/1/3", "3/1/3" '單選-選項數有3筆時儲存
                            For i = 1 To 3
                                If i = 1 Then
                                    sql = sql2 + txt_ans2_1.Text.Replace("'", "''") & "','" & chk_isans("1") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                ElseIf i = 2 Then
                                    sql = sql2 + txt_ans2_2.Text.Replace("'", "''") & "','" & chk_isans("2") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                Else
                                    sql = sql2 + txt_ans2_3.Text.Replace("'", "''") & "','" & chk_isans("3") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                End If
                            Next
                        Case "2/1/4", "3/1/4" '單選-選項數有4筆時儲存
                            For i = 1 To 4
                                If i = 1 Then
                                    sql = sql2 + txt_ans2_1.Text.Replace("'", "''") & "','" & chk_isans("1") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                ElseIf i = 2 Then
                                    sql = sql2 + txt_ans2_2.Text.Replace("'", "''") & "','" & chk_isans("2") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                ElseIf i = 3 Then
                                    sql = sql2 + txt_ans2_3.Text.Replace("'", "''") & "','" & chk_isans("3") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                Else
                                    sql = sql2 + txt_ans2_4.Text.Replace("'", "''") & "','" & chk_isans("4") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                End If
                            Next
                        Case "2/1/5", "3/1/5" '單選-選項數有5筆時儲存
                            For i = 1 To 5
                                If i = 1 Then
                                    sql = sql2 + txt_ans2_1.Text.Replace("'", "''") & "','" & chk_isans("1") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                ElseIf i = 2 Then
                                    sql = sql2 + txt_ans2_2.Text.Replace("'", "''") & "','" & chk_isans("2") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                ElseIf i = 3 Then
                                    sql = sql2 + txt_ans2_3.Text.Replace("'", "''") & "','" & chk_isans("3") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                ElseIf i = 4 Then
                                    sql = sql2 + txt_ans2_4.Text.Replace("'", "''") & "','" & chk_isans("4") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                Else
                                    sql = sql2 + txt_ans2_5.Text.Replace("'", "''") & "','" & chk_isans("5") & "')"
                                    DbAccess.ExecuteNonQuery(sql, objconn)
                                End If
                            Next
                        Case "4/0/0" '問答題選項存儲
                            sql = sql2 + txt_ans4.Text.Replace("'", "''") & "','Y')"
                            DbAccess.ExecuteNonQuery(sql, objconn)
                    End Select
                    '----------end----------

                    Dim sUrl As String = ""
                    sUrl = ""
                    sUrl &= "EXAM_02_002.aspx?un=add"
                    sUrl &= "&DistID=" & rq_DistID

                    '判斷使用者繼續新增
                    '----------start----------
                    strScript = "<script language=""javascript"">" + vbCrLf
                    strScript += "if (window.confirm('資料新增成功!\n請問是否繼續新增？')){" + vbCrLf
                    'strScript += "  location.href='exam_02_002.aspx?un=add';" + vbCrLf
                    strScript += "  location.href='" & sUrl & "';" + vbCrLf
                    strScript += "} else {;" + vbCrLf
                    strScript += "  location.href='exam_02_001.aspx';" + vbCrLf
                    strScript &= "}" & vbCrLf
                    strScript &= "</script>"
                    Page.RegisterStartupScript("ring", strScript)
                    '----------end----------
                End If
            End If
        Else '判斷動作為-修改資料
            sql_check += " and eqid<>" & rq_eqid

            If TIMS.Get_SQLRecordCount(sql_check, objconn) > 0 Then '判斷資料是否重複
                Common.MessageBox(Me, "此題目名稱重複!")
                Exit Sub
            Else
                '執行修改資料
                '----------start----------
                Dim dt As DataTable = Nothing
                Dim da As SqlDataAdapter = Nothing
                Dim str As String = ""
                Dim arr() As String
                'Dim conn As New SqlConnection
                'conn = DbAccess.GetConnection
                sql = ""
                sql += " update exam_question"
                sql += " set etid=" & ddl_cETID.SelectedValue & ""
                sql += " ,question='" & stxt_question & "'"
                sql += " ,qtype=" & ddl_qtype.SelectedValue & ""
                sql += " ,distid='" & rq_DistID & "'"
                If Me.chkStopUse.Checked Then
                    sql += ",StopUse='Y'"
                Else
                    sql += ",StopUse=NULL"
                End If
                sql += " ,modifyacct='" & sm.UserInfo.UserID & "',modifydate=getdate() "
                'sql += ",modifyacct='" & sm.UserInfo.UserID & "',modifydate='" & FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4)
                sql += " where eqid=" & rq_eqid
                DbAccess.ExecuteNonQuery(sql, objconn)
                Select Case ddl_qtype.SelectedValue
                    Case "1"
                        sql = "update exam_answer set isans='" & rdo_ans1.SelectedValue & "' where eqid=" & rq_eqid
                        DbAccess.ExecuteNonQuery(sql, objconn)
                    Case "4"
                        sql = "update exam_answer set answer='" & txt_ans4.Text & "' where eqid=" & rq_eqid
                        DbAccess.ExecuteNonQuery(sql, objconn)
                    Case Else
                        sql = "select eaid from exam_answer where eqid=" & rq_eqid
                        dt = DbAccess.GetDataTable(sql)
                        For Each dr In dt.Rows
                            str += dr("eaid").ToString & Cst_Split_flag
                        Next
                        arr = Split(str, Cst_Split_flag)
                        If ddl_select2.SelectedValue >= "2" Then
                            sql = "update exam_answer set answer='" & txt_ans2_1.Text & "',isans='" & chk_isans("1") & "' "
                            sql += "where eaid=" & arr(0)
                            DbAccess.ExecuteNonQuery(sql, objconn)

                            sql = "update exam_answer set answer='" & txt_ans2_2.Text & "',isans='" & chk_isans("2") & "' "
                            sql += "where eaid=" & arr(1)
                            DbAccess.ExecuteNonQuery(sql, objconn)
                        End If
                        If ddl_select2.SelectedValue >= "3" Then
                            sql = "update exam_answer set answer='" & txt_ans2_3.Text & "',isans='" & chk_isans("3") & "' "
                            sql += "where eaid=" & arr(2)
                            DbAccess.ExecuteNonQuery(sql, objconn)
                        End If
                        If ddl_select2.SelectedValue >= "4" Then
                            sql = "update exam_answer set answer='" & txt_ans2_4.Text & "',isans='" & chk_isans("4") & "' "
                            sql += "where eaid=" & arr(3)
                            DbAccess.ExecuteNonQuery(sql, objconn)
                        End If
                        If ddl_select2.SelectedValue >= "5" Then
                            sql = "update exam_answer set answer='" & txt_ans2_5.Text & "',isans='" & chk_isans("5") & "' "
                            sql += "where eaid=" & arr(4)
                            DbAccess.ExecuteNonQuery(sql, objconn)
                        End If
                End Select
                Dim MRqID As String = TIMS.Get_MRqID(Me)
                Common.RespWrite(Me, "<script>alert('資料修改成功!');location.href='exam_02_001.aspx?un=reedit&sch=" & rq_sch & "&id=" & MRqID & "';</script>")
            End If
            '----------end----------
        End If

    End Sub

    Private Sub btn_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_exit.Click
        If rq_un = "add" Then
            TIMS.Utl_Redirect1(Me, "exam_02_001.aspx")
        Else
            Dim sUrl As String = ""
            sUrl = ""
            sUrl &= "EXAM_02_001.aspx?un=reedit"
            sUrl &= "&ID=" & TIMS.Get_MRqID(Me)
            sUrl &= "&DistID=" & rq_DistID
            sUrl &= "&eqid=" & rq_eqid
            sUrl &= "&sch=1"
            sUrl &= "&aryid=" & rq_aryid
            TIMS.Utl_Redirect1(Me, sUrl)
            'Common.RespWrite(Me, "<script>location.href='../02/EXAM_02_001.aspx?un=reedit&sch=" & rq_sch & "&id=" & Request("id") & "'</script>")
        End If
    End Sub

    '題目類別控制
    Private Sub ddl_qtype_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_qtype.SelectedIndexChanged
        Select Case ddl_qtype.SelectedValue
            Case "0", "1", "4"
                ddl_select2.Enabled = False
            Case "2", "3"
                ddl_select2.Enabled = True
        End Select
        ddl_select2.SelectedValue = "0"
    End Sub

    '選項改變時所呈現table設定
    Sub chg_table(ByVal parament As Integer)
        Select Case parament
            Case 0 '未選擇
                tab_select1.Visible = False
                tab_select2.Visible = False
                tab_select3.Visible = False
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 1 '是非題
                tab_select1.Visible = True
                tab_select2.Visible = False
                tab_select3.Visible = False
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 2 '項目數二
                tab_select1.Visible = False
                tab_select2.Visible = True
                tab_select3.Visible = False
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 3 '項目數三
                tab_select1.Visible = False
                tab_select2.Visible = True
                tab_select3.Visible = True
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 4 '項目數四
                tab_select1.Visible = False
                tab_select2.Visible = True
                tab_select3.Visible = True
                tab_select4.Visible = True
                tab_select5.Visible = False
                tab_select6.Visible = False
            Case 5 '項目數五
                tab_select1.Visible = False
                tab_select2.Visible = True
                tab_select3.Visible = True
                tab_select4.Visible = True
                tab_select5.Visible = True
                tab_select6.Visible = False
            Case 6 '問答題
                tab_select1.Visible = False
                tab_select2.Visible = False
                tab_select3.Visible = False
                tab_select4.Visible = False
                tab_select5.Visible = False
                tab_select6.Visible = True
        End Select
    End Sub

    '將checkbox勾選選項轉換成字串
    Public Function chk_isans(ByVal ans As String) As String
        Dim str As String = ""
        Select Case ans
            Case "1"
                If chk_ans2_1.Checked = True Then
                    str = "Y"
                Else
                    str = "N"
                End If
            Case "2"
                If chk_ans2_2.Checked = True Then
                    str = "Y"
                Else
                    str = "N"
                End If
            Case "3"
                If chk_ans2_3.Checked = True Then
                    str = "Y"
                Else
                    str = "N"
                End If
            Case "4"
                If chk_ans2_4.Checked = True Then
                    str = "Y"
                Else
                    str = "N"
                End If
            Case "5"
                If chk_ans2_5.Checked = True Then
                    str = "Y"
                Else
                    str = "N"
                End If
        End Select
        Return str
    End Function

    '代入DB值
    Sub edit_search(ByVal qtype As Int16, ByVal stype As Int16, ByVal total As Int16)
        Dim sql As String = ""
        Dim str As String = ""
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim arr() As String
        sql = "select answer,isans from exam_answer where 1=1 and eqid=" & rq_eqid
        Select Case qtype
            Case 1
                dr = DbAccess.GetOneRow(sql, objconn)
                rdo_ans1.SelectedValue = dr("isans")
            Case 2, 3
                dt = DbAccess.GetDataTable(sql, objconn)
                For Each dr In dt.Rows
                    str += dr("answer").ToString & Cst_Split_flag
                    str += dr("isans").ToString & Cst_Split_flag
                Next
                arr = Split(str, Cst_Split_flag)
                Select Case stype
                    Case 1
                        If total >= 2 Then
                            txt_ans2_1.Text = arr(0)
                            txt_ans2_2.Text = arr(2)
                            If arr(1) = "Y" Then
                                chk_ans2_1.Checked = True
                            End If
                            If arr(3) = "Y" Then
                                chk_ans2_2.Checked = True
                            End If
                        End If
                        If total >= 3 Then
                            txt_ans2_3.Text = arr(4)
                            If arr(5) = "Y" Then
                                chk_ans2_3.Checked = True
                            End If
                        End If
                        If total >= 4 Then
                            txt_ans2_4.Text = arr(6)
                            If arr(7) = "Y" Then
                                chk_ans2_4.Checked = True
                            End If
                        End If
                        If total = 5 Then
                            txt_ans2_5.Text = arr(8)
                            If arr(9) = "Y" Then
                                chk_ans2_5.Checked = True
                            End If
                        End If
                End Select
            Case 4
                dr = DbAccess.GetOneRow(sql, objconn)
                txt_ans4.Text = dr("answer")
        End Select
    End Sub

    Private Sub ddl_etid_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_etid.SelectedIndexChanged
        SET_ddlcETID()
    End Sub
End Class
