Partial Class EXAM_03_002_R
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = dg_Sch

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            If Session("_SearchStr") IsNot Nothing Then
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
                tab_view.Visible = False
                Session("_SearchStr") = Nothing
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button5.Attributes("onclick") = "choose_class();"

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub btn_sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_sch.Click
        search()
    End Sub

    Sub search()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch) '顯示列數不正確

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        txt_examdate.Text = TIMS.Cdate3(TIMS.ClearSQM(txt_examdate.Text))

        Dim sql As String = ""
        sql += " SELECT DISTINCT" & vbCrLf
        sql += "    ec.ocid" & vbCrLf
        sql += "    ,ec.isonline ,ec.avail" & vbCrLf
        sql += "    ,convert(varchar(5), ec.examsdate, 108) examSTIME" & vbCrLf
        sql += "    ,convert(varchar(5), ec.examedate, 108) examETIME" & vbCrLf
        sql += "    ,cc.RID, ie.DistID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.classcname,cc.cycltype) classcname" & vbCrLf
        sql += "    ,CONVERT(varchar, cc.examdate, 111) examDATE" & vbCrLf
        sql += " FROM exam_classdata ec" & vbCrLf
        sql += " JOIN class_classinfo cc ON cc.ocid=ec.ocid" & vbCrLf
        sql += " JOIN id_examtype ie ON ie.etid=ec.etid" & vbCrLf
        sql += " JOIN auth_relship ar ON ar.RID=cc.RID" & vbCrLf
        sql += " WHERE cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        If RIDValue.Value <> "" Then
            sql += " AND cc.RID LIKE '" & RIDValue.Value & "%'" & vbCrLf
        Else
            sql += " AND cc.RID LIKE '" & sm.UserInfo.RID & "%'" & vbCrLf
        End If
        If sm.UserInfo.DistID <> "000" Then sql += " AND ie.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        If OCIDValue1.Value <> "" Then sql += " AND cc.ocid=" & OCIDValue1.Value & vbCrLf
        If txt_examdate.Text <> "" Then sql += " AND cc.examdate='" & txt_examdate.Text & "'" & vbCrLf

#Region "(No Use)"

        'If sm.UserInfo.DistID <> "000" Then
        '    sql += " and ie.DistID='" & sm.UserInfo.DistID & "'"
        'End If
        'If Len(OCIDValue1.Value) > 0 Then
        '    sql += " and cc.ocid=" & OCIDValue1.Value
        'End If
        'If Len(txt_examdate.Text) > 0 Then
        '    sql += " and cc.examdate='" & txt_examdate.Text & "'"
        'End If

#End Region

        msg.Visible = True
        tab_view.Visible = False
        PageControler1.Visible = False

        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)
        If TIMS.dtNODATA(dt1) Then Return

        msg.Visible = False
        tab_view.Visible = True
        PageControler1.Visible = True

        PageControler1.PageDataTable = dt1
        PageControler1.ControlerLoad()
    End Sub

    Private Sub dg_Sch_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnque As Button = e.Item.FindControl("btn_prt_que")
                Dim btnans As Button = e.Item.FindControl("btn_prt_ans")
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                If IsDBNull(drv("examdate")) Then e.Item.Cells(4).Text = "未填寫"
                Dim str_TYPEOE As String = If($"{drv("isonline")}" = "N", "一般筆試", "線上考試")
                e.Item.Cells(3).Text = str_TYPEOE
                Dim STRCMD1 As String = $"{drv("ocid")},{drv("isonline")}"
                btnque.CommandArgument = STRCMD1
                btnans.CommandArgument = STRCMD1
        End Select

    End Sub

    Private Sub dg_Sch_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        Dim arr() As String = Split(e.CommandArgument, ",")
        create_view(e.CommandName, arr(0), arr(1))
    End Sub

    Sub create_view(ByVal rpt_type As String, ByVal ocid As String, ByVal isonline As String)
        tab_sch.Visible = False
        tab_view.Visible = False
        rpt_view.Visible = True

        Dim sSortType As String = "1" '1:依順序(default) 2:亂數
        Dim int_score As Int16

        Dim dt_qtype As DataTable
        Dim dt_que As DataTable
        Dim dt_ans As DataTable
        Dim dr As DataRow
        'Dim dr_empty As DataRow
        Me.ViewState("rpt_type") = rpt_type

        Dim pms As New Hashtable From {{"OCID", ocid}, {"isonline", isonline}}
        Dim sql As String = "
SELECT cc.OCID
,concat(dbo.FN_CYEAR2(ip.YEARS),'年度',dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE)) CLASSNAME
,concat('甄試時間共' , examtime ,'分鐘(總分)') examtime
,ec.score,ec.sorttype,ec.isonline 
FROM exam_classdata ec
JOIN class_classinfo cc ON cc.ocid=ec.ocid
JOIN id_plan ip ON ip.planid=cc.planid
WHERE ec.ocid =@OCID AND ec.isonline =@isonline 
"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms)

        sSortType = "1"
        If dt.Rows.Count > 0 Then sSortType = Convert.ToString(dt.Rows(0)("SortType"))
        For i As Int16 = 0 To dt.Rows.Count - 1
            int_score = int_score + Int(dt.Rows(i)("score"))
        Next
        'lbl_distname.Text="職業訓練局 " & TIMS.GET_DistName(sm.UserInfo.DistID, objconn)   'dt.Rows(0)("distname") 
        lbl_distname.Text = "勞動力發展署 " & TIMS.GET_DISTNAME(objconn, sm.UserInfo.DistID)
        lbl_classname.Text = dt.Rows(0)("classname")
        lbl_examtime.Text = dt.Rows(0)("examtime") + int_score.ToString + "分)"
        dt.Clear()

        Select Case sSortType '1:依順序(default) 2:亂數
            Case "2" '1:依順序(default) 2:亂數
                dt.Columns.Add("title")
                dt.Columns.Add("data")
                'question
                Dim pms2 As New Hashtable From {{"OCID", ocid}, {"SortType", sSortType}}
                Dim sql2 As String = "
SELECT ecq.ecid ,eq.eqid ,ec.qtype ,eq.question,ecq.sort
FROM exam_classquestion ecq
JOIN exam_question eq ON eq.eqid=ecq.eqid
JOIN exam_classdata ec ON ec.ecid=ecq.ecid
WHERE ecq.ocid=@OCID AND ec.sorttype=@SortType
ORDER BY ecq.sort
"
                dt_que = DbAccess.GetDataTable(sql2, objconn, pms2)

                Dim num1 As Integer = 0
                For j As Int16 = 0 To dt_que.Rows.Count - 1
                    dr = dt.NewRow
                    num1 += 1
                    If rpt_type = "que" Then
                        '列印題目卷
                        If dt_que.Rows(j)("qtype") <> "4" Then
                            dr("title") = $"(　){num1}."
                            dr("data") = $"{dt_que.Rows(j)("question")}"
                        Else
                            dr("title") = $"{num1}."
                            dr("data") = $"{dt_que.Rows(j)("question")}"
                        End If
                    Else
                        '列印解答卷
                        Dim pmsqa As New Hashtable From {{"eqid", dt_que.Rows(j)("eqid")}}
                        Dim sqlqa As String = " SELECT answer ,isans FROM exam_answer WHERE eqid=@eqid"
                        Dim dt_que_ans As DataTable = DbAccess.GetDataTable(sqlqa, objconn, pmsqa)
                        If dt_que.Rows(j)("qtype") = "1" Then '是非題
                            If dt_que_ans.Rows(0)("isans") = "Y" Then
                                'dr("title")="(O)" & (j + 1).ToString & "."
                                dr("title") = $"(O){num1}."
                                dr("data") = $"{dt_que.Rows(j)("question")}"
                            Else
                                'dr("title")="(X)" & (j + 1).ToString & "."
                                dr("title") = $"(X){num1}."
                                dr("data") = $"{dt_que.Rows(j)("question")}"
                            End If
                        ElseIf dt_que.Rows(j)("qtype") = "4" Then '問答題
                            'dr("title")=(j + 1).ToString & "."
                            dr("title") = $"{num1}."
                            dr("data") = $"{dt_que.Rows(j)("question")}Ans:{dt_que_ans.Rows(0)("answer")}"
                        Else '選擇題&複選題
                            Dim str As String = ""
                            For k As Int16 = 0 To dt_que_ans.Rows.Count - 1
                                If dt_que_ans.Rows(k)("isans") = "Y" Then
                                    If str = "" Then
                                        str = $"{(k + 1)}"
                                    Else
                                        str = $"{str},{(k + 1)}"
                                    End If
                                End If
                            Next
                            'dr("title")="(" & str & ")" & (j + 1).ToString & "."
                            dr("title") = $"({str}){num1}."
                            dr("data") = $"{dt_que.Rows(j)("question")}"
                        End If
                    End If

                    'answer
                    sql = " SELECT eqid ,answer FROM exam_answer WHERE eqid=" & dt_que.Rows(j)("eqid")
                    dt_ans = DbAccess.GetDataTable(sql, objconn)
                    For k As Int16 = 0 To dt_ans.Rows.Count - 1
                        If dt_ans.Rows.Count <> 1 Then dr("data") += "　" & "(" & (k + 1).ToString & ")" & dt_ans.Rows(k)("answer")
                    Next
                    dt.Rows.Add(dr)
                Next
            Case Else '1:依順序(default) 2:亂數
                dt.Columns.Add("title")
                dt.Columns.Add("data")
                'qtpye
                Dim pms_QTYPE As New Hashtable From {{"OCID", ocid}, {"ISONLINE", isonline}}
                Dim sql_QTYPE As String = "
SELECT ecd.ecid
,concat(CASE CONVERT(varchar, ecd.qtype) WHEN '1' THEN '是非題' WHEN '2' THEN '選擇題' WHEN '3' THEN '複選題' WHEN '4' THEN '問答題' END
 ,'(共' ,ecd.num, '題每題' ,format(score/num,'N2'), '分，共' ,ecd.score,'分)') data
FROM exam_classdata ecd
JOIN id_examtype ie ON ie.etid=ecd.etid
LEFT JOIN id_examtype iep ON ie.parent=iep.etid
WHERE ecd.OCID=@OCID AND isonline=@ISONLINE
"
                dt_qtype = DbAccess.GetDataTable(sql_QTYPE, objconn, pms_QTYPE)

                Dim num1 As Integer = 0
                For i As Int16 = 0 To dt_qtype.Rows.Count - 1
                    ''題型
                    'dr=dt.NewRow
                    'dr("title")=check_num(i + 1) & "、"
                    'dr("data")=dt_qtype.Rows(i)("data")
                    'dt.Rows.Add(dr)

                    'question
                    Dim pms_que As New Hashtable From {{"OCID", ocid}, {"ECID", dt_qtype.Rows(i)("ecid")}}
                    Dim sql_que As String = "
SELECT ecq.ecid ,eq.eqid ,ec.qtype ,eq.question,ecq.sort
FROM exam_classquestion ecq
JOIN exam_question eq ON eq.eqid=ecq.eqid
JOIN exam_classdata ec ON ec.ecid=ecq.ecid
 WHERE ecq.OCID=@OCID AND ecq.ECID=@ECID
ORDER BY ecq.sort
"
                    dt_que = DbAccess.GetDataTable(sql_que, objconn, pms_que)
                    For j As Int16 = 0 To dt_que.Rows.Count - 1
                        dr = dt.NewRow
                        num1 += 1
                        If rpt_type = "que" Then
                            '列印題目卷
                            If dt_que.Rows(j)("qtype") <> "4" Then
                                'dr("title")="(　)" & (j + 1).ToString & "."
                                dr("title") = "(　)" & CStr(num1) & "."
                                dr("data") = dt_que.Rows(j)("question")
                            Else
                                'dr("title")=(j + 1).ToString & "."
                                dr("title") = CStr(num1) & "."
                                dr("data") = dt_que.Rows(j)("question")
                            End If
                        Else
                            '列印解答卷
                            Dim dt_que_ans As DataTable
                            sql = $" SELECT answer ,isans FROM exam_answer WHERE eqid={dt_que.Rows(j)("eqid")}"
                            dt_que_ans = DbAccess.GetDataTable(sql, objconn)
                            If dt_que.Rows(j)("qtype") = "1" Then '是非題
                                If dt_que_ans.Rows(0)("isans") = "Y" Then
                                    'dr("title")="(O)" & (j + 1).ToString & "."
                                    dr("title") = "(O)" & CStr(num1) & "."
                                    dr("data") = dt_que.Rows(j)("question")
                                Else
                                    'dr("title")="(X)" & (j + 1).ToString & "."
                                    dr("title") = "(X)" & CStr(num1) & "."
                                    dr("data") = dt_que.Rows(j)("question")
                                End If
                            ElseIf dt_que.Rows(j)("qtype") = "4" Then '問答題
                                'dr("title")=(j + 1).ToString & "."
                                dr("title") = CStr(num1) & "."
                                dr("data") = dt_que.Rows(j)("question") & "Ans:" & dt_que_ans.Rows(0)("answer")
                            Else '選擇題&複選題
                                Dim str As String = ""
                                For k As Int16 = 0 To dt_que_ans.Rows.Count - 1
                                    If dt_que_ans.Rows(k)("isans") = "Y" Then
                                        If str = "" Then
                                            str = (k + 1).ToString
                                        Else
                                            str = str + "," + (k + 1).ToString
                                        End If
                                    End If
                                Next
                                'dr("title")="(" & str & ")" & (j + 1).ToString & "."
                                dr("title") = "(" & str & ")" & CStr(num1) & "."
                                dr("data") = dt_que.Rows(j)("question")
                            End If
                        End If

                        'answer
                        sql = " SELECT eqid ,answer FROM exam_answer WHERE eqid=" & dt_que.Rows(j)("eqid")
                        dt_ans = DbAccess.GetDataTable(sql, objconn)
                        For k As Int16 = 0 To dt_ans.Rows.Count - 1
                            If dt_ans.Rows.Count <> 1 Then dr("data") += "　" & "(" & (k + 1).ToString & ")" & dt_ans.Rows(k)("answer")
                        Next
                        dt.Rows.Add(dr)
                    Next
                    'dr_empty=dt.NewRow
                    'dr_empty("title")="　"
                    'dr_empty("data")="　"
                    'dt.Rows.Add(dr_empty)
                Next
        End Select

        dg_view.DataSource = dt
        dg_view.DataBind()
    End Sub

    Private Sub btn_prt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_prt.Click
        Dim str As String
        If Me.ViewState("rpt_type") = "que" Then
            str = lbl_classname.Text + "學科甄試題目.doc"
        Else
            str = lbl_classname.Text + "學科甄試解答.doc"
        End If
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")

        '提示使用者是否要儲存檔案
        Dim sFileName As String = HttpUtility.UrlEncode(str, System.Text.Encoding.UTF8)
        Response.AddHeader("content-disposition", "attachment; filename=" & sFileName)

        '文件內容指定為work
        Response.ContentType = "application/vnd.ms-word;charset=utf-8"

        '繪出要輸出的html內容
        Dim strContent As New System.Text.StringBuilder
        Dim stringWrite As New System.IO.StringWriter(strContent)
        Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)

        div1.RenderControl(htmlWrite)

        strContent.Replace("<HTML>", "")
        strContent.Replace("</HTML>", "")

        Common.RespWrite(Me, strContent)

        '結束程式執行
        Response.End()
    End Sub

    Private Sub btn_lev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev.Click
        tab_sch.Visible = True
        tab_view.Visible = True
        rpt_view.Visible = False
    End Sub

#Region "(No Use)"

    '阿拉伯數字變中文數字
    'Private Function check_num(ByVal int_num As Int64)
    '    Dim str As String
    '    Select Case int_num
    '        Case 1
    '            str="一"
    '        Case 2
    '            str="二"
    '        Case 3
    '            str="三"
    '        Case 4
    '            str="四"
    '        Case 5
    '            str="五"
    '        Case 6
    '            str="六"
    '        Case 7
    '            str="七"
    '        Case 8
    '            str="八"
    '        Case 9
    '            str="九"
    '        Case 10
    '            str="十"
    '        Case Else
    '            str=CStr(int_num)
    '    End Select
    '    Return str
    'End Function

#End Region
End Class
