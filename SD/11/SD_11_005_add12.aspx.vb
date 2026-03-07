Partial Class SD_11_005_add12
    Inherits AuthBasePage

    'SELECT * FROM STUD_QUESTIONFIN WHERE ROWNUM <=10 --產投 '受訓學員訓後動態調查表

    'Q1 1.學員目前的近況為何？1留任原公司,2轉換至同產業公司,3轉換至不同產業的公司,4待業中
    'Q2 2.學員於結訓後薪資有提升嗎？1大幅提升,2小幅提升,3沒有變化,4小幅減少,5大幅減少
    'Q3 3.學員的職位有變化嗎？1升遷,2調職,3沒有變化,4降職
    'Q4 4.學員對目前工作的滿意度是否有變化？1大幅提升,2小幅提升,3沒有變化,4小幅減少,5大幅減少
    'Q5 5.學員目前的工作內容是否與參訓課程內容相關？ 1非常相關,2相關,3尚可,4不相關,5非常不相關
    '(停用) Q6 學員是否同意參加訓練對目前工作表現或第二專長培育有幫助？1幫助非常大,2幫助頗多,3略有幫助,4幫助有限,5完全沒幫助
    'Q6_7 6-1.學員是否同意參加訓練對目前工作表現有幫助 ?  1幫助非常大,2幫助頗多,3略有幫助,4幫助有限,5完全沒幫助
    'Q6_8 6-2.學員是否同意參加訓練對第二專長培育有幫助 ?  1幫助非常大,2幫助頗多,3略有幫助,4幫助有限,5完全沒幫助
    '(停用) Q7 承上題，參加本項訓練對學員的幫助是在哪方面 ?  1對適應工作環境有幫助,2對目前工作績效有幫助,3對轉換工作跑道有幫助
    'Q8  7.學員是否有繼續參與進修訓練的意願 ?  1非常想參與,2想參與,3尚無想法,4不想參與,5非常不想參與
    'Q9_1_Note 8.學員認為還需要加強哪方面的專業知識使工作進行得更順利 ? 答1
    'Q9_2_Note 8.學員認為還需要加強哪方面的專業知識使工作進行得更順利 ? 答2
    'Q9_3_Note 8.學員認為還需要加強哪方面的專業知識使工作進行得更順利 ? 答3
    '(2015啟用)
    'Q10_1_Note 9.學員常跟本課程的哪些學員、教師或職員連絡？答1
    'Q10_2_Note 9.學員常跟本課程的哪些學員、教師或職員連絡？答2
    'Q10_3_Note 9.學員常跟本課程的哪些學員、教師或職員連絡？答3

    Const ss_QuestionFinSearchStr As String = "QuestionFinSearchStr"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在-------------------------- Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        ProcessType.Value = Request("ProcessType")
        Re_OCID.Value = Convert.ToString(Request("OCID"))
        Re_SOCID.Value = Convert.ToString(Request("SOCID")) '不一定有資料
        Re_ID.Value = Convert.ToString(Request("ID"))
        ProcessType.Value = TIMS.ClearSQM(ProcessType.Value)
        Re_OCID.Value = TIMS.ClearSQM(Re_OCID.Value)
        Re_SOCID.Value = TIMS.ClearSQM(Re_SOCID.Value)
        Re_ID.Value = TIMS.ClearSQM(Re_ID.Value)

#Region "(No Use)"

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    Re_ID.Value = Request("ID")
        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '       'Dim FunDr As DataRow
        '        FunDr = FunDrArray(0)
        '        Select Case ProcessType.Value
        '            Case "Update"
        '                Button1.Enabled = True
        '                If FunDr("Mod") = "0" And FunDr("Del") = "0" Then
        '                    Button1.Enabled = False
        '                    TIMS.Tooltip(Button1, "(Mod.Del)無權限使用該功能", True)
        '                End If
        '            Case "Insert", "next"
        '                Button1.Enabled = False
        '                If FunDr("Adds") = "1" Then
        '                    Button1.Enabled = True
        '                    'next_but.Enabled = True
        '                Else
        '                    TIMS.Tooltip(Button1, "(Adds)無權限使用該功能", True)
        '                    'TIMS.Tooltip(next_but, "(Adds)無權限使用該功能", True)
        '                End If
        '        End Select
        '    End If
        'End If

#End Region

        If Not IsPostBack Then
            If Convert.ToString(Request("OCID")) = "" Then
                Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
                Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
                Exit Sub
            End If
            'Re_OCID.Value = Convert.ToString(Request("OCID"))
            'Re_SOCID.Value = Convert.ToString(Request("SOCID")) '不一定有資料
            'Re_OCID.Value = TIMS.ClearSQM(Re_OCID.Value)
            'Re_SOCID.Value = TIMS.ClearSQM(Re_SOCID.Value)
            Call LoadCreateData1()
            SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
            Button1.Attributes.Add("OnClick", "return ChkData();")
            Q1.Attributes("onclick") = "Q1Is4();"
        End If
    End Sub

    '載入資料
    Sub LoadCreateData1()
        If Not Session(ss_QuestionFinSearchStr) Is Nothing Then
            Me.ViewState(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
            Session(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
        Else
            Session(ss_QuestionFinSearchStr) = Me.ViewState(ss_QuestionFinSearchStr)
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.StudentID" & vbCrLf
        sql &= " ,a.Name   + '(' + ISNULL(a.StudID,'') + ')' COLLATE CHINESE_TAIWAN_STROKE_CS_AS Name" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " FROM V_STUDENTINFO a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.OCID = @OCID " & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = Re_OCID.Value
            dt.Load(.ExecuteReader())
            'dt = DbAccess.GetDataTable(oCmd.CommandText, objconn, oCmd.Parameters)
        End With
        If dt.Rows.Count = 0 Then
            Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
            Exit Sub
        End If

        dt.DefaultView.Sort = "StudentID"
        With SOCID
            .DataSource = dt
            .DataTextField = "Name"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

        Re_OCID.Value = TIMS.ClearSQM(Re_OCID.Value)
        Re_SOCID.Value = TIMS.ClearSQM(Re_SOCID.Value)

        If Re_SOCID.Value <> "" Then
            Common.SetListItem(SOCID, Re_SOCID.Value)
            Select Case ProcessType.Value
                Case "del" '刪除
                    Dim sqlstrdel As String
                    sqlstrdel = " DELETE STUD_QUESTIONFIN WHERE SOCID = '" & Re_SOCID.Value & "' "
                    DbAccess.ExecuteNonQuery(sqlstrdel, objconn)
                    'Common.MessageBox(Me, "資料刪除完成", "SD_11_005.aspx?ID=" & Request("ID"))
                    Dim RqID As String = TIMS.Get_MRqID(Me)
                    Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID
                    TIMS.BlockAlert(Me, "資料刪除完成", uUrl1)
                    Exit Sub

                Case "check" '查詢
                    Call create(Re_SOCID.Value)
                    Button1.Enabled = False
                    next_but.Enabled = False
                    TIMS.Tooltip(Button1, "僅供查詢", True)
                    TIMS.Tooltip(next_but, "僅供查詢", True)
                Case "Edit" '修改
                    Call create(Re_SOCID.Value)
                    Button1.Enabled = True
                    next_but.Enabled = False
                    TIMS.Tooltip(next_but, "僅供修改", True)
                Case "next"
                    Call MoveNext() '儲存 自動呼叫下一筆
            End Select
        End If
    End Sub

    '建立學員顯示及答案 (依產投)
    Private Sub create(ByVal StrSOCID As String)
        StrSOCID = TIMS.ClearSQM(StrSOCID)
        Dim sqlstr As String = ""
        sqlstr = ""
        sqlstr &= " SELECT b.StudentID ,c.name ,b.StudStatus ,b.RejectTDate1 ,b.RejectTDate2 ,d.OrgKind "
        sqlstr += " FROM CLASS_CLASSINFO a "
        sqlstr += " JOIN CLASS_STUDENTSOFCLASS b ON a.ocid = b.ocid "
        sqlstr += " JOIN plan_planinfo p ON a.PlanID = p.PlanID AND a.comIDNO = p.comIDNO AND a.SeqNO = p.SeqNO "
        sqlstr += " JOIN stud_studentinfo c ON b.sid = c.sid "
        sqlstr += " JOIN Org_OrgInfo d ON d.ComIDNO = a.ComIDNO "
        sqlstr += " WHERE 1=1 "
        sqlstr += "    AND p.TPlanID IN (" & TIMS.Cst_TPlanID28_2 & ") "
        sqlstr += "    AND b.SOCID = '" & StrSOCID & "' "
        Dim row As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        If row Is Nothing Then
            Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
            Exit Sub
        End If
        Me.Label_Name.Text = Convert.ToString(row("name"))
        Me.Label_Stud.Text = "--"
        If Convert.ToString(row("StudentID")) <> "" Then
            Me.Label_Stud.Text = Convert.ToString(row("StudentID"))
        End If
        'Dim okind As String = Convert.ToString(row("OrgKind"))  '**by Milor 20080424
        Me.Label_Status.Text = TIMS.GET_STUDSTATUS_N23(row("StudStatus"), row("RejectTDate1"), row("RejectTDate2")) '"在訓"

        'Dim sqlstr As String
        sqlstr = ""
        sqlstr = "SELECT * FROM STUD_QUESTIONFIN WHERE SOCID = '" & StrSOCID & "'"
        Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        If Not row_list Is Nothing Then
            If Convert.ToString(row_list("Q1")) <> "" Then
                Common.SetListItem(Q1, Convert.ToString(row_list("Q1")))
            End If
            'If Not IsDBNull(row_list("Q1")) Then  Common.SetListItem(Q1, row_list("Q1"))
            If Convert.ToString(row_list("Q1")) = "4" Then
                Q2.Enabled = False
                Q3.Enabled = False
                Q4.Enabled = False
                Q5.Enabled = False
            Else
                If Not IsDBNull(row_list("Q2")) Then Common.SetListItem(Q2, row_list("Q2"))
                If Not IsDBNull(row_list("Q3")) Then Common.SetListItem(Q3, row_list("Q3"))
                If Not IsDBNull(row_list("Q4")) Then Common.SetListItem(Q4, row_list("Q4"))
                If Not IsDBNull(row_list("Q5")) Then Common.SetListItem(Q5, row_list("Q5"))
            End If
            'If Not IsDBNull(row_list("Q2")) Then  Common.SetListItem(Q2, row_list("Q2"))
            'If Not IsDBNull(row_list("Q3")) Then  Common.SetListItem(Q3, row_list("Q3"))
            'If Not IsDBNull(row_list("Q4")) Then  Common.SetListItem(Q4, row_list("Q4"))
            'If Not IsDBNull(row_list("Q5")) Then  Common.SetListItem(Q5, row_list("Q5"))
            'If Not IsDBNull(row_list("Q6")) Then Common.SetListItem(Q6, row_list("Q6"))
            If Not IsDBNull(row_list("Q6_7")) Then Common.SetListItem(Q6_7, row_list("Q6_7"))
            If Not IsDBNull(row_list("Q6_8")) Then Common.SetListItem(Q6_8, row_list("Q6_8"))
            'If Not IsDBNull(row_list("Q7")) Then Common.SetListItem(Q7, row_list("Q7"))
            If Not IsDBNull(row_list("Q8")) Then Common.SetListItem(Q8, row_list("Q8"))
            If Not IsDBNull(row_list("Q9_1_Note")) Then Q9_1_Note.Text = row_list("Q9_1_Note").ToString Else Q9_1_Note.Text = ""
            If Not IsDBNull(row_list("Q9_2_Note")) Then Q9_2_Note.Text = row_list("Q9_2_Note").ToString Else Q9_2_Note.Text = ""
            If Not IsDBNull(row_list("Q9_3_Note")) Then Q9_3_Note.Text = row_list("Q9_3_Note").ToString Else Q9_3_Note.Text = ""
            If Not IsDBNull(row_list("Q10_1_Note")) Then Q10_1_Note.Text = row_list("Q10_1_Note").ToString Else Q10_1_Note.Text = ""
            If Not IsDBNull(row_list("Q10_2_Note")) Then Q10_2_Note.Text = row_list("Q10_2_Note").ToString Else Q10_2_Note.Text = ""
            If Not IsDBNull(row_list("Q10_3_Note")) Then Q10_3_Note.Text = row_list("Q10_3_Note").ToString Else Q10_3_Note.Text = ""
            'If Not IsDBNull(row_list("BusName")) Then BusName.Text = row_list("BusName").ToString Else BusName.Text = ""
            'If Not IsDBNull(row_list("Q11")) Then  Common.SetListItem(Q11, row_list("Q11"))
            'If Not IsDBNull(row_list("Q12")) Then  Common.SetListItem(Q12, row_list("Q12"))
            'If Not IsDBNull(row_list("Q13")) Then  Common.SetListItem(Q13, row_list("Q13"))
            'If Not IsDBNull(row_list("Q14")) Then  Common.SetListItem(Q14, row_list("Q14"))
            'If Not IsDBNull(row_list("Q15")) Then  Common.SetListItem(Q15, row_list("Q15"))
            'If Not IsDBNull(row_list("Q16")) Then  Common.SetListItem(Q16, row_list("Q16"))
        Else
            Q1.SelectedIndex = -1
            Q2.SelectedIndex = -1
            Q3.SelectedIndex = -1
            Q4.SelectedIndex = -1
            Q5.SelectedIndex = -1
            'Q6.SelectedIndex = -1
            Q6_7.SelectedIndex = -1
            Q6_8.SelectedIndex = -1
            'Q7.SelectedIndex = -1
            Q8.SelectedIndex = -1
            Q9_1_Note.Text = ""
            Q9_2_Note.Text = ""
            Q9_3_Note.Text = ""
            Q10_1_Note.Text = ""
            Q10_2_Note.Text = ""
            Q10_3_Note.Text = ""
            'BusName.Text = ""
            'Q11.SelectedIndex = -1
            'Q12.SelectedIndex = -1
            'Q13.SelectedIndex = -1
            'Q14.SelectedIndex = -1
            'Q15.SelectedIndex = -1
            'Q16.SelectedIndex = -1
        End If
    End Sub

    '檢查是否為 最後一筆學員
    Private Sub check_last()
        If Not Session(ss_QuestionFinSearchStr) Is Nothing Then
            Me.ViewState(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
            Session(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
        Else
            Session(ss_QuestionFinSearchStr) = Me.ViewState(ss_QuestionFinSearchStr)
        End If
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "alert('已為此班級中最後一筆學員!!');" + vbCrLf
        'strScript += "location.href ='SD_11_005.aspx?ProcessType=Back&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("", strScript)
        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID & "&ProcessType=Back"
        TIMS.BlockAlert(Me, "已為此班級中最後一筆學員", uUrl1)
    End Sub

    '回上頁
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Session("QuestionFinSearchStr") = Me.ViewState("QuestionFinSearchStr")
        If Not Session(ss_QuestionFinSearchStr) Is Nothing Then
            Me.ViewState(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
            Session(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
        Else
            Session(ss_QuestionFinSearchStr) = Me.ViewState(ss_QuestionFinSearchStr)
        End If
        TIMS.Utl_Redirect1(Me, "SD_11_005.aspx?ProcessType=Back&ID=" & Re_ID.Value & "")
    End Sub

    '不儲存 呼叫下一筆
    Private Sub next_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles next_but.Click
        Call MoveNext()
    End Sub

    '下一筆
    Private Sub MoveNext()
        If SOCID.Items.Count > 0 Then
            Dim NowIndex As Integer
            Dim MaxIndex As Integer
            MaxIndex = SOCID.Items.Count - 1
            NowIndex = SOCID.SelectedIndex
            If NowIndex = MaxIndex Then
                Call check_last()
            Else
                SOCID.SelectedIndex = NowIndex + 1
                Re_SOCID.Value = SOCID.SelectedValue
                Call create(SOCID.SelectedValue)
            End If
        End If
    End Sub

    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        Re_SOCID.Value = SOCID.SelectedValue
        Call create(SOCID.SelectedValue)
    End Sub

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If SOCID.SelectedValue = "" Then
            Common.MessageBox(Me, "請選擇有效學員!!")
            Exit Sub
        End If
        If Not Session(ss_QuestionFinSearchStr) Is Nothing Then
            Me.ViewState(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
            Session(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
        Else
            Session(ss_QuestionFinSearchStr) = Me.ViewState(ss_QuestionFinSearchStr)
        End If
        Call SaveData1(SOCID.SelectedValue, objconn)
        Common.AddClientScript(Page, "insert_next();")
#Region "(No Use)"

        'Try
        'Catch ex As Exception
        '    DbAccess.RollbackTrans(objTrans)
        '    Common.MessageBox(Me, ex.ToString)
        '    'objconn.Close()
        '    Throw ex
        'End Try

#End Region
    End Sub

    Sub SaveData1(ByVal vSOCID As String, ByRef tConn As SqlConnection)
        Call TIMS.OpenDbConn(tConn)
        Dim sql As String = ""
        Dim flagI As Integer = 0 '0:修改 1:新增
        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim dtTable As DataTable = Nothing
        Dim objTrans As SqlTransaction = Nothing
        objTrans = DbAccess.BeginTrans(tConn)
        Try
            sql = " SELECT * FROM STUD_QUESTIONFIN WHERE SOCID = '" & vSOCID & "' "
            dtTable = DbAccess.GetDataTable(sql, sqlAdapter, objTrans)
            Dim dr_row As DataRow
            If dtTable.Rows.Count = 0 Then
                flagI = 1 '0:修改 1:新增
                dr_row = dtTable.NewRow
                dr_row("SOCID") = Val(vSOCID)
            Else
                dr_row = dtTable.Rows(0) '0:修改 1:新增
            End If
            If Q1.SelectedIndex = -1 Then dr_row("Q1") = Convert.DBNull Else dr_row("Q1") = Q1.SelectedValue
            If Q2.SelectedIndex = -1 Then dr_row("Q2") = Convert.DBNull Else dr_row("Q2") = Q2.SelectedValue
            If Q3.SelectedIndex = -1 Then dr_row("Q3") = Convert.DBNull Else dr_row("Q3") = Q3.SelectedValue
            If Q4.SelectedIndex = -1 Then dr_row("Q4") = Convert.DBNull Else dr_row("Q4") = Q4.SelectedValue
            If Q5.SelectedIndex = -1 Then dr_row("Q5") = Convert.DBNull Else dr_row("Q5") = Q5.SelectedValue
            'If Q6.SelectedIndex = -1 Then dr_row("Q6") = Convert.DBNull Else dr_row("Q6") = Q6.SelectedValue
            If Q6_7.SelectedIndex = -1 Then dr_row("Q6_7") = Convert.DBNull Else dr_row("Q6_7") = Q6_7.SelectedValue
            If Q6_8.SelectedIndex = -1 Then dr_row("Q6_8") = Convert.DBNull Else dr_row("Q6_8") = Q6_8.SelectedValue
            'If Q7.SelectedIndex = -1 Then dr_row("Q7") = Convert.DBNull Else dr_row("Q7") = Q7.SelectedValue
            If Q8.SelectedIndex = -1 Then dr_row("Q8") = Convert.DBNull Else dr_row("Q8") = Q8.SelectedValue
            If Q9_1_Note.Text = "" Then dr_row("Q9_1_Note") = Convert.DBNull Else dr_row("Q9_1_Note") = Q9_1_Note.Text
            If Q9_2_Note.Text = "" Then dr_row("Q9_2_Note") = Convert.DBNull Else dr_row("Q9_2_Note") = Q9_2_Note.Text
            If Q9_3_Note.Text = "" Then dr_row("Q9_3_Note") = Convert.DBNull Else dr_row("Q9_3_Note") = Q9_3_Note.Text
            If Q10_1_Note.Text = "" Then dr_row("Q10_1_Note") = Convert.DBNull Else dr_row("Q10_1_Note") = Q10_1_Note.Text
            If Q10_2_Note.Text = "" Then dr_row("Q10_2_Note") = Convert.DBNull Else dr_row("Q10_2_Note") = Q10_2_Note.Text
            If Q10_3_Note.Text = "" Then dr_row("Q10_3_Note") = Convert.DBNull Else dr_row("Q10_3_Note") = Q10_3_Note.Text
            'If BusName.Text = "" Then dr_row("BusName") = Convert.DBNull Else dr_row("BusName") = BusName.Text
            'If Q11.SelectedIndex = -1 Then dr_row("Q11") = Convert.DBNull Else dr_row("Q11") = Q11.SelectedValue
            'If Q12.SelectedIndex = -1 Then dr_row("Q12") = Convert.DBNull Else dr_row("Q12") = Q12.SelectedValue
            'If Q13.SelectedIndex = -1 Then dr_row("Q13") = Convert.DBNull Else dr_row("Q13") = Q13.SelectedValue
            'If Q14.SelectedIndex = -1 Then dr_row("Q14") = Convert.DBNull Else dr_row("Q14") = Q14.SelectedValue
            'If Q15.SelectedIndex = -1 Then dr_row("Q15") = Convert.DBNull Else dr_row("Q15") = Q15.SelectedValue
            'If Q16.SelectedIndex = -1 Then dr_row("Q16") = Convert.DBNull Else dr_row("Q16") = Q16.SelectedValue
            dr_row("DaSource") = "2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
            dr_row("ModifyAcct") = sm.UserInfo.UserID
            dr_row("ModifyDate") = Now()
            If flagI = 1 Then dtTable.Rows.Add(dr_row) '新增
            DbAccess.UpdateDataTable(dtTable, sqlAdapter, objTrans)
            DbAccess.CommitTrans(objTrans)
            'Call TIMS.CloseDbConn(tConn)
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            Throw ex
        End Try
    End Sub
End Class