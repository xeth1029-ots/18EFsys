Partial Class SD_11_005_add17
    Inherits AuthBasePage

    'SELECT * FROM STUD_QUESTIONFIN WHERE ROWNUM <=10 --產投 '受訓學員訓後動態調查表

    Const ss_QuestionFinSearchStr As String = "QuestionFinSearchStr"
    Const cst_ptUpdate As String = "Update"
    Const cst_ptInsert As String = "Insert"
    Const cst_ptNext As String = "next"
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

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
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End

        ProcessType.Value = TIMS.ClearSQM(Request("ProcessType"))
        Re_OCID.Value = TIMS.ClearSQM(Request("OCID"))
        Re_SOCID.Value = TIMS.ClearSQM(Request("SOCID")) '不一定有資料
        Re_ID.Value = TIMS.ClearSQM(Request("ID"))

        If Not IsPostBack Then
            If Re_OCID.Value = "" Then
                'Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
                'Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
                Dim RqID As String = TIMS.Get_MRqID(Me)
                Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID
                TIMS.BlockAlert(Me, "傳入資料有誤，請重新操作該功能", uUrl1)
                Exit Sub
            End If
            Call LoadCreateData1()

            ddlSOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
            btnSave1.Attributes.Add("OnClick", "return ChkData();")
            'Q1.Attributes("onclick") = "Q1Is4();"
            rblQ1_1.Attributes("onclick") = "Clk_Q1x4();"
            rblQ1_2.Attributes("onclick") = "Clk_Q1x4();"
            rblQ1_3.Attributes("onclick") = "Clk_Q1x4();"
            rblQ1_4.Attributes("onclick") = "Clk_Q1x4();"
            rblQ1_5.Attributes("onclick") = "Clk_Q1x4();"
            rblQ1_6.Attributes("onclick") = "Clk_Q1x4();"
        End If
    End Sub

    Sub CHECK_SESS1()
        'Session("QuestionFinSearchStr") = Me.ViewState("QuestionFinSearchStr")
        If Session(ss_QuestionFinSearchStr) IsNot Nothing Then
            Me.ViewState(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
            Session(ss_QuestionFinSearchStr) = Session(ss_QuestionFinSearchStr)
        Else
            If Me.ViewState(ss_QuestionFinSearchStr) IsNot Nothing Then
                Session(ss_QuestionFinSearchStr) = Me.ViewState(ss_QuestionFinSearchStr)
            End If
        End If
    End Sub

    '載入資料
    Sub LoadCreateData1()
        Call CHECK_SESS1()
        Call TIMS.OpenDbConn(objconn)

        Dim s_parms As New Hashtable
        s_parms.Add("OCID", Re_OCID.Value)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.STUDENTID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(a.STUDENTID) STUDID2" & vbCrLf
        sql &= " ,concat(a.NAME,'(',dbo.fn_CSTUDID2(a.STUDENTID),')') NAME2" & vbCrLf
        sql &= " ,a.STUDSTATUS" & vbCrLf
        sql &= " ,q.DASOURCE" & vbCrLf
        sql &= " FROM V_STUDENTINFO a" & vbCrLf
        sql &= " LEFT JOIN STUD_QUESTIONFIN q on q.SOCID=a.SOCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND a.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql &= " AND a.Years=  '" & sm.UserInfo.Years & "' " & vbCrLf
        Else
            sql &= " AND a.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        sql &= " AND a.OCID=@OCID" & vbCrLf
        'sql &= " AND a.TPLANID='28'" & vbCrLf 'sql &= " AND a.YEARS='2021'" & vbCrLf 'sql &= " AND a.OCID=133342" & vbCrLf
        sql &= " AND ISNULL(q.DASOURCE, 0) != 1" & vbCrLf
        'sql &= " ORDER BY STUDID2" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, s_parms)
        If dt.Rows.Count = 0 Then
            Dim RqID As String = TIMS.Get_MRqID(Me)
            Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID
            TIMS.BlockAlert(Me, "傳入資料有誤，請重新操作該功能", uUrl1)
            Exit Sub
        End If

        dt.DefaultView.Sort = "STUDENTID"
        With ddlSOCID
            .DataSource = dt
            .DataTextField = "NAME2"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

        Re_OCID.Value = TIMS.ClearSQM(Re_OCID.Value)
        Re_SOCID.Value = TIMS.ClearSQM(Re_SOCID.Value)
        If Re_SOCID.Value = "" Then Exit Sub

        Common.SetListItem(ddlSOCID, Re_SOCID.Value)

        Select Case ProcessType.Value'ProcessTypeVal
            Case "del" '刪除
                Dim RqID As String = TIMS.Get_MRqID(Me)
                Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID

                Dim s_FINDASOURCE As String = Get_FINDASOURCE(Re_SOCID.Value, objconn)
                Dim flag_SAVE_Enabled As Boolean = If(s_FINDASOURCE = "1", False, True)
                If Not flag_SAVE_Enabled Then
                    'btnSave1.Enabled = flag_SAVE_Enabled
                    TIMS.BlockAlert(Me, "學員外網填寫，不可刪除!!", uUrl1)
                    Exit Sub
                End If

                Dim sqlstrdel As String
                sqlstrdel = "DELETE STUD_QUESTIONFIN WHERE SOCID= '" & Re_SOCID.Value & "'"
                DbAccess.ExecuteNonQuery(sqlstrdel, objconn)
                TIMS.BlockAlert(Me, "資料刪除完成", uUrl1)
                Exit Sub

            Case "check" '查詢
                Call SHOW_STUDDATA1(Re_SOCID.Value, objconn)
            Case "Edit" '修改
                Call SHOW_STUDDATA1(Re_SOCID.Value, objconn)
            Case cst_ptNext '"next"'下一筆
                Call MoveNext() '儲存 自動呼叫下一筆
        End Select

    End Sub

    '清理答案
    Sub sub_Clear1()
        rblQ1_1.Checked = False
        rblQ1_2.Checked = False
        rblQ1_3.Checked = False
        rblQ1_4.Checked = False
        rblQ1_5.Checked = False
        rblQ1_6.Checked = False
        Q1abc.SelectedIndex = -1
        'Call TIMS.SetMyValue(Q1abc, "")
        Q2.SelectedIndex = -1
        Q3.SelectedIndex = -1
        Q4.SelectedIndex = -1
        Q5.SelectedIndex = -1
        Q6.SelectedIndex = -1
        Call TIMS.SetCblValue(Q7, "")
        Q211.SelectedIndex = -1
        Q212.SelectedIndex = -1
        Q213.SelectedIndex = -1
        Q214.SelectedIndex = -1
        Q215.SelectedIndex = -1
        Q216.SelectedIndex = -1
        Q217.SelectedIndex = -1
        Q218.SelectedIndex = -1
        Q221.SelectedIndex = -1
        Q222.SelectedIndex = -1
        Q223.SelectedIndex = -1
        Q224.SelectedIndex = -1
        Q225.SelectedIndex = -1
        Q226.SelectedIndex = -1
        Q3_Note.Text = ""
    End Sub

    '建立學員顯示及答案 (依產投)
    Sub SHOW_STUDDATA1(ByVal strSOCID As String, ByRef oConn As SqlConnection)
        Call sub_Clear1()
        strSOCID = TIMS.ClearSQM(strSOCID)
        If strSOCID = "" Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT b.StudentID ,c.name ,b.StudStatus " & vbCrLf
        sql &= "  ,(CASE WHEN b.StudStatus = 1 THEN '在訓' WHEN b.StudStatus = 2 THEN '離訓' WHEN b.StudStatus = 3 THEN '退訓' WHEN b.StudStatus = 4 THEN '續訓' WHEN b.StudStatus = 5 THEN '結訓' ELSE '在訓' END) StudStatus2 " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, b.RejectTDate1, 111) RejectTDate1 " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, b.RejectTDate2, 111) RejectTDate2 ,d.OrgKind " & vbCrLf
        sql &= " FROM CLASS_CLASSINFO a " & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON a.ocid = b.ocid " & vbCrLf
        sql &= " JOIN PLAN_PLANINFO p ON a.PlanID = p.PlanID and a.comIDNO = p.comIDNO and a.SeqNO = p.SeqNO" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO c ON b.sid = c.sid " & vbCrLf
        sql &= " JOIN ORG_ORGINFO d ON d.ComIDNO = a.ComIDNO " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND p.TPlanID IN (" & TIMS.Cst_TPlanID28_2 & ") " & vbCrLf
        sql &= " AND b.SOCID = '" & strSOCID & "' " & vbCrLf
        Dim row As DataRow = DbAccess.GetOneRow(sql, oConn)
        If row Is Nothing Then
            'Common.RespWrite(Me, "<script>alert('傳入資料有誤，請重新操作該功能');</script>")
            'Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
            'Exit Sub
            Dim RqID As String = TIMS.Get_MRqID(Me)
            Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID
            TIMS.BlockAlert(Me, "傳入資料有誤，請重新操作該功能", uUrl1)
            Exit Sub
        End If

        Label_Name.Text = Convert.ToString(row("name"))
        Label_Stud.Text = If(Convert.ToString(row("StudentID")) <> "", Convert.ToString(row("StudentID")), "--")

        'Dim okind As String = Convert.ToString(row("OrgKind"))  '**by Milor 20080424
        Dim str As String = Convert.ToString(row("StudStatus2"))
        Me.Label_Status.Text = Convert.ToString(row("StudStatus2"))
        Select Case Convert.ToString(row("StudStatus"))
            Case "2"
                str &= "(" & Convert.ToString(row("RejectTDate1")) & ")"
                'Me.Label_Status.Text = str
            Case "3"
                str &= "(" & Convert.ToString(row("StudStatus2")) & ")"
        End Select
        Me.Label_Status.Text = str

        Call SHOW_STUDDATA2(strSOCID, oConn)

        Dim strScript1 As String = ""
        strScript1 = "<script>Clk_Q1x4();</script>" & vbCrLf
        TIMS.RegisterStartupScript(Me, TIMS.xBlockName, strScript1)
    End Sub

    'DASOURCE  '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。) 2:TIMS系統() 學員填寫不可列印。
    ''' <summary> 檢核DASOURCE </summary>
    ''' <param name="strSOCID"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Public Shared Function Get_FINDASOURCE(ByVal strSOCID As String, ByRef oConn As SqlConnection) As String
        strSOCID = TIMS.ClearSQM(strSOCID)
        If strSOCID = "" Then Return ""
        Dim rst As String = ""
        Dim sqlstr As String = ""
        sqlstr = "SELECT DASOURCE FROM STUD_QUESTIONFIN WHERE SOCID = '" & strSOCID & "'"
        Dim dr2 As DataRow = DbAccess.GetOneRow(sqlstr, oConn)
        If dr2 Is Nothing Then Return rst 'Exit Sub

        'DASOURCE  '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。) 2:TIMS系統() 學員填寫不可列印。
        'Hid_DASOURCE.Value = Convert.ToString(dr2("DASOURCE"))
        'Dim flag_SAVE_Enabled As Boolean = If(Hid_DASOURCE.Value = "1", False, True)
        rst = Convert.ToString(dr2("DASOURCE")) 'Hid_DASOURCE.Value
        Return rst
    End Function

    '建立學員顯示及答案 (依產投)2
    Sub SHOW_STUDDATA2(ByVal strSOCID As String, ByRef oConn As SqlConnection)
        strSOCID = TIMS.ClearSQM(strSOCID)
        If strSOCID = "" Then Return

        'Dim sqlstr As String
        Dim sqlstr As String = ""
        sqlstr = "SELECT * FROM STUD_QUESTIONFIN WHERE SOCID = '" & strSOCID & "'"
        Dim dr2 As DataRow = DbAccess.GetOneRow(sqlstr, oConn)
        If dr2 Is Nothing Then Exit Sub
        'DASOURCE  '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。) 2:TIMS系統() 學員填寫不可列印。
        Hid_DASOURCE.Value = Convert.ToString(dr2("DASOURCE"))
        Dim flag_SAVE_Enabled As Boolean = If(Hid_DASOURCE.Value = "1", False, True)
        btnSave1.Enabled = flag_SAVE_Enabled
        If Not flag_SAVE_Enabled Then
            TIMS.Tooltip(btnSave1, "學員外網填寫，不可修改", True)
        Else
            TIMS.Tooltip(btnSave1, "")
        End If

        '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
        '新增「填寫來源」檢核機制，由系統判斷填寫來源為產投報名網或TIMS系統：
        '1.若為報名網，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰，無法逕行修改，僅保留填寫狀態供訓練單位查詢。
        '2.若為TIMS系統，保留訓練單位各項功能(新增、修改、查詢、列印、清除重填)。
        '3.若有訓練單位先於TIMS系統協助填寫，學員後於報名網修正之情形者，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰。
        'PS:訓練單位登入,若是學員於外網填寫,是不可修改,不可查詢,不可列印,不可清除重填   但若是分署登入,若是學員於外網填寫,是不可修改,可查詢,可列印,不可清除重填

        If Hid_DASOURCE.Value = "1" Then
            '委訓單位
            If sm.UserInfo.LID = 2 Then
                Dim RqID As String = TIMS.Get_MRqID(Me)
                Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID
                TIMS.BlockAlert(Me, "學員外網填寫，不可查看", uUrl1)
                Exit Sub
            Else
                Select Case ProcessType.Value 'ProcessTypeVal
                    Case "check" '查詢 
                        btnSave1.Enabled = False
                        btnNext1.Enabled = False
                        TIMS.Tooltip(btnSave1, "僅供查詢", True)
                        TIMS.Tooltip(btnNext1, "僅供查詢", True)
                    Case Else 'Case "del" '刪除 Case "Edit" '修改 Case cst_ptNext '"next"'下一筆
                        Dim RqID As String = TIMS.Get_MRqID(Me)
                        Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID
                        TIMS.BlockAlert(Me, "學員外網填寫，不可修改", uUrl1)
                        Exit Sub
                End Select
            End If
        End If

        Select Case ProcessType.Value'ProcessTypeVal
            Case "check" '查詢
                btnSave1.Enabled = False
                btnNext1.Enabled = False
                TIMS.Tooltip(btnSave1, "僅供查詢", True)
                TIMS.Tooltip(btnNext1, "僅供查詢", True)
        End Select

        rblQ1_1.Checked = If(Convert.ToString(dr2("Q1")) = "1", True, False)
        rblQ1_2.Checked = If(Convert.ToString(dr2("Q1")) = "2", True, False)
        rblQ1_3.Checked = If(Convert.ToString(dr2("Q1")) = "3", True, False)
        rblQ1_4.Checked = If(Convert.ToString(dr2("Q1")) = "4", True, False)
        rblQ1_5.Checked = If(Convert.ToString(dr2("Q1")) = "5", True, False)
        rblQ1_6.Checked = If(Convert.ToString(dr2("Q1")) = "6", True, False)

        Dim MyValue As String = ""
        If Convert.ToString(dr2("Q1a")) = "Y" Then MyValue = "a"
        If Convert.ToString(dr2("Q1b")) = "Y" Then MyValue = "b"
        If Convert.ToString(dr2("Q1c")) = "Y" Then MyValue = "c"
        Common.SetListItem(Q1abc, MyValue)

        Common.SetListItem(Q2, Convert.ToString(dr2("Q2")))
        Common.SetListItem(Q3, Convert.ToString(dr2("Q3")))
        Common.SetListItem(Q4, Convert.ToString(dr2("Q4")))
        Common.SetListItem(Q5, Convert.ToString(dr2("Q5")))
        Common.SetListItem(Q6, Convert.ToString(dr2("Q8")))
        MyValue = ""
        If Convert.ToString(dr2("Q7MR1")) = "Y" Then
            If MyValue <> "" Then MyValue &= ","
            MyValue &= "1"
        End If
        If Convert.ToString(dr2("Q7MR2")) = "Y" Then
            If MyValue <> "" Then MyValue &= ","
            MyValue &= "2"
        End If
        If Convert.ToString(dr2("Q7MR3")) = "Y" Then
            If MyValue <> "" Then MyValue &= ","
            MyValue &= "3"
        End If
        If Convert.ToString(dr2("Q7MR4")) = "Y" Then
            If MyValue <> "" Then MyValue &= ","
            MyValue &= "4"
        End If
        TIMS.SetCblValue(Q7, MyValue)

        Common.SetListItem(Q211, Convert.ToString(dr2("Q211")))
        Common.SetListItem(Q212, Convert.ToString(dr2("Q212")))
        Common.SetListItem(Q213, Convert.ToString(dr2("Q213")))
        Common.SetListItem(Q214, Convert.ToString(dr2("Q214")))
        Common.SetListItem(Q215, Convert.ToString(dr2("Q215")))
        Common.SetListItem(Q216, Convert.ToString(dr2("Q216")))
        Common.SetListItem(Q217, Convert.ToString(dr2("Q217")))
        Common.SetListItem(Q218, Convert.ToString(dr2("Q218")))
        Common.SetListItem(Q221, Convert.ToString(dr2("Q221")))
        Common.SetListItem(Q222, Convert.ToString(dr2("Q222")))
        Common.SetListItem(Q223, Convert.ToString(dr2("Q223")))
        Common.SetListItem(Q224, Convert.ToString(dr2("Q224")))
        Common.SetListItem(Q225, Convert.ToString(dr2("Q225")))
        Common.SetListItem(Q226, Convert.ToString(dr2("Q226")))
        Q3_Note.Text = Convert.ToString(dr2("Q3_NOTE"))
    End Sub

    '檢查是否為 最後一筆學員
    Private Sub check_last()
        Call CHECK_SESS1()
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "alert('已為此班級中最後一筆學員!!');" + vbCrLf
        'strScript += "location.href ='SD_11_005.aspx?ProcessType=Back&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("", strScript)
        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim uUrl1 As String = "SD/11/SD_11_005.aspx?ID=" & RqID & "&ProcessType=Back"
        TIMS.BlockAlert(Me, "已為此班級中最後一筆學員", uUrl1)
        'Exit Sub
    End Sub

    '回上1頁
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack1.Click
        Call CHECK_SESS1()

        TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "SD_11_005.aspx?ProcessType=Back&ID=" & Re_ID.Value & "")
    End Sub

    '不儲存 呼叫下1筆
    Private Sub next_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext1.Click
        Call MoveNext() '自動呼叫下一筆
    End Sub

    ''' <summary> 下一筆 </summary>
    Private Sub MoveNext()
        'Dim flag_chktest As Boolean = TIMS.sUtl_ChkTest()
        'Dim slogMsg1 As String = ""
        'slogMsg1 &= "org:ListText:" & TIMS.GetListText(SOCID) & vbCrLf
        'slogMsg1 &= "org:ListValue:" & TIMS.GetListValue(SOCID) & vbCrLf
        'If flag_chktest Then TIMS.writeLog(Me, slogMsg1)

        If ddlSOCID.Items.Count > 0 Then
            Dim iNowIndex As Integer = ddlSOCID.SelectedIndex
            Dim iMaxIndex As Integer = ddlSOCID.Items.Count - 1
            'MaxIndex = SOCID.Items.Count - 1
            'NowIndex = SOCID.SelectedIndex
            If iNowIndex = iMaxIndex Then
                Call check_last()
                Exit Sub
            End If

            ddlSOCID.SelectedIndex = iNowIndex + 1
            Dim v_ddlSOCID As String = TIMS.GetListValue(ddlSOCID)
            Re_SOCID.Value = v_ddlSOCID 'SOCID.SelectedValue
            'Dim slogMsg1 As String = ""
            'slogMsg1 = ""
            'slogMsg1 &= "next:iNowIndex:" & iNowIndex & vbCrLf
            'slogMsg1 &= "next:iMaxIndex:" & iMaxIndex & vbCrLf
            'slogMsg1 &= "next:ListText:" & TIMS.GetListText(SOCID) & vbCrLf
            'slogMsg1 &= "next:ListValue:" & TIMS.GetListValue(SOCID) & vbCrLf
            'If flag_chktest Then TIMS.writeLog(Me, slogMsg1)
            Call SHOW_STUDDATA1(v_ddlSOCID, objconn)
        End If
    End Sub

    Private Sub ddlSOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlSOCID.SelectedIndexChanged
        Dim v_ddlSOCID As String = TIMS.GetListValue(ddlSOCID)
        Re_SOCID.Value = v_ddlSOCID 'SOCID.SelectedValue
        Call SHOW_STUDDATA1(v_ddlSOCID, objconn)
    End Sub

    '儲存
    Sub SaveData1(ByVal vSOCID As String, ByRef tConn As SqlConnection)
        Call TIMS.OpenDbConn(tConn)
        Dim sMODIFYACCT As String = sm.UserInfo.UserID
        Const cst_DASOURCE As String = "2" ' "2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()

        Dim i_sql As String = ""
        Dim u_sql As String = ""

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO STUD_QUESTIONFIN ( " & vbCrLf
        sql &= "  SOCID, Q1, Q2, Q3, Q4, Q5, Q8, Q1A, Q1B, Q1C, Q7MR1, Q7MR2, Q7MR3, Q7MR4 " & vbCrLf
        sql &= "  ,Q211 ,Q212 ,Q213 ,Q214 ,Q215 ,Q216 ,Q217 ,Q218 ,Q221 ,Q222 ,Q223 ,Q224 ,Q225 ,Q226 ,Q3_NOTE " & vbCrLf
        sql &= "  ,DASOURCE ,MODIFYACCT ,MODIFYDATE " & vbCrLf
        sql &= "  ) VALUES ( " & vbCrLf
        sql &= "  @SOCID, @Q1, @Q2, @Q3, @Q4, @Q5, @Q8, @Q1A, @Q1B, @Q1C, @Q7MR1, @Q7MR2, @Q7MR3, @Q7MR4 " & vbCrLf
        sql &= "  ,@Q211, @Q212, @Q213, @Q214, @Q215, @Q216, @Q217, @Q218, @Q221, @Q222, @Q223, @Q224, @Q225, @Q226, @Q3_NOTE " & vbCrLf
        sql &= "  ,@DASOURCE, @MODIFYACCT, GETDATE() " & vbCrLf
        sql &= " ) " & vbCrLf
        i_sql = sql
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " UPDATE STUD_QUESTIONFIN " & vbCrLf
        sql &= " SET Q1 = @Q1 " & vbCrLf
        sql &= " ,Q2 = @Q2" & vbCrLf
        sql &= " ,Q3 = @Q3" & vbCrLf
        sql &= " ,Q4 = @Q4" & vbCrLf
        sql &= " ,Q5 = @Q5" & vbCrLf
        sql &= " ,Q8 = @Q8" & vbCrLf
        sql &= " ,Q1A = @Q1A" & vbCrLf
        sql &= " ,Q1B = @Q1B" & vbCrLf
        sql &= " ,Q1C = @Q1C" & vbCrLf
        sql &= " ,Q7MR1 = @Q7MR1" & vbCrLf
        sql &= " ,Q7MR2 = @Q7MR2" & vbCrLf
        sql &= " ,Q7MR3 = @Q7MR3" & vbCrLf
        sql &= " ,Q7MR4 = @Q7MR4" & vbCrLf
        sql &= " ,Q211 = @Q211" & vbCrLf
        sql &= " ,Q212 = @Q212" & vbCrLf
        sql &= " ,Q213 = @Q213" & vbCrLf
        sql &= " ,Q214 = @Q214" & vbCrLf
        sql &= " ,Q215 = @Q215" & vbCrLf
        sql &= " ,Q216 = @Q216" & vbCrLf
        sql &= " ,Q217 = @Q217" & vbCrLf
        sql &= " ,Q218 = @Q218" & vbCrLf
        sql &= " ,Q221 = @Q221" & vbCrLf
        sql &= " ,Q222 = @Q222" & vbCrLf
        sql &= " ,Q223 = @Q223" & vbCrLf
        sql &= " ,Q224 = @Q224" & vbCrLf
        sql &= " ,Q225 = @Q225" & vbCrLf
        sql &= " ,Q226 = @Q226" & vbCrLf
        sql &= " ,Q3_NOTE = @Q3_NOTE " & vbCrLf
        sql &= " ,DASOURCE = @DASOURCE " & vbCrLf
        sql &= " ,MODIFYACCT = @MODIFYACCT " & vbCrLf
        sql &= " ,MODIFYDATE = GETDATE() " & vbCrLf
        sql &= " WHERE SOCID = @SOCID " & vbCrLf
        u_sql = sql
        Dim uCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " SELECT 'X' X FROM STUD_QUESTIONFIN WHERE SOCID = @SOCID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = vSOCID
            dt.Load(.ExecuteReader())
        End With

        'flagI:: 0:修改 1:新增
        Dim flagI As Integer = 0
        If dt.Rows.Count = 0 Then flagI = 1

        Dim s_FINDASOURCE As String = Get_FINDASOURCE(vSOCID, objconn)
        Dim flag_SAVE_Enabled As Boolean = If(s_FINDASOURCE = "1", False, True)
        If Not flag_SAVE_Enabled Then
            btnSave1.Enabled = flag_SAVE_Enabled
            'TIMS.Tooltip(btnSave1, "學員外網填寫，不可修改", True)
            Common.MessageBox(Me, "學員外網填寫，不可修改!!")
            Exit Sub
        End If

        Dim Q1value As String = ""
        Dim Q1a As String = ""
        Dim Q1b As String = ""
        Dim Q1c As String = ""
        Select Case True
            Case rblQ1_1.Checked '1.留任原公司
                Q1value = 1
            Case rblQ1_2.Checked '2.轉換至同產業的公司
                Q1value = 2
            Case rblQ1_3.Checked '3.轉換至不同產業的公司
                Q1value = 3
            Case rblQ1_4.Checked '4.創業
                Q1value = 4
            Case rblQ1_5.Checked '5.已離職，待業中
                Q1value = 5
            Case rblQ1_6.Checked '6.其他
                Q1value = 6
        End Select

        Select Case Q1abc.SelectedValue
            Case "a"
                Q1a = "Y"
            Case "b"
                Q1b = "Y"
            Case "c"
                Q1c = "Y"
        End Select

        Dim Q7MR1 As String = ""
        Dim Q7MR2 As String = ""
        Dim Q7MR3 As String = ""
        Dim Q7MR4 As String = ""
        For Each Items1 As ListItem In Q7.Items
            If Items1.Selected Then
                Select Case Items1.Value
                    Case "1"
                        Q7MR1 = "Y"
                    Case "2"
                        Q7MR2 = "Y"
                    Case "3"
                        Q7MR3 = "Y"
                    Case "4"
                        Q7MR4 = "Y"
                End Select
            End If
        Next

        Q3_Note.Text = TIMS.ClearSQM(Q3_Note.Text)

        'flagI:: 0:修改 1:新增
        If flagI = 1 Then
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("SOCID", SqlDbType.VarChar).Value = vSOCID
                .Parameters.Add("Q1", SqlDbType.VarChar).Value = TIMS.GetValue1(Q1value)
                .Parameters.Add("Q2", SqlDbType.VarChar).Value = TIMS.GetValue1(Q2.SelectedValue)
                .Parameters.Add("Q3", SqlDbType.VarChar).Value = TIMS.GetValue1(Q3.SelectedValue)
                .Parameters.Add("Q4", SqlDbType.VarChar).Value = TIMS.GetValue1(Q4.SelectedValue)
                .Parameters.Add("Q5", SqlDbType.VarChar).Value = TIMS.GetValue1(Q5.SelectedValue)
                .Parameters.Add("Q8", SqlDbType.VarChar).Value = TIMS.GetValue1(Q6.SelectedValue)
                .Parameters.Add("Q1A", SqlDbType.VarChar).Value = TIMS.GetValue1(Q1a)
                .Parameters.Add("Q1B", SqlDbType.VarChar).Value = TIMS.GetValue1(Q1b)
                .Parameters.Add("Q1C", SqlDbType.VarChar).Value = TIMS.GetValue1(Q1c)
                .Parameters.Add("Q7MR1", SqlDbType.VarChar).Value = TIMS.GetValue1(Q7MR1)
                .Parameters.Add("Q7MR2", SqlDbType.VarChar).Value = TIMS.GetValue1(Q7MR2)
                .Parameters.Add("Q7MR3", SqlDbType.VarChar).Value = TIMS.GetValue1(Q7MR3)
                .Parameters.Add("Q7MR4", SqlDbType.VarChar).Value = TIMS.GetValue1(Q7MR4)
                .Parameters.Add("Q211", SqlDbType.VarChar).Value = TIMS.GetValue1(Q211.SelectedValue)
                .Parameters.Add("Q212", SqlDbType.VarChar).Value = TIMS.GetValue1(Q212.SelectedValue)
                .Parameters.Add("Q213", SqlDbType.VarChar).Value = TIMS.GetValue1(Q213.SelectedValue)
                .Parameters.Add("Q214", SqlDbType.VarChar).Value = TIMS.GetValue1(Q214.SelectedValue)
                .Parameters.Add("Q215", SqlDbType.VarChar).Value = TIMS.GetValue1(Q215.SelectedValue)
                .Parameters.Add("Q216", SqlDbType.VarChar).Value = TIMS.GetValue1(Q216.SelectedValue)
                .Parameters.Add("Q217", SqlDbType.VarChar).Value = TIMS.GetValue1(Q217.SelectedValue)
                .Parameters.Add("Q218", SqlDbType.VarChar).Value = TIMS.GetValue1(Q218.SelectedValue)
                .Parameters.Add("Q221", SqlDbType.VarChar).Value = TIMS.GetValue1(Q221.SelectedValue)
                .Parameters.Add("Q222", SqlDbType.VarChar).Value = TIMS.GetValue1(Q222.SelectedValue)
                .Parameters.Add("Q223", SqlDbType.VarChar).Value = TIMS.GetValue1(Q223.SelectedValue)
                .Parameters.Add("Q224", SqlDbType.VarChar).Value = TIMS.GetValue1(Q224.SelectedValue)
                .Parameters.Add("Q225", SqlDbType.VarChar).Value = TIMS.GetValue1(Q225.SelectedValue)
                .Parameters.Add("Q226", SqlDbType.VarChar).Value = TIMS.GetValue1(Q226.SelectedValue)
                .Parameters.Add("Q3_NOTE", SqlDbType.NVarChar).Value = TIMS.GetValue1(Q3_Note.Text)
                .Parameters.Add("DASOURCE", SqlDbType.VarChar).Value = cst_DASOURCE '"2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sMODIFYACCT 'sm.UserInfo.UserID
                '.ExecuteNonQuery()
            End With
            DbAccess.ExecuteNonQuery(i_sql, objconn, iCmd.Parameters)
        Else
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("Q1", SqlDbType.VarChar).Value = TIMS.GetValue1(Q1value)
                .Parameters.Add("Q2", SqlDbType.VarChar).Value = TIMS.GetValue1(Q2.SelectedValue)
                .Parameters.Add("Q3", SqlDbType.VarChar).Value = TIMS.GetValue1(Q3.SelectedValue)
                .Parameters.Add("Q4", SqlDbType.VarChar).Value = TIMS.GetValue1(Q4.SelectedValue)
                .Parameters.Add("Q5", SqlDbType.VarChar).Value = TIMS.GetValue1(Q5.SelectedValue)
                .Parameters.Add("Q8", SqlDbType.VarChar).Value = TIMS.GetValue1(Q6.SelectedValue)
                .Parameters.Add("Q1A", SqlDbType.VarChar).Value = TIMS.GetValue1(Q1a)
                .Parameters.Add("Q1B", SqlDbType.VarChar).Value = TIMS.GetValue1(Q1b)
                .Parameters.Add("Q1C", SqlDbType.VarChar).Value = TIMS.GetValue1(Q1c)
                .Parameters.Add("Q7MR1", SqlDbType.VarChar).Value = TIMS.GetValue1(Q7MR1)
                .Parameters.Add("Q7MR2", SqlDbType.VarChar).Value = TIMS.GetValue1(Q7MR2)
                .Parameters.Add("Q7MR3", SqlDbType.VarChar).Value = TIMS.GetValue1(Q7MR3)
                .Parameters.Add("Q7MR4", SqlDbType.VarChar).Value = TIMS.GetValue1(Q7MR4)
                .Parameters.Add("Q211", SqlDbType.VarChar).Value = TIMS.GetValue1(Q211.SelectedValue)
                .Parameters.Add("Q212", SqlDbType.VarChar).Value = TIMS.GetValue1(Q212.SelectedValue)
                .Parameters.Add("Q213", SqlDbType.VarChar).Value = TIMS.GetValue1(Q213.SelectedValue)
                .Parameters.Add("Q214", SqlDbType.VarChar).Value = TIMS.GetValue1(Q214.SelectedValue)
                .Parameters.Add("Q215", SqlDbType.VarChar).Value = TIMS.GetValue1(Q215.SelectedValue)
                .Parameters.Add("Q216", SqlDbType.VarChar).Value = TIMS.GetValue1(Q216.SelectedValue)
                .Parameters.Add("Q217", SqlDbType.VarChar).Value = TIMS.GetValue1(Q217.SelectedValue)
                .Parameters.Add("Q218", SqlDbType.VarChar).Value = TIMS.GetValue1(Q218.SelectedValue)
                .Parameters.Add("Q221", SqlDbType.VarChar).Value = TIMS.GetValue1(Q221.SelectedValue)
                .Parameters.Add("Q222", SqlDbType.VarChar).Value = TIMS.GetValue1(Q222.SelectedValue)
                .Parameters.Add("Q223", SqlDbType.VarChar).Value = TIMS.GetValue1(Q223.SelectedValue)
                .Parameters.Add("Q224", SqlDbType.VarChar).Value = TIMS.GetValue1(Q224.SelectedValue)
                .Parameters.Add("Q225", SqlDbType.VarChar).Value = TIMS.GetValue1(Q225.SelectedValue)
                .Parameters.Add("Q226", SqlDbType.VarChar).Value = TIMS.GetValue1(Q226.SelectedValue)
                .Parameters.Add("Q3_NOTE", SqlDbType.NVarChar).Value = TIMS.GetValue1(Q3_Note.Text)
                .Parameters.Add("DASOURCE", SqlDbType.VarChar).Value = cst_DASOURCE '"2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sMODIFYACCT 'sm.UserInfo.UserID
                .Parameters.Add("SOCID", SqlDbType.VarChar).Value = vSOCID
                '.ExecuteNonQuery()
            End With
            DbAccess.ExecuteNonQuery(u_sql, objconn, uCmd.Parameters)
        End If
    End Sub

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave1.Click
        Call CHECK_SESS1()

        Dim v_ddlSOCID As String = TIMS.GetListValue(ddlSOCID)

        If v_ddlSOCID = "" Then
            Common.MessageBox(Me, "請選擇有效學員!!")
            Exit Sub
        End If

        Call SaveData1(v_ddlSOCID, objconn)
        Common.AddClientScript(Page, "insert_next('" & v_ddlSOCID & "');")

        'Try
        'Catch ex As Exception
        '    DbAccess.RollbackTrans(objTrans)
        '    Common.MessageBox(Me, ex.ToString)
        '    'objconn.Close()
        '    Throw ex
        'End Try
    End Sub

    Protected Sub btnClear1_Click(sender As Object, e As EventArgs) Handles btnClear1.Click
        Call sub_Clear1()
    End Sub
End Class