Public Class CP_01_001_add_06
    Inherits AuthBasePage

    'CLASS_VISITOR3
    'CLASS_VISITOR
    'VIEW_VISITOR(CLASS_VISITOR/CLASS_VISITOR3)
    'CP_01_001_03*.jrxml
    'CP_01_001_add_03*.jrxml
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    'Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            '新增 或修改 時利用 Session("SearchStr") 傳遞班級。
            If Session("SearchStr") IsNot Nothing Then
                ViewState("SearchStr") = Session("SearchStr")
                center.Text = TIMS.GetMyValue(Me.ViewState("SearchStr"), "center")
                RIDValue.Value = TIMS.GetMyValue(Me.ViewState("SearchStr"), "RIDValue")
                TMID1.Text = TIMS.GetMyValue(Me.ViewState("SearchStr"), "TMID1")
                OCID1.Text = TIMS.GetMyValue(Me.ViewState("SearchStr"), "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(Me.ViewState("SearchStr"), "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(Me.ViewState("SearchStr"), "OCIDValue1")
                EndDate.Value = TIMS.GetMyValue(Me.ViewState("SearchStr"), "end_date")
                '新增或修改都會傳遞此值。
                LabVISCOUNT.Text = TIMS.GetMyValue(Me.ViewState("SearchStr"), "VISCOUNT")
                Session("SearchStr") = Session("SearchStr")
            End If
            ViewState("_SearchStr") = Session("_SearchStr")
            Session("_SearchStr") = Session("_SearchStr")

            If OCIDValue1.Value <> "" Then create(OCIDValue1.Value, "")
            '修改動作。
            If Request("OCID") <> "" AndAlso Request("SeqNo") <> "" Then create(Request("OCID"), Request("SeqNo"))
            'If Request("DOCID") <> "" Then  create(Request("DOCID"), "")
        End If

        'Button1.Attributes("onclick") = "javascript:return chkdata()"

        '不提供儲存鈕。
        If Request("view") = "1" Then
            Button1.Visible = False
            TIMS.Tooltip(Button1, "僅供檢視")
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        HIDTPlanID.Value = sm.UserInfo.TPlanID
#Region "(No Use)"

        'D3c.Attributes("onclick") = "return check('" & D3c.ClientID & "','" & D3c2.ClientID & "')"
        'D4C.Attributes("onclick") = "return check('" & D4c.ClientID & "','" & D4c2.ClientID & "')"
        ''D5c.Attributes("onclick") = "return check('" & D5c.ClientID & "','" & D5c2.ClientID & "')"
        ''D6c.Attributes("onclick") = "return check('" & D6c.ClientID & "','" & D6c2.ClientID & "')"
        'D7c.Attributes("onclick") = "return check('" & D7c.ClientID & "','" & D7c2.ClientID & "')"
        ''D8c.Attributes("onclick") = "return check('" & D8c.ClientID & "','" & D8c2.ClientID & "')"
        'D3C1.Attributes("onclick") = "return check('" & D3c2.ClientID & "','" & D3c.ClientID & "')"
        'D4c2.Attributes("onclick") = "return check('" & D4c2.ClientID & "','" & D4c.ClientID & "')"
        ''D5c2.Attributes("onclick") = "return check('" & D5c2.ClientID & "','" & D5c.ClientID & "')"
        ''D6c2.Attributes("onclick") = "return check('" & D6c2.ClientID & "','" & D6c.ClientID & "')"
        'D7c2.Attributes("onclick") = "return check('" & D7c2.ClientID & "','" & D7c.ClientID & "')"
        ''D8c2.Attributes("onclick") = "return check('" & D8c2.ClientID & "','" & D8c.ClientID & "')"

#End Region
    End Sub

    '查詢班級。
    Sub create(ByVal OCID As String, ByVal SeqNo As String)

        '(檢視或修改)一定會有OCID
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        If drCC Is Nothing Then Return

        center.Text = drCC("OrgName").ToString()
        RIDValue.Value = drCC("RID").ToString()
        TMID1.Text = "[" & drCC("TrainID").ToString() & "]" & drCC("TrainName").ToString()
        TMIDValue1.Value = drCC("TMID").ToString()
        OCID1.Text = drCC("CLASSCNAME2").ToString() '第" & Val(drCC("CyclType")) & "期"
        OCIDValue1.Value = drCC("OCID")
        center.Enabled = False
        TMID1.Enabled = False
        OCID1.Enabled = False
        Button2.Disabled = True
        Button3.Disabled = True

        If SeqNo = "" Then Return

        Dim parms As New Hashtable
        parms.Add("OCID", OCID)
        If SeqNo <> "" Then parms.Add("SEQNO", SeqNo)

        Dim sql As String = ""
        sql = ""
        sql &= " SELECT * FROM CLASS_VISITOR3" & vbCrLf
        sql &= " WHERE 1=1 AND OCID =@OCID" & vbCrLf
        If SeqNo <> "" Then sql &= " AND SEQNO =@SEQNO" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)

        If dr Is Nothing Then Return

        APPLYDATE.Text = TIMS.Cdate3(dr("APPLYDATE")) '訪查日期
        APPLYDATEHH1.Text = Convert.ToString(dr("APPLYDATEHH1"))
        APPLYDATEMI1.Text = Convert.ToString(dr("APPLYDATEMI1"))
        APPLYDATEHH2.Text = Convert.ToString(dr("APPLYDATEHH2"))
        APPLYDATEMI2.Text = Convert.ToString(dr("APPLYDATEMI2"))
        VISTIMES.Text = Convert.ToString(dr("VISTIMES"))
        If Convert.ToString(dr("DATA1")) <> "" Then Common.SetListItem(DATA1, Convert.ToString(dr("DATA1")))
        If Convert.ToString(dr("DATA2")) <> "" Then Common.SetListItem(DATA2, Convert.ToString(dr("DATA2")))
        If Convert.ToString(dr("DATA3")) <> "" Then Common.SetListItem(DATA3, Convert.ToString(dr("DATA3")))
        'If Convert.ToString(dr("DATA4")) <> "" Then Common.SetListItem(DATA4, Convert.ToString(dr("DATA4")))
        'If Convert.ToString(dr("DATA5")) <> "" Then Common.SetListItem(DATA5, Convert.ToString(dr("DATA5")))
        'If Convert.ToString(dr("DATA6")) <> "" Then Common.SetListItem(DATA6, Convert.ToString(dr("DATA6")))
        DATACOPY1.Text = Convert.ToString(dr("DATACOPY1"))
        DATACOPY2.Text = Convert.ToString(dr("DATACOPY2"))
        DATACOPY3.Text = Convert.ToString(dr("DATACOPY3"))
        'DATACOPY4.Text = Convert.ToString(dr("DATACOPY4"))
        'DATACOPY5.Text = Convert.ToString(dr("DATACOPY5"))
        'DATACOPY6.Text = Convert.ToString(dr("DATACOPY6"))
        If Convert.ToString(dr("D3C")) <> "" Then
            Select Case Convert.ToString(dr("D3C"))
                Case "1"
                    D3C1.Checked = True
                Case "2"
                    D3C2.Checked = True
                Case "3"
                    D3C3.Checked = True
            End Select
        End If
        'If Convert.ToString(dr("D4C")) <> "" Then Common.SetListItem(D4C, Convert.ToString(dr("D4C")))
        'If Convert.ToString(dr("D5C")) <> "" Then Common.SetListItem(D5C, Convert.ToString(dr("D5C")))
        'If Convert.ToString(dr("D6C")) <> "" Then Common.SetListItem(D6C, Convert.ToString(dr("D6C")))
        'If Convert.ToString(dr("DATA62")) <> "" Then Common.SetListItem(DATA62, Convert.ToString(dr("DATA62")))
        'DATACOPY62.Text = Convert.ToString(dr("DATACOPY62"))
        'If Convert.ToString(dr("D62C")) <> "" Then Common.SetListItem(D62C, Convert.ToString(dr("D62C")))
        ITEM7NOTE2.Text = Convert.ToString(dr("ITEM7NOTE2"))
        D1CMM.Text = Convert.ToString(dr("D1CMM"))
        D1CDD.Text = Convert.ToString(dr("D1CDD"))
        D2CMM.Text = Convert.ToString(dr("D2CMM"))
        D2CDD.Text = Convert.ToString(dr("D2CDD"))
        D3CMM.Text = Convert.ToString(dr("D3CMM"))
        D3CDD.Text = Convert.ToString(dr("D3CDD"))
        D3NOTE.Text = Convert.ToString(dr("D3NOTE"))
        APPROVEDCOUNT.Text = Convert.ToString(dr("APPROVEDCOUNT"))
        AUTHCOUNT.Text = Convert.ToString(dr("AUTHCOUNT"))
        TURTHCOUNT.Text = Convert.ToString(dr("TURTHCOUNT"))
        TURNOUTCOUNT.Text = Convert.ToString(dr("TURNOUTCOUNT"))
        TRUANCYCOUNT.Text = Convert.ToString(dr("TRUANCYCOUNT"))
        LEAVECOUNT.Text = Convert.ToString(dr("LEAVECOUNT")) '離訓
        REJECTCOUNT.Text = Convert.ToString(dr("REJECTCOUNT")) '退訓
        'ADVJOBCOUNT.Text = Convert.ToString(dr("ADVJOBCOUNT"))'含提前就業

        If Convert.ToString(dr("ITEM1_1")) <> "" Then Common.SetListItem(ITEM1_1, Convert.ToString(dr("ITEM1_1")))
        If Convert.ToString(dr("ITEM1_2")) <> "" Then Common.SetListItem(ITEM1_2, Convert.ToString(dr("ITEM1_2")))
        ITEM1_COUR.Text = Convert.ToString(dr("ITEM1_COUR"))
        If Convert.ToString(dr("ITEM1_3")) <> "" Then Common.SetListItem(ITEM1_3, Convert.ToString(dr("ITEM1_3")))
        ITEM1_TEACHER.Text = Convert.ToString(dr("ITEM1_TEACHER"))
        ITEM1_ASSISTANT.Text = Convert.ToString(dr("ITEM1_ASSISTANT"))
        If Convert.ToString(dr("ITEM2_1")) <> "" Then Common.SetListItem(ITEM2_1, Convert.ToString(dr("ITEM2_1")))
        If Convert.ToString(dr("ITEM2_2")) <> "" Then Common.SetListItem(ITEM2_2, Convert.ToString(dr("ITEM2_2")))
        'If Convert.ToString(dr("ITEM2_3")) <> "" Then Common.SetListItem(ITEM2_3, Convert.ToString(dr("ITEM2_3")))
        If Convert.ToString(dr("ITEM3_1")) <> "" Then Common.SetListItem(ITEM3_1, Convert.ToString(dr("ITEM3_1")))
        If Convert.ToString(dr("ITEM3_2")) <> "" Then Common.SetListItem(ITEM3_2, Convert.ToString(dr("ITEM3_2")))
        'If Convert.ToString(dr("ITEM3_3")) <> "" Then Common.SetListItem(ITEM3_3, Convert.ToString(dr("ITEM3_3")))
        'If Convert.ToString(dr("ITEM3_4")) <> "" Then Common.SetListItem(ITEM3_4, Convert.ToString(dr("ITEM3_4")))
        'If Convert.ToString(dr("ITEM3_5")) <> "" Then Common.SetListItem(ITEM3_5, Convert.ToString(dr("ITEM3_5")))
        'If Convert.ToString(dr("ITEM4_1")) <> "" Then Common.SetListItem(ITEM4_1, Convert.ToString(dr("ITEM4_1")))
        'If Convert.ToString(dr("ITEM4_2")) <> "" Then Common.SetListItem(ITEM4_2, Convert.ToString(dr("ITEM4_2")))
        'ITEM4NOTE.Text = Convert.ToString(dr("ITEM4NOTE"))
        'If Convert.ToString(dr("ITEM4_3")) <> "" Then Common.SetListItem(ITEM4_3, Convert.ToString(dr("ITEM4_3")))
        ITEM7NOTE.Text = Convert.ToString(dr("ITEM7NOTE"))

        ITEM1PROS.Text = Convert.ToString(dr("ITEM1PROS"))
        ITEM2PROS.Text = Convert.ToString(dr("ITEM2PROS"))
        ITEM3PROS.Text = Convert.ToString(dr("ITEM3PROS"))
        'ITEM4PROS.Text = Convert.ToString(dr("ITEM4PROS"))
        ITEM1NOTE.Text = Convert.ToString(dr("ITEM1NOTE"))
        ITEM2NOTE.Text = Convert.ToString(dr("ITEM2NOTE"))
        ITEM3NOTE.Text = Convert.ToString(dr("ITEM3NOTE"))

        '訪查單位綜合建議
        ITEM31NOTE.Text = Convert.ToString(dr("ITEM31NOTE"))
        '訓練單位缺失處理
        If Convert.ToString(dr("ITEM32")) <> "" Then Common.SetListItem(ITEM32, Convert.ToString(dr("ITEM32")))
        ITEM32NOTE.Text = Convert.ToString(dr("ITEM32NOTE"))

        'CURSENAME.Text = Convert.ToString(dr("CURSENAME"))
        VISITORNAME.Text = Convert.ToString(dr("VISITORNAME"))

    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""
        If VISTIMES.Text <> "" Then VISTIMES.Text = Trim(VISTIMES.Text)
        If VISTIMES.Text = "" Then Errmsg += "請填 本次為第次訪視！ " & vbCrLf
        If APPLYDATE.Text <> "" Then APPLYDATE.Text = Trim(APPLYDATE.Text)
        If APPLYDATE.Text = "" Then
            Errmsg += "請填選查訪日期！ " & vbCrLf
        Else
            If Not TIMS.IsDate1(APPLYDATE.Text) Then Errmsg += "請填選正確查訪日期！ " & vbCrLf
        End If
        If Request("OCID") = "" Then '表示新增狀態
            If OCIDValue1.Value = "" Then Errmsg += "班別代碼有誤，請確認點選職類/班別！ " & vbCrLf
            If OCIDValue1.Value <> "" Then
                Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
                If drCC Is Nothing Then Errmsg += "班別代碼有誤，請確認點選職類/班別！ " & vbCrLf
            End If
        Else
            If OCIDValue1.Value <> Request("OCID") Then Errmsg += "班別代碼有誤，修改模式下不可重新選職類/班別！ " & vbCrLf
            If OCIDValue1.Value = Request("OCID") Then
                Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
                If drCC Is Nothing Then Errmsg += "班別代碼有誤，請確認點選職類/班別！ " & vbCrLf
            End If
        End If

        APPROVEDCOUNT.Text = TIMS.ClearSQM(APPROVEDCOUNT.Text)
        AUTHCOUNT.Text = TIMS.ClearSQM(AUTHCOUNT.Text)
        TURTHCOUNT.Text = TIMS.ClearSQM(TURTHCOUNT.Text)
        TURNOUTCOUNT.Text = TIMS.ClearSQM(TURNOUTCOUNT.Text)
        TRUANCYCOUNT.Text = TIMS.ClearSQM(TRUANCYCOUNT.Text)
        'LEAVECOUNT. 離訓
        LEAVECOUNT.Text = TIMS.ClearSQM(LEAVECOUNT.Text)
        REJECTCOUNT.Text = TIMS.ClearSQM(REJECTCOUNT.Text)

        'If ADVJOBCOUNT.Text <> "" Then ADVJOBCOUNT.Text = Trim(ADVJOBCOUNT.Text)
        If DATA1.SelectedValue = "" Then Errmsg += "請選擇 教學(訓練)日誌選項！ " & vbCrLf
        If DATA2.SelectedValue = "" Then Errmsg += "請選擇 學員簽到(退)表 選項！ " & vbCrLf
        If DATA3.SelectedValue = "" Then Errmsg += "請選擇 請假單 選項！ " & vbCrLf
        'If DATA4.SelectedValue = "" Then Errmsg += "請選擇 退訓／提前就業申請表 選項！ " & vbCrLf
        'If DATA5.SelectedValue = "" Then Errmsg += "請選擇 職業訓練生活津貼補助印領清冊 選項！ " & vbCrLf
        'If DATA6.SelectedValue = "" Then Errmsg += "請選擇 參訓學員辦理勞工保險加退保紀錄 選項！ " & vbCrLf
        'If DATA7.SelectedValue = "" Then Errmsg += "請選擇 出缺勤狀況 選項！ " & vbCrLf
        If DATACOPY1.Text <> "" Then DATACOPY1.Text = Trim(DATACOPY1.Text)
        If DATACOPY2.Text <> "" Then DATACOPY2.Text = Trim(DATACOPY2.Text)
        If DATACOPY3.Text <> "" Then DATACOPY3.Text = Trim(DATACOPY3.Text)
        'If DATACOPY4.Text <> "" Then DATACOPY4.Text = Trim(DATACOPY4.Text)
        'If DATACOPY5.Text <> "" Then DATACOPY5.Text = Trim(DATACOPY5.Text)
        'If DATACOPY6.Text <> "" Then DATACOPY6.Text = Trim(DATACOPY6.Text)
        If D1CMM.Text <> "" Then D1CMM.Text = TIMS.ClearSQM(D1CMM.Text)
        If D1CDD.Text <> "" Then D1CDD.Text = TIMS.ClearSQM(D1CDD.Text)
        If D2CMM.Text <> "" Then D2CMM.Text = TIMS.ClearSQM(D2CMM.Text)
        If D2CDD.Text <> "" Then D2CDD.Text = TIMS.ClearSQM(D2CDD.Text)
        If D3CMM.Text <> "" Then D3CMM.Text = TIMS.ClearSQM(D3CMM.Text)
        If D3CDD.Text <> "" Then D3CDD.Text = TIMS.ClearSQM(D3CDD.Text)
        Dim iD3C As Integer = 0
        If D3C1.Checked Then iD3C = 1
        If D3C2.Checked Then iD3C = 2
        If D3C3.Checked Then iD3C = 3
        If iD3C = 0 Then Errmsg += "請選擇 請假單 備註 選項！ " & vbCrLf
        'If D4C.SelectedValue = "" Then Errmsg += "請選擇 退訓／提前就業申請表 備註 選項！ " & vbCrLf
        'If D5C.SelectedValue = "" Then Errmsg += "請選擇 職業訓練生活津貼補助印領清冊 備註 選項！ " & vbCrLf
        'If D5C.SelectedValue = "" Then Errmsg += "請選擇 參訓學員辦理勞工保險加退保紀錄 備註 選項！ " & vbCrLf
        If Not TIMS.Check123(APPROVEDCOUNT.Text) Then Errmsg += "核定人數 必須為數字！ " & vbCrLf
        If Not TIMS.Check123(AUTHCOUNT.Text) Then Errmsg += "開訓人數 必須為數字！ " & vbCrLf
        If Not TIMS.Check123(TURTHCOUNT.Text) Then Errmsg += "實到人數 必須為數字！ " & vbCrLf
        If Not TIMS.Check123(TURNOUTCOUNT.Text) Then Errmsg += "請假人數 必須為數字！ " & vbCrLf
        If Not TIMS.Check123(TRUANCYCOUNT.Text) Then Errmsg += "缺(曠)課人數 必須為數字！ " & vbCrLf
        'LEAVECOUNT. 離訓
        If Not TIMS.Check123(LEAVECOUNT.Text) Then Errmsg += "離訓人數 必須為數字！ " & vbCrLf
        If Not TIMS.Check123(REJECTCOUNT.Text) Then Errmsg += "退訓人數 必須為數字！ " & vbCrLf
        'If Not TIMS.Check123(ADVJOBCOUNT.Text) Then Errmsg += "提前就業 必須為數字！ " & vbCrLf

        If ITEM1_1.SelectedValue = "" Then Errmsg += "請回答有無週(月)課程表? " & vbCrLf
        If ITEM1_2.SelectedValue = "" Then Errmsg += "請回答是否依課程表授課? " & vbCrLf
        If ITEM1_COUR.Text <> "" Then ITEM1_COUR.Text = Trim(ITEM1_COUR.Text)
        If ITEM1_COUR.Text = "" Then Errmsg += "請輸入課目或課題為何? " & vbCrLf
        If ITEM1_3.SelectedValue = "" Then Errmsg += "請回答教師與助教是否與計畫相符? " & vbCrLf
        If ITEM1_TEACHER.Text <> "" Then ITEM1_TEACHER.Text = Trim(ITEM1_TEACHER.Text)
        If ITEM1_TEACHER.Text = "" Then Errmsg += "請輸入教師姓名? " & vbCrLf
        If ITEM1_ASSISTANT.Text <> "" Then ITEM1_ASSISTANT.Text = Trim(ITEM1_ASSISTANT.Text)
        If ITEM2_1.SelectedValue = "" Then Errmsg += "請回答有無書籍(講義)領用表? " & vbCrLf
        If ITEM2_2.SelectedValue = "" Then Errmsg += "請回答有無材料領用表? " & vbCrLf
        'If ITEM2_3.SelectedValue = "" Then Errmsg += "請回答訓練設施設備是否依契約提供學員使用? " & vbCrLf
        If ITEM3_1.SelectedValue = "" Then Errmsg += "請回答教學(訓練)日誌是否確實填寫? " & vbCrLf
        If ITEM3_2.SelectedValue = "" Then Errmsg += "請回答有否按時呈主管核閱? " & vbCrLf
        'If ITEM3_3.SelectedValue = "" Then Errmsg += "請回答學員生活、就業輔導與管理機制是否依契約規範辦理? " & vbCrLf
        'If ITEM3_4.SelectedValue = "" Then Errmsg += "請回答是否依契約規範提供學員問題反應申訴管道? " & vbCrLf
        'If ITEM3_5.SelectedValue = "" Then Errmsg += "請回答是否依契約規範公告學員權益義務或編製參訓學員服務手冊? " & vbCrLf
        'If ITEM4_1.SelectedValue = "" Then Errmsg += "請回答是否依規定於開訓後15日內收齊職業訓練生活津貼申請書及相關證明文件後送委訓單位審查? " & vbCrLf
        'If ITEM4_2.SelectedValue = "" Then Errmsg += "請回答培訓單位於收到本署所屬分署核撥之津貼後，是否按月即時（不超過3個工作日）轉發給受訓學員? " & vbCrLf
        'If ITEM4_3.SelectedValue = "" Then Errmsg += "請回答申請人離、退訓時，培訓單位是否按月覈實繳回職業訓練生活津貼? " & vbCrLf
        'If ITEM4NOTE.Text <> "" Then ITEM4NOTE.Text = Trim(ITEM4NOTE.Text)
        'If ITEM4_2.SelectedValue = "3" Then
        '    If ITEM4NOTE.Text = "" Then Errmsg += "請回答 費用(津貼)收核狀況 免填原因說明? " & vbCrLf
        'End If
        If ITEM32.SelectedValue = "" Then Errmsg += "請選擇 訓練單位缺失處理? " & vbCrLf

        'If CURSENAME.Text <> "" Then CURSENAME.Text = Trim(CURSENAME.Text)
        If VISITORNAME.Text <> "" Then VISITORNAME.Text = Trim(VISITORNAME.Text)
        'If CURSENAME.Text = "" Then Errmsg += "請輸入培訓單位人員姓名? " & vbCrLf
        If VISITORNAME.Text = "" Then Errmsg += "請輸入訪視人員姓名? " & vbCrLf
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Dim rOCID As String = Request("OCID")
        Dim rSeqNo As String = Request("SeqNo")
        rOCID = TIMS.ClearSQM(rOCID)
        rSeqNo = TIMS.ClearSQM(rSeqNo)
        If rOCID <> "" Then
            '表示修改依傳入序號。
            If rSeqNo = "" Then Exit Sub
            Dim iSEQNO As Integer = rSeqNo
            Call UPDATE_DATA(iSEQNO)
        End If
        If rOCID = "" Then Call INSERT_DATA() '表示新增狀態

        Session("SearchStr") = If(Session("SearchStr") IsNot Nothing, Session("SearchStr"), ViewState("SearchStr"))
        Session("_SearchStr") = If(Session("_SearchStr") IsNot Nothing, Session("_SearchStr"), ViewState("_SearchStr"))

        Common.RespWrite(Me, "<script> alert('儲存成功');")
        If Request("DOCID") <> "" Then
            Common.RespWrite(Me, "location.href='CP_01_001.aspx?ID=" & Request("ID") & "&DOCID=" & Request("DOCID") & "';</script>")
        Else
            Common.RespWrite(Me, "location.href='CP_01_001.aspx?ID=" & Request("ID") & "';</script>")
        End If
    End Sub

    '回查詢頁面。
    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Session("SearchStr") = If(Session("SearchStr") IsNot Nothing, Session("SearchStr"), ViewState("SearchStr"))
        Session("_SearchStr") = If(Session("_SearchStr") IsNot Nothing, Session("_SearchStr"), ViewState("_SearchStr"))
        'If Request("DOCID") <> "" Then
        '    TIMS.Utl_Redirect1(Me, "CP_01_001.aspx?ID=" & Request("ID") & "&DOCID=" & Request("DOCID"))
        'Else
        '    TIMS.Utl_Redirect1(Me, "CP_01_001.aspx?ID=" & Request("ID"))
        'End If
        Dim url1 As String = ""
        url1 = "CP_01_001.aspx?ID=" & Request("ID")
        If TIMS.ClearSQM(Request("DOCID")) <> "" Then url1 = "CP_01_001.aspx?ID=" & Request("ID") & "&DOCID=" & Request("DOCID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    'INPUT OCIDValue1.Value OUTPUT: MAX(SEQNO)+1(NEW)
    Function sUtl_GetSEQNO() As Integer
        Dim rst As Integer = 1
        Dim sql As String = ""
        sql = " SELECT MAX(SEQNO) SEQNO FROM CLASS_VISITOR3 WHERE OCID = @OCID "
        Dim dtS As New DataTable
        Dim sCMD As New SqlCommand(sql, objconn)
        With sCMD
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCIDValue1.Value)
            dtS.Load(.ExecuteReader())
        End With
        If dtS.Rows.Count > 0 Then
            If Convert.ToString(dtS.Rows(0)("SEQNO")) <> "" Then rst = Val(dtS.Rows(0)("SEQNO")) + 1
        End If
        Return rst
    End Function

    '新增
    Sub INSERT_DATA()
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql = ""
        sql &= " INSERT INTO CLASS_VISITOR3 (" & vbCrLf
        sql &= " OCID ,SEQNO" & vbCrLf
        sql &= " ,APPLYDATE" & vbCrLf
        sql &= " ,APPLYDATEHH1" & vbCrLf
        sql &= " ,APPLYDATEMI1" & vbCrLf
        sql &= " ,APPLYDATEHH2" & vbCrLf
        sql &= " ,APPLYDATEMI2" & vbCrLf
        sql &= " ,VISTIMES" & vbCrLf
        sql &= " ,DATA1" & vbCrLf
        sql &= " ,DATA2" & vbCrLf
        sql &= " ,DATA3" & vbCrLf
        sql &= " ,DATACOPY1" & vbCrLf
        sql &= " ,DATACOPY2" & vbCrLf
        sql &= " ,DATACOPY3" & vbCrLf
        sql &= " ,D3C" & vbCrLf
        sql &= " ,ITEM7NOTE2" & vbCrLf
        sql &= " ,D1CMM" & vbCrLf
        sql &= " ,D1CDD" & vbCrLf
        sql &= " ,D2CMM" & vbCrLf
        sql &= " ,D2CDD" & vbCrLf
        sql &= " ,D3CMM" & vbCrLf
        sql &= " ,D3CDD" & vbCrLf
        sql &= " ,D3NOTE" & vbCrLf
        sql &= " ,APPROVEDCOUNT" & vbCrLf
        sql &= " ,AUTHCOUNT" & vbCrLf
        sql &= " ,TURTHCOUNT" & vbCrLf
        sql &= " ,TURNOUTCOUNT" & vbCrLf
        sql &= " ,TRUANCYCOUNT" & vbCrLf
        'LEAVECOUNT. 離訓
        sql &= " ,LEAVECOUNT" & vbCrLf
        sql &= " ,REJECTCOUNT" & vbCrLf
        'sql &= "   ,ADVJOBCOUNT" & vbCrLf
        sql &= " ,ITEM1_1" & vbCrLf
        sql &= " ,ITEM1_2" & vbCrLf
        sql &= " ,ITEM1_COUR" & vbCrLf
        sql &= " ,ITEM1_3" & vbCrLf
        sql &= " ,ITEM1_TEACHER" & vbCrLf
        sql &= " ,ITEM1_ASSISTANT" & vbCrLf
        sql &= " ,ITEM2_1" & vbCrLf
        sql &= " ,ITEM2_2" & vbCrLf
        'sql &= "   ,ITEM2_3" & vbCrLf
        sql &= " ,ITEM3_1" & vbCrLf
        sql &= " ,ITEM3_2" & vbCrLf
        sql &= " ,ITEM7NOTE" & vbCrLf
        sql &= " ,ITEM1PROS" & vbCrLf
        sql &= " ,ITEM2PROS" & vbCrLf
        sql &= " ,ITEM3PROS" & vbCrLf
        sql &= " ,ITEM1NOTE" & vbCrLf
        sql &= " ,ITEM2NOTE" & vbCrLf
        sql &= " ,ITEM3NOTE" & vbCrLf
        sql &= " ,ITEM31NOTE,ITEM32,ITEM32NOTE" & vbCrLf
        'sql &= " ,CURSENAME" & vbCrLf
        sql &= " ,VISITORNAME" & vbCrLf
        sql &= " ,RID" & vbCrLf
        sql &= " ,MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE" & vbCrLf
        sql &= " ) VALUES (" & vbCrLf
        sql &= " @OCID ,@SEQNO" & vbCrLf
        sql &= " ,@APPLYDATE" & vbCrLf
        sql &= " ,@APPLYDATEHH1" & vbCrLf
        sql &= " ,@APPLYDATEMI1" & vbCrLf
        sql &= " ,@APPLYDATEHH2" & vbCrLf
        sql &= " ,@APPLYDATEMI2" & vbCrLf
        sql &= " ,@VISTIMES" & vbCrLf
        sql &= " ,@DATA1" & vbCrLf
        sql &= " ,@DATA2" & vbCrLf
        sql &= " ,@DATA3" & vbCrLf
        sql &= " ,@DATACOPY1" & vbCrLf
        sql &= " ,@DATACOPY2" & vbCrLf
        sql &= " ,@DATACOPY3" & vbCrLf
        sql &= " ,@D3C" & vbCrLf
        sql &= " ,@ITEM7NOTE2" & vbCrLf
        sql &= " ,@D1CMM" & vbCrLf
        sql &= " ,@D1CDD" & vbCrLf
        sql &= " ,@D2CMM" & vbCrLf
        sql &= " ,@D2CDD" & vbCrLf
        sql &= " ,@D3CMM" & vbCrLf
        sql &= " ,@D3CDD" & vbCrLf
        sql &= " ,@D3NOTE" & vbCrLf
        sql &= " ,@APPROVEDCOUNT" & vbCrLf
        sql &= " ,@AUTHCOUNT" & vbCrLf
        sql &= " ,@TURTHCOUNT" & vbCrLf
        sql &= " ,@TURNOUTCOUNT" & vbCrLf
        sql &= " ,@TRUANCYCOUNT" & vbCrLf
        'LEAVECOUNT. 離訓
        sql &= " ,@LEAVECOUNT" & vbCrLf
        sql &= " ,@REJECTCOUNT" & vbCrLf
        'sql &= "   ,@ADVJOBCOUNT" & vbCrLf
        sql &= " ,@ITEM1_1" & vbCrLf
        sql &= " ,@ITEM1_2" & vbCrLf
        sql &= " ,@ITEM1_COUR" & vbCrLf
        sql &= " ,@ITEM1_3" & vbCrLf
        sql &= " ,@ITEM1_TEACHER" & vbCrLf
        sql &= " ,@ITEM1_ASSISTANT" & vbCrLf
        sql &= " ,@ITEM2_1" & vbCrLf
        sql &= " ,@ITEM2_2" & vbCrLf
        'sql &= "   ,@ITEM2_3" & vbCrLf
        sql &= " ,@ITEM3_1" & vbCrLf
        sql &= " ,@ITEM3_2" & vbCrLf
        sql &= " ,@ITEM7NOTE" & vbCrLf
        sql &= " ,@ITEM1PROS" & vbCrLf
        sql &= " ,@ITEM2PROS" & vbCrLf
        sql &= " ,@ITEM3PROS" & vbCrLf
        sql &= " ,@ITEM1NOTE" & vbCrLf
        sql &= " ,@ITEM2NOTE" & vbCrLf
        sql &= " ,@ITEM3NOTE" & vbCrLf
        sql &= " ,@ITEM31NOTE,@ITEM32,@ITEM32NOTE" & vbCrLf
        'sql &= " ,@CURSENAME" & vbCrLf
        sql &= " ,@VISITORNAME" & vbCrLf
        sql &= " ,@RID" & vbCrLf
        sql &= " ,@MODIFYACCT" & vbCrLf
        sql &= " ,GETDATE()" & vbCrLf
        sql &= " )" & vbCrLf
        Dim iCMD As New SqlCommand(sql, objconn)

        Dim iSEQNO As Integer = sUtl_GetSEQNO()
        Dim iD3C As Integer = If(D3C1.Checked, 1, If(D3C2.Checked, 2, If(D3C3.Checked, 3, 0)))
        'Dim dt As New DataTable
        'Dim oCmd As New SqlCommand(sql, objconn)
        With iCMD
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCIDValue1.Value)
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = iSEQNO

            .Parameters.Add("APPLYDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(APPLYDATE.Text)
            .Parameters.Add("APPLYDATEHH1", SqlDbType.VarChar).Value = APPLYDATEHH1.Text
            .Parameters.Add("APPLYDATEMI1", SqlDbType.VarChar).Value = APPLYDATEMI1.Text
            .Parameters.Add("APPLYDATEHH2", SqlDbType.VarChar).Value = APPLYDATEHH2.Text
            .Parameters.Add("APPLYDATEMI2", SqlDbType.VarChar).Value = APPLYDATEMI2.Text
            .Parameters.Add("VISTIMES", SqlDbType.Int).Value = Val(VISTIMES.Text)
            .Parameters.Add("DATA1", SqlDbType.VarChar).Value = If(DATA1.SelectedValue = "", Convert.DBNull, DATA1.SelectedValue)
            .Parameters.Add("DATA2", SqlDbType.VarChar).Value = If(DATA2.SelectedValue = "", Convert.DBNull, DATA2.SelectedValue)
            .Parameters.Add("DATA3", SqlDbType.VarChar).Value = If(DATA3.SelectedValue = "", Convert.DBNull, DATA3.SelectedValue)
            .Parameters.Add("DATACOPY1", SqlDbType.VarChar).Value = DATACOPY1.Text
            .Parameters.Add("DATACOPY2", SqlDbType.VarChar).Value = DATACOPY2.Text
            .Parameters.Add("DATACOPY3", SqlDbType.VarChar).Value = DATACOPY3.Text
            '.Parameters.Add("DATACOPY4", SqlDbType.VarChar).Value = DATACOPY4.Text
            '.Parameters.Add("DATACOPY5", SqlDbType.VarChar).Value = DATACOPY5.Text
            '.Parameters.Add("DATACOPY6", SqlDbType.VarChar).Value = DATACOPY6.Text
            '.Parameters.Add("D1C", SqlDbType.VarChar).Value = D1C
            '.Parameters.Add("D2C", SqlDbType.VarChar).Value = D2C
            .Parameters.Add("D3C", SqlDbType.Int).Value = If(iD3C = 0, Convert.DBNull, iD3C)
            .Parameters.Add("ITEM7NOTE2", SqlDbType.VarChar).Value = If(ITEM7NOTE2.Text = "", Convert.DBNull, ITEM7NOTE2.Text)
            .Parameters.Add("D1CMM", SqlDbType.VarChar).Value = If(D1CMM.Text = "", Convert.DBNull, D1CMM.Text)
            .Parameters.Add("D1CDD", SqlDbType.VarChar).Value = If(D1CDD.Text = "", Convert.DBNull, D1CDD.Text)
            .Parameters.Add("D2CMM", SqlDbType.VarChar).Value = If(D2CMM.Text = "", Convert.DBNull, D2CMM.Text)
            .Parameters.Add("D2CDD", SqlDbType.VarChar).Value = If(D2CDD.Text = "", Convert.DBNull, D2CDD.Text)
            .Parameters.Add("D3CMM", SqlDbType.VarChar).Value = If(D3CMM.Text = "", Convert.DBNull, D3CMM.Text)
            .Parameters.Add("D3CDD", SqlDbType.VarChar).Value = If(D3CDD.Text = "", Convert.DBNull, D3CDD.Text)
            '.Parameters.Add("D1NOTE", SqlDbType.VarChar).Value = D1NOTE.text
            '.Parameters.Add("D2NOTE", SqlDbType.VarChar).Value = D2NOTE.text
            .Parameters.Add("D3NOTE", SqlDbType.VarChar).Value = D3NOTE.Text
            '.Parameters.Add("D5NOTE", SqlDbType.VarChar).Value = D5NOTE.text
            '.Parameters.Add("D6NOTE", SqlDbType.VarChar).Value = D6NOTE.text
            .Parameters.Add("APPROVEDCOUNT", SqlDbType.Int).Value = Val(APPROVEDCOUNT.Text)
            .Parameters.Add("AUTHCOUNT", SqlDbType.Int).Value = Val(AUTHCOUNT.Text)
            .Parameters.Add("TURTHCOUNT", SqlDbType.Int).Value = Val(TURTHCOUNT.Text)
            .Parameters.Add("TURNOUTCOUNT", SqlDbType.Int).Value = Val(TURNOUTCOUNT.Text)
            .Parameters.Add("TRUANCYCOUNT", SqlDbType.Int).Value = Val(TRUANCYCOUNT.Text)
            'LEAVECOUNT. 離訓
            .Parameters.Add("LEAVECOUNT", SqlDbType.Int).Value = Val(LEAVECOUNT.Text)
            .Parameters.Add("REJECTCOUNT", SqlDbType.Int).Value = Val(REJECTCOUNT.Text)
            '.Parameters.Add("ADVJOBCOUNT", SqlDbType.Int).Value = Val(ADVJOBCOUNT.Text)
            .Parameters.Add("ITEM1_1", SqlDbType.VarChar).Value = If(ITEM1_1.SelectedValue = "", Convert.DBNull, ITEM1_1.SelectedValue)
            .Parameters.Add("ITEM1_2", SqlDbType.VarChar).Value = If(ITEM1_2.SelectedValue = "", Convert.DBNull, ITEM1_2.SelectedValue)
            .Parameters.Add("ITEM1_COUR", SqlDbType.VarChar).Value = ITEM1_COUR.Text
            .Parameters.Add("ITEM1_3", SqlDbType.VarChar).Value = If(ITEM1_3.SelectedValue = "", Convert.DBNull, ITEM1_3.SelectedValue)
            .Parameters.Add("ITEM1_TEACHER", SqlDbType.VarChar).Value = ITEM1_TEACHER.Text
            .Parameters.Add("ITEM1_ASSISTANT", SqlDbType.VarChar).Value = ITEM1_ASSISTANT.Text
            .Parameters.Add("ITEM2_1", SqlDbType.VarChar).Value = If(ITEM2_1.SelectedValue = "", Convert.DBNull, ITEM2_1.SelectedValue)
            .Parameters.Add("ITEM2_2", SqlDbType.VarChar).Value = If(ITEM2_2.SelectedValue = "", Convert.DBNull, ITEM2_2.SelectedValue)
            '.Parameters.Add("ITEM2_3", SqlDbType.VarChar).Value = If(ITEM2_3.SelectedValue = "", Convert.DBNull, ITEM2_3.SelectedValue)
            .Parameters.Add("ITEM3_1", SqlDbType.VarChar).Value = If(ITEM3_1.SelectedValue = "", Convert.DBNull, ITEM3_1.SelectedValue)
            .Parameters.Add("ITEM3_2", SqlDbType.VarChar).Value = If(ITEM3_2.SelectedValue = "", Convert.DBNull, ITEM3_2.SelectedValue)
            .Parameters.Add("ITEM7NOTE", SqlDbType.VarChar).Value = ITEM7NOTE.Text
            .Parameters.Add("ITEM1PROS", SqlDbType.VarChar).Value = ITEM1PROS.Text
            .Parameters.Add("ITEM2PROS", SqlDbType.VarChar).Value = ITEM2PROS.Text
            .Parameters.Add("ITEM3PROS", SqlDbType.VarChar).Value = ITEM3PROS.Text
            .Parameters.Add("ITEM1NOTE", SqlDbType.VarChar).Value = ITEM1NOTE.Text
            .Parameters.Add("ITEM2NOTE", SqlDbType.VarChar).Value = ITEM2NOTE.Text
            .Parameters.Add("ITEM3NOTE", SqlDbType.VarChar).Value = ITEM3NOTE.Text
            .Parameters.Add("ITEM31NOTE", SqlDbType.VarChar).Value = ITEM31NOTE.Text
            .Parameters.Add("ITEM32", SqlDbType.VarChar).Value = If(ITEM32.SelectedValue = "", Convert.DBNull, ITEM32.SelectedValue)
            .Parameters.Add("ITEM32NOTE", SqlDbType.VarChar).Value = ITEM32NOTE.Text

            '.Parameters.Add("CURSENAME", SqlDbType.VarChar).Value = CURSENAME.Text '培訓單位人員姓名
            .Parameters.Add("VISITORNAME", SqlDbType.VarChar).Value = VISITORNAME.Text '訪視人員姓名
            .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
            'dt.Load(.ExecuteReader())
            '.ExecuteNonQuery()  'edit，by:20181011
            DbAccess.ExecuteNonQuery(iCMD.CommandText, objconn, iCMD.Parameters)  'edit，by:20181011
            'rst = .ExecuteScalar()
        End With
    End Sub

    '修改
    Sub UPDATE_DATA(ByVal iSEQNO As Integer)
        'CLASS_VISITOR3
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql = ""
        sql &= " UPDATE CLASS_VISITOR3 " & vbCrLf
        sql &= " SET APPLYDATE = @APPLYDATE " & vbCrLf
        sql &= " ,APPLYDATEHH1 = @APPLYDATEHH1 " & vbCrLf
        sql &= " ,APPLYDATEMI1 = @APPLYDATEMI1 " & vbCrLf
        sql &= " ,APPLYDATEHH2 = @APPLYDATEHH2 " & vbCrLf
        sql &= " ,APPLYDATEMI2 = @APPLYDATEMI2 " & vbCrLf
        sql &= " ,VISTIMES = @VISTIMES " & vbCrLf
        sql &= " ,DATA1 = @DATA1 " & vbCrLf
        sql &= " ,DATA2 = @DATA2 " & vbCrLf
        sql &= " ,DATA3 = @DATA3 " & vbCrLf
        sql &= " ,DATACOPY1 = @DATACOPY1 " & vbCrLf
        sql &= " ,DATACOPY2 = @DATACOPY2 " & vbCrLf
        sql &= " ,DATACOPY3 = @DATACOPY3 " & vbCrLf
        sql &= " ,D3C = @D3C " & vbCrLf
        sql &= " ,ITEM7NOTE2 = @ITEM7NOTE2 " & vbCrLf
        sql &= " ,D1CMM = @D1CMM " & vbCrLf
        sql &= " ,D1CDD = @D1CDD " & vbCrLf
        sql &= " ,D2CMM = @D2CMM " & vbCrLf
        sql &= " ,D2CDD = @D2CDD " & vbCrLf
        sql &= " ,D3CMM = @D3CMM " & vbCrLf
        sql &= " ,D3CDD = @D3CDD " & vbCrLf
        sql &= " ,D3NOTE = @D3NOTE " & vbCrLf
        sql &= " ,APPROVEDCOUNT = @APPROVEDCOUNT " & vbCrLf
        sql &= " ,AUTHCOUNT = @AUTHCOUNT " & vbCrLf
        sql &= " ,TURTHCOUNT = @TURTHCOUNT " & vbCrLf
        sql &= " ,TURNOUTCOUNT = @TURNOUTCOUNT " & vbCrLf
        sql &= " ,TRUANCYCOUNT = @TRUANCYCOUNT " & vbCrLf
        'LEAVECOUNT. 離訓
        sql &= " ,LEAVECOUNT = @LEAVECOUNT " & vbCrLf
        sql &= " ,REJECTCOUNT = @REJECTCOUNT " & vbCrLf
        'sql &= "   ,ADVJOBCOUNT = @ADVJOBCOUNT " & vbCrLf
        sql &= " ,ITEM1_1 = @ITEM1_1 " & vbCrLf
        sql &= " ,ITEM1_2 = @ITEM1_2 " & vbCrLf
        sql &= " ,ITEM1_COUR = @ITEM1_COUR " & vbCrLf
        sql &= " ,ITEM1_3 = @ITEM1_3 " & vbCrLf
        sql &= " ,ITEM1_TEACHER = @ITEM1_TEACHER " & vbCrLf
        sql &= " ,ITEM1_ASSISTANT = @ITEM1_ASSISTANT " & vbCrLf
        sql &= " ,ITEM2_1 = @ITEM2_1 " & vbCrLf
        sql &= " ,ITEM2_2 = @ITEM2_2 " & vbCrLf
        'sql &= "   ,ITEM2_3 = @ITEM2_3 " & vbCrLf
        sql &= " ,ITEM3_1 = @ITEM3_1 " & vbCrLf
        sql &= " ,ITEM3_2 = @ITEM3_2 " & vbCrLf
        sql &= " ,ITEM7NOTE = @ITEM7NOTE " & vbCrLf
        sql &= " ,ITEM1PROS = @ITEM1PROS " & vbCrLf
        sql &= " ,ITEM2PROS = @ITEM2PROS " & vbCrLf
        sql &= " ,ITEM3PROS = @ITEM3PROS " & vbCrLf
        sql &= " ,ITEM1NOTE = @ITEM1NOTE " & vbCrLf
        sql &= " ,ITEM2NOTE = @ITEM2NOTE " & vbCrLf
        sql &= " ,ITEM3NOTE = @ITEM3NOTE " & vbCrLf
        sql &= " ,ITEM31NOTE =@ITEM31NOTE" & vbCrLf
        sql &= " ,ITEM32 =@ITEM32" & vbCrLf
        sql &= " ,ITEM32NOTE =@ITEM32NOTE" & vbCrLf

        'sql &= " ,CURSENAME = @CURSENAME " & vbCrLf
        sql &= " ,VISITORNAME = @VISITORNAME " & vbCrLf
        sql &= " ,RID = @RID " & vbCrLf
        sql &= " ,MODIFYACCT = @MODIFYACCT " & vbCrLf
        sql &= " ,MODIFYDATE=getdate() " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "  AND OCID = @OCID " & vbCrLf
        sql &= "  AND SEQNO = @SEQNO " & vbCrLf
        Dim uCMD As New SqlCommand(sql, objconn)

        Dim iD3C As Integer = If(D3C1.Checked, 1, If(D3C2.Checked, 2, If(D3C3.Checked, 3, 0)))
        'Dim dt As New DataTable
        'Dim oCmd As New SqlCommand(sql, objconn)
        With uCMD
            .Parameters.Clear()
            .Parameters.Add("APPLYDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(APPLYDATE.Text)
            .Parameters.Add("APPLYDATEHH1", SqlDbType.VarChar).Value = APPLYDATEHH1.Text
            .Parameters.Add("APPLYDATEMI1", SqlDbType.VarChar).Value = APPLYDATEMI1.Text
            .Parameters.Add("APPLYDATEHH2", SqlDbType.VarChar).Value = APPLYDATEHH2.Text
            .Parameters.Add("APPLYDATEMI2", SqlDbType.VarChar).Value = APPLYDATEMI2.Text
            .Parameters.Add("VISTIMES", SqlDbType.Int).Value = Val(VISTIMES.Text)
            .Parameters.Add("DATA1", SqlDbType.VarChar).Value = If(DATA1.SelectedValue = "", Convert.DBNull, DATA1.SelectedValue)
            .Parameters.Add("DATA2", SqlDbType.VarChar).Value = If(DATA2.SelectedValue = "", Convert.DBNull, DATA2.SelectedValue)
            .Parameters.Add("DATA3", SqlDbType.VarChar).Value = If(DATA3.SelectedValue = "", Convert.DBNull, DATA3.SelectedValue)
            .Parameters.Add("DATACOPY1", SqlDbType.VarChar).Value = DATACOPY1.Text
            .Parameters.Add("DATACOPY2", SqlDbType.VarChar).Value = DATACOPY2.Text
            .Parameters.Add("DATACOPY3", SqlDbType.VarChar).Value = DATACOPY3.Text
            .Parameters.Add("D3C", SqlDbType.Int).Value = iD3C
            .Parameters.Add("ITEM7NOTE2", SqlDbType.VarChar).Value = If(ITEM7NOTE2.Text = "", Convert.DBNull, ITEM7NOTE2.Text)
            .Parameters.Add("D1CMM", SqlDbType.VarChar).Value = If(D1CMM.Text = "", Convert.DBNull, D1CMM.Text)
            .Parameters.Add("D1CDD", SqlDbType.VarChar).Value = If(D1CDD.Text = "", Convert.DBNull, D1CDD.Text)
            .Parameters.Add("D2CMM", SqlDbType.VarChar).Value = If(D2CMM.Text = "", Convert.DBNull, D2CMM.Text)
            .Parameters.Add("D2CDD", SqlDbType.VarChar).Value = If(D2CDD.Text = "", Convert.DBNull, D2CDD.Text)
            .Parameters.Add("D3CMM", SqlDbType.VarChar).Value = If(D3CMM.Text = "", Convert.DBNull, D3CMM.Text)
            .Parameters.Add("D3CDD", SqlDbType.VarChar).Value = If(D3CDD.Text = "", Convert.DBNull, D3CDD.Text)
            .Parameters.Add("D3NOTE", SqlDbType.VarChar).Value = D3NOTE.Text
            .Parameters.Add("APPROVEDCOUNT", SqlDbType.Int).Value = Val(APPROVEDCOUNT.Text)
            .Parameters.Add("AUTHCOUNT", SqlDbType.Int).Value = Val(AUTHCOUNT.Text)
            .Parameters.Add("TURTHCOUNT", SqlDbType.Int).Value = Val(TURTHCOUNT.Text)
            .Parameters.Add("TURNOUTCOUNT", SqlDbType.Int).Value = Val(TURNOUTCOUNT.Text)
            .Parameters.Add("TRUANCYCOUNT", SqlDbType.Int).Value = Val(TRUANCYCOUNT.Text)
            'LEAVECOUNT. 離訓
            .Parameters.Add("LEAVECOUNT", SqlDbType.Int).Value = Val(LEAVECOUNT.Text)
            .Parameters.Add("REJECTCOUNT", SqlDbType.Int).Value = Val(REJECTCOUNT.Text)
            '.Parameters.Add("ADVJOBCOUNT", SqlDbType.Int).Value = Val(ADVJOBCOUNT.Text)
            .Parameters.Add("ITEM1_1", SqlDbType.VarChar).Value = If(ITEM1_1.SelectedValue = "", Convert.DBNull, ITEM1_1.SelectedValue)
            .Parameters.Add("ITEM1_2", SqlDbType.VarChar).Value = If(ITEM1_2.SelectedValue = "", Convert.DBNull, ITEM1_2.SelectedValue)
            .Parameters.Add("ITEM1_COUR", SqlDbType.VarChar).Value = ITEM1_COUR.Text
            .Parameters.Add("ITEM1_3", SqlDbType.VarChar).Value = If(ITEM1_3.SelectedValue = "", Convert.DBNull, ITEM1_3.SelectedValue)
            .Parameters.Add("ITEM1_TEACHER", SqlDbType.VarChar).Value = ITEM1_TEACHER.Text
            .Parameters.Add("ITEM1_ASSISTANT", SqlDbType.VarChar).Value = ITEM1_ASSISTANT.Text
            .Parameters.Add("ITEM2_1", SqlDbType.VarChar).Value = If(ITEM2_1.SelectedValue = "", Convert.DBNull, ITEM2_1.SelectedValue)
            .Parameters.Add("ITEM2_2", SqlDbType.VarChar).Value = If(ITEM2_2.SelectedValue = "", Convert.DBNull, ITEM2_2.SelectedValue)
            '.Parameters.Add("ITEM2_3", SqlDbType.VarChar).Value = If(ITEM2_3.SelectedValue = "", Convert.DBNull, ITEM2_3.SelectedValue)
            .Parameters.Add("ITEM3_1", SqlDbType.VarChar).Value = If(ITEM3_1.SelectedValue = "", Convert.DBNull, ITEM3_1.SelectedValue)
            .Parameters.Add("ITEM3_2", SqlDbType.VarChar).Value = If(ITEM3_2.SelectedValue = "", Convert.DBNull, ITEM3_2.SelectedValue)
            .Parameters.Add("ITEM7NOTE", SqlDbType.VarChar).Value = ITEM7NOTE.Text
            .Parameters.Add("ITEM1PROS", SqlDbType.VarChar).Value = ITEM1PROS.Text
            .Parameters.Add("ITEM2PROS", SqlDbType.VarChar).Value = ITEM2PROS.Text
            .Parameters.Add("ITEM3PROS", SqlDbType.VarChar).Value = ITEM3PROS.Text
            .Parameters.Add("ITEM1NOTE", SqlDbType.VarChar).Value = ITEM1NOTE.Text
            .Parameters.Add("ITEM2NOTE", SqlDbType.VarChar).Value = ITEM2NOTE.Text
            .Parameters.Add("ITEM3NOTE", SqlDbType.VarChar).Value = ITEM3NOTE.Text
            .Parameters.Add("ITEM31NOTE", SqlDbType.VarChar).Value = ITEM31NOTE.Text
            .Parameters.Add("ITEM32", SqlDbType.VarChar).Value = If(ITEM32.SelectedValue = "", Convert.DBNull, ITEM32.SelectedValue)
            .Parameters.Add("ITEM32NOTE", SqlDbType.VarChar).Value = ITEM32NOTE.Text

            '.Parameters.Add("CURSENAME", SqlDbType.VarChar).Value = CURSENAME.Text '培訓單位人員姓名
            .Parameters.Add("VISITORNAME", SqlDbType.VarChar).Value = VISITORNAME.Text '訪視人員姓名
            .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
            'dt.Load(.ExecuteReader())
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCIDValue1.Value)
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = iSEQNO
            '.ExecuteNonQuery()  'edit，by:20181011
            DbAccess.ExecuteNonQuery(uCMD.CommandText, objconn, uCMD.Parameters)  'edit，by:20181011
            'rst = .ExecuteScalar()
        End With
    End Sub
End Class