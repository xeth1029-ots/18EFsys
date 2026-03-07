Partial Class TC_04_002_add2
    Inherits AuthBasePage

    '(產投)
    Dim i_gSeqno As Integer = 0 '共用序號使用

    Dim ff3 As String = ""
    Dim GPlanID As String = ""
    Dim GComIDNO As String = ""
    Dim GSeqNO As String = ""
    'Dim vsSEnterDate As String = ""
    'Dim vsFEnterDate As String = ""

    Const Cst_通過 As String = "通過" 'Y
    Const Cst_不通過 As String = "不通過" 'N
    Const Cst_退件修正 As String = "退件修正" 'R
    Const cst_sSeqNoCh_Y As String = "Y"
    Const cst_sSeqNoCh_N As String = "N"
    Const cst_sSeqNoCh_R As String = "R"

    Const cst_Msg_1 As String = "----請填寫不通過的原因----"
    Const cst_Msg_2 As String = "<FONT COLOR=RED>委訓單位申請班級資料變更</FONT>"
    Const cst_Err_1 As String = "請填寫 不通過原因內容" & vbCrLf
    Const cst_Err_2 As String = "請填寫 退件修正原因內容" & vbCrLf
    Const cst_errmsg4 As String = "使用者登入計畫有誤，不提供儲存!!"

    '#Region "FUNCTION "
    Function GetClassID(ByVal obj As DropDownList, ByVal TPlanID As String, ByVal TMID As String, Optional ByVal DistID As String = "") As DropDownList
        Dim sql As String = ""
        sql &= " SELECT ClassName, ClassID FROM ID_CLASS "
        sql &= " WHERE TPlanID='" & TPlanID & "' AND TMID ='" & TMID & "'"
        If DistID <> "" Then sql &= " AND DistID='" & DistID & "'"
        sql &= " ORDER BY CLSID "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        With obj
            .DataSource = dt
            .DataTextField = "ClassName"
            .DataValueField = "ClassID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        Return obj
    End Function

    Function Get_VerSeqNo(ByVal obj As DropDownList) As DropDownList
        With obj
            .Items.Clear()
            .Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            .Items.Add(New ListItem(Cst_通過, cst_sSeqNoCh_Y))
            .Items.Add(New ListItem(Cst_不通過, cst_sSeqNoCh_N))
            .Items.Add(New ListItem(Cst_退件修正, cst_sSeqNoCh_R))
        End With
        Return obj
    End Function

    Sub CreateTrainDesc()
        Dim parms As New Hashtable From {{"PlanID", GPlanID}, {"ComIDNO", GComIDNO}, {"SeqNO", GSeqNO}}
        Dim sql As String = ""
        sql &= " SELECT * FROM PLAN_TRAINDESC" & vbCrLf
        sql &= " WHERE PlanID=@PlanID and ComIDNO=@ComIDNO and SeqNO=@SeqNO" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        Datagrid3Table.Style.Item("display") = "none"
        If dt.Rows.Count = 0 Then Return

        Datagrid3Table.Style.Item("display") = "inline"
        With Datagrid3
            .DataSource = dt
            .DataKeyField = "PTDID"
            .DataBind()
        End With
    End Sub

    '建立上課時間
    Sub CreateClassTime()
        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow
        sql = "SELECT * FROM Plan_OnClass WHERE PlanID='" & GPlanID & "' and ComIDNO='" & GComIDNO & "' and SeqNO='" & GSeqNO & "'"
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.Columns("POCID").AutoIncrement = True
        dt.Columns("POCID").AutoIncrementSeed = -1
        dt.Columns("POCID").AutoIncrementStep = -1
        If dt.Rows.Count = 0 Then
            DataGrid1Table.Visible = False
        Else
            DataGrid1Table.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    Sub CreateItem()
        PlanYear = TIMS.GetSyear(PlanYear)
        '' TPeriod = TIMS.Get_HourRan(TPeriod)
        Call TIMS.Get_ClassCatelog(ClassCate, objconn)
        'CapDegree = TIMS.Get_Degree(CapDegree, 1, objconn)
        CapDegree = TIMS.Get_Degree(CapDegree, 2, objconn)

        ClassID = GetClassID(ClassID, Request("TPlanID"), Request("TMID"), sm.UserInfo.DistID)
        VerSeqNo_ch1 = Get_VerSeqNo(VerSeqNo_ch1)
        Common.SetListItem(Me.VerSeqNo_ch1, "Y")
        VerReason_ch1.Text = cst_Msg_1
    End Sub

    Sub Get_Plan_TrainPlace(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String)
        'Dim dr As DataRow
        'Dim dt As DataTable
        'Dim OtherMsg As String = ""

        Dim v_OtherMsg1 As String = ""
        Dim oParms As New Hashtable From {{"PTCOMIDNO", ComIDNO}, {"PLANID", PlanID}, {"COMIDNO", ComIDNO}, {"SEQNO", SeqNo}}
        Dim objstr As String = ""
        objstr &= " SELECT b.connum ,b.hwdesc ,b.OtherDesc" & vbCrLf
        objstr &= " ,b.PTID ,b.PLACEID ,b.ClassIFICation" & vbCrLf
        objstr &= " FROM PLAN_PLANINFO a" & vbCrLf
        objstr &= " JOIN PLAN_TRAINPLACE b ON a.COMIDNO=b.COMIDNO AND a.SCIPLACEID=b.PLACEID AND b.COMIDNO=@PTCOMIDNO" & vbCrLf
        objstr &= " WHERE a.PLANID=@PLANID AND a.COMIDNO=@COMIDNO AND a.SEQNO=@SEQNO" & vbCrLf
        objstr &= " AND b.ClassIFICation IN (1,3)" & vbCrLf '學科共用。
        Dim dt1 As DataTable = DbAccess.GetDataTable(objstr, objconn, oParms)

        Dim v_OtherMsg2 As String = ""
        Dim oParms2 As New Hashtable From {{"PTCOMIDNO", ComIDNO}, {"PLANID", PlanID}, {"COMIDNO", ComIDNO}, {"SEQNO", SeqNo}}
        Dim objstr2 As String = ""
        objstr2 &= " SELECT b.connum ,b.hwdesc ,b.OtherDesc" & vbCrLf
        objstr2 &= " ,b.PTID,b.PLACEID,b.ClassIFICation" & vbCrLf
        objstr2 &= " FROM PLAN_PLANINFO a" & vbCrLf
        objstr2 &= " JOIN PLAN_TRAINPLACE b ON a.COMIDNO=b.COMIDNO AND a.TECHPLACEID=b.PLACEID AND b.COMIDNO=@PTCOMIDNO" & vbCrLf
        objstr2 &= " WHERE a.PLANID=@PLANID AND a.COMIDNO=@COMIDNO AND a.SEQNO=@SEQNO" & vbCrLf
        objstr2 &= " AND b.ClassIFICation IN (2,3)" & vbCrLf '術科共用。
        Dim dt2 As DataTable = DbAccess.GetDataTable(objstr2, objconn, oParms2)

        If dt1.Rows.Count > 0 Then
            Dim dr1 As DataRow = dt1.Rows(0)
            Tnum2.Text = dr1("connum").ToString
            HwDesc2.Text = dr1("hwdesc").ToString
            v_OtherMsg1 = Convert.ToString(dr1("OtherDesc"))
        End If

        If dt2.Rows.Count > 0 Then
            Dim dr2 As DataRow = dt2.Rows(0)
            Tnum3.Text = dr2("connum").ToString
            HwDesc3.Text = dr2("hwdesc").ToString
            v_OtherMsg2 = Convert.ToString(dr2("OtherDesc"))
        End If

        Dim v_OtherMsgA As String = ""
        If v_OtherMsg1 <> "" AndAlso v_OtherMsg2 <> "" Then
            v_OtherMsgA = String.Concat(v_OtherMsg1, If(v_OtherMsg1 <> v_OtherMsg2, String.Concat(vbCrLf, v_OtherMsg2), ""))
        ElseIf v_OtherMsg1 <> "" AndAlso v_OtherMsg1.Length > 1 Then
            v_OtherMsgA = v_OtherMsg1
        ElseIf v_OtherMsg2 <> "" AndAlso v_OtherMsg2.Length > 1 Then
            v_OtherMsgA = v_OtherMsg2
        End If

        OtherDesc23.Text = v_OtherMsgA
    End Sub


    Function GET_PLANTEACHER() As DataTable
        Dim dt As DataTable

        Dim sql As String = ""
        sql &= " select a.TechID, a.TeachCName, a.DegreeID, c.Name DegreeName" & vbCrLf
        sql &= " ,replace(ISNULL(a.Specialty1,' '),',',' ')" & vbCrLf
        sql &= " +replace(ISNULL(a.Specialty2,' '),',',' ')" & vbCrLf
        sql &= " +replace(ISNULL(a.Specialty3,' '),',',' ')" & vbCrLf
        sql &= " +replace(ISNULL(a.Specialty4,' '),',',' ')" & vbCrLf
        sql &= " +replace(ISNULL(a.Specialty5,' '),',',' ') as major" & vbCrLf
        sql &= " ,CASE ISNULL(CONVERT(varchar, b.TechID),' ') when ' ' then 'N' ELSE 'Y'  END ptchk" & vbCrLf
        sql &= " FROM TEACH_TEACHERINFO a" & vbCrLf
        sql &= " LEFT JOIN KEY_DEGREE c on a.DegreeID=c.DegreeID" & vbCrLf

        sql &= " left join ( select TechID from Plan_Teacher" & vbCrLf
        sql &= " WHERE TechTYPE='A'" 'TechTYPE: A:師資/B:助教
        sql &= " AND PlanID='" & GPlanID & "'" & vbCrLf
        sql &= " AND ComIDNO='" & GComIDNO & "'" & vbCrLf
        sql &= " AND SeqNo='" & GSeqNO & "') b on a.TechID =b.TechID" & vbCrLf

        sql &= " where WorkStatus = '1'" & vbCrLf
        sql &= " and RID IN (SELECT RID" & vbCrLf
        sql &= " from Plan_PlanInfo" & vbCrLf
        sql &= " WHERE PlanID='" & GPlanID & "'" & vbCrLf
        sql &= " AND ComIDNO='" & GComIDNO & "'" & vbCrLf
        sql &= " AND SeqNo='" & GSeqNO & "')" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.Columns("TechID").AutoIncrement = True
        dt.Columns("TechID").AutoIncrementSeed = -1
        dt.Columns("TechID").AutoIncrementStep = -1
        Return dt
    End Function
    '#End Region

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
        '檢查Session是否存在 End

        TRClassid.Visible = False
        'Me.RIDValue.Value = sm.UserInfo.RID
        'GPlanID = Request("PlanID")
        'GComIDNO = Request("ComIDNO")
        'GSeqNO = Request("SeqNO")
        GPlanID = TIMS.ClearSQM(Request("PlanID"))
        GComIDNO = TIMS.ClearSQM(Request("ComIDNO"))
        GSeqNO = TIMS.ClearSQM(Request("SeqNO"))
        If GPlanID = "" Then Exit Sub
        If GComIDNO = "" Then Exit Sub
        If GSeqNO = "" Then Exit Sub

        If Not Page.IsPostBack Then
            '建立物件----Start
            Call CreateItem()
            Call Get_Plan_TrainPlace(GPlanID, GComIDNO, GSeqNO)
            Call CreateClassTime()
            Call CreateTrainDesc()
            'Call ShowReqValue()
            Call SHOW_PLAN_PLANINFO()
            Call SHOW_PLAN_VERREPORT()
            Call SHOW_PLAN_TEACHER12(Hid_RIDValue.Value)
        End If

        If Not Session("_search") Is Nothing Then ViewState("_search") = Session("_search") 'Session("_search") = Nothing

    End Sub

    '儲存
    Private Sub bt_addrow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_addrow.Click
        Const cst_verType_F As String = "F"
        Const cst_verType_S As String = "S"

        Dim sVerType As String = TIMS.GetVerType(sm.UserInfo.LID)
        Select Case sVerType 'TIMS.GetVerType(sm.UserInfo.LID)
            Case cst_verType_F, cst_verType_S
            Case Else
                Common.MessageBox(Me, "儲存失敗，該登入者無審核權限!!")
                Exit Sub
        End Select

        Dim ERRMsg As String = ""
        Dim VerReasonMsg As String = ""
        Dim strVerSeqNoCh As String = ""
        If VerSeqNo_ch1.SelectedValue <> "" Then strVerSeqNoCh = VerSeqNo_ch1.SelectedValue

        Select Case strVerSeqNoCh
            Case cst_sSeqNoCh_N
                VerReason_ch1.Text = TIMS.ClearSQM(VerReason_ch1.Text)
                If VerReason_ch1.Text = "" Then ERRMsg += cst_Err_1

            Case cst_sSeqNoCh_R
                VerReason_ch1.Text = TIMS.ClearSQM(VerReason_ch1.Text)
                If VerReason_ch1.Text = "" Then ERRMsg += cst_Err_2

        End Select
        If ERRMsg <> "" Then
            Common.MessageBox(Me, ERRMsg)
            Exit Sub
        End If

        Dim drPP As DataRow = TIMS.GetPPInfo(GPlanID, GComIDNO, GSeqNO, objconn)
        '檢核報名日期 (若OK 轉出OUT SEnterDate/FEnterDate)
        Dim vSTDate As String = TIMS.Cdate3(drPP("STDate"))
        Dim vSEnterDate As String = "" 'TIMS.GetMyValue2(htCC, "SEnterDate")
        Dim vFEnterDate As String = "" 'TIMS.GetMyValue2(htCC, "FEnterDate") 'Dim flag_chkSEnDate As Boolean = False 'false:異常
        Call TIMS.ChangeSEnterDate(vSTDate, vSEnterDate, vFEnterDate)

        Dim V_Errmsg As String = ""
        Dim flag_chkSEnDate As Boolean = True 'false:異常
        If vSEnterDate = "" Then flag_chkSEnDate = False
        If vFEnterDate = "" Then flag_chkSEnDate = False
        Hid_SENTERDATE.Value = vSEnterDate
        Hid_FENTERDATE.Value = vFEnterDate
        If Not flag_chkSEnDate Then
            '報名時間有誤不執行轉入
            V_Errmsg = "開訓時間計算報名時間有誤 不可執行審核作業!"
            Common.MessageBox(Me, V_Errmsg)
            Exit Sub
        End If

        '系統管理者(不管開訓日)
        Dim flgROLEIDx0xLIDx0 As Boolean = TIMS.IsSuperUser(sm, 1)
        'PlanMode:S:審核中/Y:已通過/R:退件修正(含不通過的)
        If Not flgROLEIDx0xLIDx0 AndAlso strVerSeqNoCh = cst_sSeqNoCh_Y Then
            '小於、等於 開訓前三天 -不可報名
            Dim flag_chkSEnDate3 As Boolean = TIMS.ChkEnterDayS3(vSTDate)
            If Not flag_chkSEnDate3 Then
                'V_Errmsg = "開訓時間 小於、等於 開訓前三天 不可報名(不執行轉入)!"
                V_Errmsg = "班級審核日距離開訓日為3日(含)內，不可執行審核作業!"
                Common.MessageBox(Me, V_Errmsg)
                Exit Sub
            End If
        End If

        VerReasonMsg = ""
        Select Case sVerType
            Case cst_verType_F
                Select Case strVerSeqNoCh
                    Case cst_sSeqNoCh_N, cst_sSeqNoCh_R
                        TIMS.Plan_VerRecord_Update(GPlanID, GComIDNO, GSeqNO, sm.UserInfo.UserID, cst_verType_F, 1, VerReason_ch1.Text, objconn)
                        VerReasonMsg &= VerReason_ch1.Text & vbCrLf
                    Case Else 'Y
                        TIMS.Plan_VerRecord_Update(GPlanID, GComIDNO, GSeqNO, sm.UserInfo.UserID, cst_verType_F, 1, "", objconn)
                End Select
            Case cst_verType_S
                Select Case strVerSeqNoCh
                    Case cst_sSeqNoCh_N, cst_sSeqNoCh_R
                        TIMS.Plan_VerRecord_Update(GPlanID, GComIDNO, GSeqNO, sm.UserInfo.UserID, cst_verType_S, 1, VerReason_ch1.Text, objconn)
                        VerReasonMsg &= VerReason_ch1.Text & vbCrLf
                    Case Else 'Y
                        TIMS.Plan_VerRecord_Update(GPlanID, GComIDNO, GSeqNO, sm.UserInfo.UserID, cst_verType_S, 1, "", objconn)
                End Select
        End Select

        '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
        Select Case sVerType'TIMS.GetVerType(sm.UserInfo.LID.ToString)
            Case cst_verType_F, cst_verType_S
                'VER@VerSeqNo_ch1.SelectedValue Y@N@R
                If Trim(VerReasonMsg) <> "" Then
                    Call TIMS.Plan_VerReprot_Update(Me, GPlanID, GComIDNO, GSeqNO, strVerSeqNoCh, objconn)
                Else
                    Call TIMS.Plan_VerReprot_Update(Me, GPlanID, GComIDNO, GSeqNO, cst_sSeqNoCh_Y, objconn)
                End If
        End Select

        '執行班級轉入-----------------start
        Select Case strVerSeqNoCh
            Case cst_sSeqNoCh_N
                V_Errmsg = "審核不通過 不執行轉入!"
                Common.MessageBox(Me, V_Errmsg)
                Exit Sub
            Case cst_sSeqNoCh_R
                V_Errmsg = "退件修正 不執行轉入!"
                Common.MessageBox(Me, V_Errmsg)
                Exit Sub
        End Select
        'Case Else 'Y  If VerReasonMsg = "" Then flag_SaveCC = True 'true:可執行轉入/false:不可執行

        '檢核計畫與班級轉入
        V_Errmsg = ""
        Dim flag_SaveCC As Boolean = False 'true:可執行轉入/false:不可執行
        flag_SaveCC = xChk_CC1(V_Errmsg)
        If V_Errmsg <> "" Then
            '審核有誤不執行轉入
            Common.MessageBox(Me, V_Errmsg) '"儲存失敗，審核有誤 不執行轉入!")
            Exit Sub
        End If
        If Not flag_SaveCC Then
            Common.MessageBox(Me, "審核狀況有誤 不執行轉入!")
            Exit Sub
        End If

        '儲存(CLASS_CLASSINFO) --班級轉入
        Call SaveData1()

        If Session("_search") Is Nothing Then Session("_search") = ViewState("_search")
        Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "TC_04_002.aspx?ID=" & Request("ID"))
    End Sub

    '回上一頁 
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Session("_search") Is Nothing Then Session("_search") = ViewState("_search")
        Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "TC_04_002.aspx?ID=" & Request("ID"))
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Weeks As Label = e.Item.FindControl("Weeks1")
                Dim Times As Label = e.Item.FindControl("Times1")
                Dim drv As DataRowView = e.Item.DataItem
                'Dim btn1 As Button = e.Item.FindControl("Button2")
                'Dim btn2 As Button = e.Item.FindControl("Button3")

                'btn1.Enabled = Button29.Enabled
                'btn2.Enabled = Button29.Enabled
                Weeks.Text = drv("Weeks").ToString
                Times.Text = drv("Times").ToString
                'btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                'btn2.CommandArgument = drv("POCID")
            Case ListItemType.EditItem
                Dim Weeks As DropDownList = e.Item.FindControl("Weeks2")
                Dim Times As TextBox = e.Item.FindControl("Times2")
                'Dim btn1 As Button = e.Item.FindControl("Button4")
                'Dim btn2 As Button = e.Item.FindControl("Button5")
                Dim drv As DataRowView = e.Item.DataItem

                With Weeks
                    .Items.Add(New ListItem("==請選擇==", ""))
                    .Items.Add(New ListItem("星期一", "星期一"))
                    .Items.Add(New ListItem("星期二", "星期二"))
                    .Items.Add(New ListItem("星期三", "星期三"))
                    .Items.Add(New ListItem("星期四", "星期四"))
                    .Items.Add(New ListItem("星期五", "星期五"))
                    .Items.Add(New ListItem("星期六", "星期六"))
                    .Items.Add(New ListItem("星期日", "星期日"))
                End With
                Common.SetListItem(Weeks, drv("Weeks").ToString)
                Times.Text = drv("Times").ToString
                'btn1.CommandArgument = drv("POCID")
        End Select

    End Sub

    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim STrainDateLabel As Label = e.Item.FindControl("STrainDateLabel")
                Dim PNameLabel As Label = e.Item.FindControl("PNameLabel")
                Dim PHourLabel As Label = e.Item.FindControl("PHourLabel")
                Dim PContText As TextBox = e.Item.FindControl("PContText")
                Dim drpClassification1 As DropDownList = e.Item.FindControl("drpClassification1")
                Dim drpPTID As DropDownList = e.Item.FindControl("drpPTID")
                Dim cb_FARLEARNi As CheckBox = e.Item.FindControl("cb_FARLEARNi")
                Dim Tech1Value As HtmlInputHidden = e.Item.FindControl("Tech1Value")
                Dim Tech1Text As TextBox = e.Item.FindControl("Tech1Text")

                If drv("STrainDate").ToString <> "" Then
                    STrainDateLabel.Text = Common.FormatDate(drv("STrainDate").ToString)
                End If
                PNameLabel.Text = drv("PName").ToString
                PHourLabel.Text = drv("PHour").ToString
                PContText.Text = drv("PCont").ToString
                If drv("Classification1").ToString <> "" Then
                    Common.SetListItem(drpClassification1, drv("Classification1").ToString)

                    Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "ComIDNO", Hid_ComIDNO.Value)
                    If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                    Select Case drpClassification1.SelectedValue
                        Case "1" '學科
                            drpPTID = TIMS.Get_SciPTID(drpPTID, Hid_ComIDNO.Value, 1, objconn)
                        Case "2" '術科
                            drpPTID = TIMS.Get_TechPTID(drpPTID, Hid_ComIDNO.Value, 1, objconn)
                    End Select

                    If drv("PTID").ToString <> "" Then
                        Common.SetListItem(drpPTID, drv("PTID").ToString)
                    End If
                End If
                '遠距教學
                cb_FARLEARNi.Checked = If(Convert.ToString(drv("FARLEARN")).Equals("Y"), True, False)
                If drv("TechID").ToString <> "" Then
                    Tech1Value.Value = drv("TechID").ToString
                    Tech1Text.Text = TIMS.Get_TeachCName(Tech1Value.Value, objconn) '
                    'Tech1Text.Text = TIMS.Get_TeacherName(drv("TechID").ToString)
                End If
        End Select
    End Sub

    '確認各物件的屬性
    Sub SET_READ_ONLY()
        rblFuncLevel.Enabled = False
        cblTMethod.Enabled = False
        TMethodOth.Enabled = False
        tPOWERNEED1.Enabled = False
        tPOWERNEED2.Enabled = False
        tPOWERNEED3.Enabled = False
        cbPOWERNEED4.Enabled = False
        tPOWERNEED4.Enabled = False
        CapAll.Enabled = False
        RecDesc.Enabled = False
        LearnDesc.Enabled = False
        ActDesc.Enabled = False
        ResultDesc.Enabled = False
        OtherDesc.Enabled = False

        ClassName.ReadOnly = True
        ClassCate.Enabled = False
        Tnum.ReadOnly = True
        THours.ReadOnly = True
        'Times.ReadOnly = False
        start_date.ReadOnly = True
        end_date.ReadOnly = True
        IMG1.Visible = False
        IMG2.Visible = False
        'Me.ProcID.Enabled = False
        'Me.OLessonTeah1.ReadOnly = True
        CapDegree.Enabled = False
        'TPeriod.Enabled = False
        'Times.ReadOnly = True
        ClassID.Enabled = False
        'TrainDemain.ReadOnly = True
        'Me.TrainDemain.ReadOnly = True
        'Me.TrainTarget.ReadOnly = True

        'Me.OLessonTeah1.ReadOnly = True
        'Me.TeacherDesc.ReadOnly = True
        'Me.TeacherDesc2.ReadOnly = True
        'Me.Domain.ReadOnly = True

        CapAll.ReadOnly = True
        CostDesc.ReadOnly = True

        'Me.TrainMode.ReadOnly = True
        'Me.Content.ReadOnly = True

        RecDesc.ReadOnly = True
        LearnDesc.ReadOnly = True
        ActDesc.ReadOnly = True
        ResultDesc.ReadOnly = True
        OtherDesc.ReadOnly = True

        '是否為iCAP課程 / 是, 請填寫/否/ 課程相關說明
        tb_ISiCAPCOUR.Disabled = True
        RB_ISiCAPCOUR_Y.Enabled = False
        RB_ISiCAPCOUR_N.Enabled = False
        iCAPCOURDESC.ReadOnly = True '課程相關說明
        Recruit.ReadOnly = True '招訓方式
        Selmethod.ReadOnly = True '遴選方式
        Inspire.ReadOnly = True '學員激勵辦法

        CapDegree.Enabled = False
        If Request("CmdStatus") IsNot Nothing Then
            If TIMS.ClearSQM(Request("CmdStatus")) = "View" Then bt_addrow.Enabled = False
        End If
    End Sub

    'requeset-PLAN_PLANINFO
    Sub SHOW_PLAN_PLANINFO()
        '確認各物件的屬性
        Call SET_READ_ONLY()

        Dim sql As String = ""
        sql &= " SELECT pp.*" & vbCrLf
        sql &= " ,o1.ORGNAME" & vbCrLf
        sql &= " FROM dbo.PLAN_PLANINFO pp" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO O1 ON O1.COMIDNO=pp.COMIDNO" & vbCrLf
        sql &= " WHERE pp.PlanID='" & GPlanID & "'" & vbCrLf
        sql &= " AND pp.ComIDNO='" & GComIDNO & "'" & vbCrLf
        sql &= " AND pp.SeqNo='" & GSeqNO & "'" & vbCrLf
        Dim dt As DataTable = Nothing 'Plan_VerReport
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt Is Nothing Then Exit Sub
        If dt.Rows.Count <> 1 Then Exit Sub

        Dim dr1 As DataRow = dt.Rows(0)
        labORGNAME.Text = Convert.ToString(dr1("ORGNAME"))

        Hid_RIDValue.Value = Convert.ToString(dr1("RID"))

        Common.SetListItem(PlanYear, TIMS.ClearSQM(Request("PlanYear")))
        ClassName.Text = HttpUtility.UrlDecode(Request("ClassName"))
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        Common.SetListItem(ClassCate, TIMS.ClearSQM(Request("ClassCate")))
        Tnum.Text = TIMS.ClearSQM(Request("Tnum"))
        THours.Text = TIMS.ClearSQM(Request("THours"))
        start_date.Text = TIMS.ClearSQM(Request("STDate"))
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.ClearSQM(Request("FDDate"))
        end_date.Text = TIMS.Cdate3(end_date.Text)
        'If Request("ProcID") <> "" Then
        '    Common.SetListItem(ProcID, Request("ProcID"))
        'End If
        Dim vCapDegree As String = TIMS.ClearSQM(Request("CapDegree"))
        If vCapDegree <> "" Then
            Common.SetListItem(Me.CapDegree, vCapDegree)
        End If

        DefGovCost.Text = TIMS.ClearSQM(Request("DefGovCost"))
        If DefGovCost.Text = "" Then DefGovCost.Text = 0
        DefStdCost.Text = TIMS.ClearSQM(Request("DefStdCost"))
        If DefStdCost.Text = "" Then DefStdCost.Text = 0
        TotalCost.Text = CInt(DefGovCost.Text) + CInt(DefStdCost.Text)
        If Tnum.Text <> "" Then
            DefGovCost_Tnum.Text = CInt(DefGovCost.Text) / CInt(Tnum.Text)
            DefStdCost_Tnum.Text = CInt(DefStdCost.Text) / CInt(Tnum.Text)
        End If
    End Sub

    'PLAN_VERREPORT
    Sub SHOW_PLAN_VERREPORT() 'ByVal htSS As Hashtable)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT pv.*" & vbCrLf
        sql &= " ,ISNULL(pp.AppliedResult,' ') AppliedResult" & vbCrLf
        sql &= " FROM dbo.PLAN_VERREPORT pv" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO pp on pv.PlanID = pp.PlanID AND pv.ComIDNO = pp.ComIDNO AND pv.SeqNO = pp.SeqNo" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND pv.PlanID='" & GPlanID & "'" & vbCrLf
        sql &= " AND pv.ComIDNO='" & GComIDNO & "'" & vbCrLf
        sql &= " AND pv.SeqNo='" & GSeqNO & "'" & vbCrLf
        Dim dt As DataTable = Nothing 'Plan_VerReport
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            'Common.RespWrite(Me, "<script>alert('請先執行開班計畫表資料維護作業 ');</script>")
            'TIMS.CloseDbConn(objconn)
            'Server.Transfer("TC_04_002.aspx?ID=" & Request("ID"))
            Dim url1 As String = "TC_04_002.aspx?ID=" & Request("ID")
            Common.MessageBox(Me, "請先執行開班計畫表資料維護作業", url1)
            Exit Sub
        End If
        'If dt Is Nothing Then Exit Sub
        'If dt.Rows.Count = 0 Then Exit Sub

        Dim dr As DataRow = dt.Rows(0)
        Common.SetListItem(rblFuncLevel, Convert.ToString(dr("FuncLevel")))
        If Convert.ToString(dr("TMethod")) <> "" Then TIMS.SetCblValue(cblTMethod, Convert.ToString(dr("TMethod")))
        TMethodOth.Text = Convert.ToString(dr("TMethodOth"))
        'Common.SetListItem(ClassID, dr("ClassID"))
        TIMS.PL_settextbox1(tPOWERNEED1, dr("POWERNEED1"))
        TIMS.PL_settextbox1(tPOWERNEED2, dr("POWERNEED2"))
        TIMS.PL_settextbox1(tPOWERNEED3, dr("POWERNEED3"))
        cbPOWERNEED4.Checked = False
        If Convert.ToString(dr("POWERNEED4CHK")) = TIMS.cst_YES Then cbPOWERNEED4.Checked = True
        If cbPOWERNEED4.Checked AndAlso Convert.ToString(dr("POWERNEED4")) <> "" Then TIMS.PL_settextbox1(tPOWERNEED4, dr("POWERNEED4"))
        'If tPlanCause.Text = "" AndAlso Convert.ToString(dr("PlanCause")) <> "" Then tPlanCause.Text = Convert.ToString(dr("PlanCause"))
        'If tPurScience.Text = "" AndAlso Convert.ToString(dr("PurScience")) <> "" Then tPurScience.Text = Convert.ToString(dr("PurScience"))
        'If tPurTech.Text = "" AndAlso Convert.ToString(dr("PurTech")) <> "" Then tPurTech.Text = Convert.ToString(dr("PurTech"))
        'If tPurMoral.Text = "" AndAlso Convert.ToString(dr("PurMoral")) <> "" Then tPurMoral.Text = Convert.ToString(dr("PurMoral"))
        CapAll.Text = dr("CapAll").ToString
        'If CostDesc.Text = "" Then
        '    If dr("CostDesc").ToString <> "" Then CostDesc.Text = dr("CostDesc").ToString
        'End If
        'Me.TrainMode.Enabled = False
        'Me.TrainMode.Text = "(請勾選教學方法)"
        RecDesc.Text = Convert.ToString(dr("RecDesc")) '.ToString
        LearnDesc.Text = Convert.ToString(dr("LearnDesc")) '.ToString
        ActDesc.Text = Convert.ToString(dr("ActDesc")) '.ToString
        ResultDesc.Text = Convert.ToString(dr("ResultDesc")) '.ToString
        OtherDesc.Text = Convert.ToString(dr("OtherDesc")) '.ToString

        'chk_RecDesc.Checked = False
        'If RecDesc.Text <> "" Then chk_RecDesc.Checked = True
        'chk_LearnDesc.Checked = False
        'If LearnDesc.Text <> "" Then chk_LearnDesc.Checked = True
        'chk_ActDesc.Checked = False
        'If ActDesc.Text <> "" Then chk_ActDesc.Checked = True
        'chk_ResultDesc.Checked = False
        'If ResultDesc.Text <> "" Then chk_ResultDesc.Checked = True
        'chk_OtherDesc.Checked = False
        'If OtherDesc.Text <> "" Then chk_OtherDesc.Checked = True
        'Me.Recruit.Text = dr("Recruit").ToString
        'Me.Inspire.Text = dr("Inspire").ToString
        'TGovExamCY.Checked = False
        'TGovExamCN.Checked = False
        'Select Case Convert.ToString(dr("TGovExam"))
        '    Case "Y"
        '        TGovExamCY.Checked = True
        '    Case "N"
        '        TGovExamCN.Checked = True
        'End Select
        'Me.TGovExamName.Text = dr("TGovExamName").ToString
        'chkMEMO8C1.Checked = False
        'chkMEMO8C2.Checked = False
        'Me.txtMemo8.Text = ""
        'If Convert.ToString(dr("memo8")) <> "" Then chkMEMO8C1.Checked = True
        'If Convert.ToString(dr("memo82")) <> "" Then
        '    chkMEMO8C2.Checked = True
        '    txtMemo8.Text = Convert.ToString(dr("memo82"))
        'End If

        'dbo.fn_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'B', 407731)
        'Dim strTDA As String = GET_TeacherDesc_AB("A")
        'Dim strTDB As String = GET_TeacherDesc_AB("B")
        'TeacherDesc_A.Text = strTDA 'Convert.ToString(drv("TeacherDesc"))
        'TeacherDesc_B.Text = strTDB 'Convert.ToString(drv("TeacherDesc"))



        'Dim dr As DataRow = dt.Rows(0)
        Common.SetListItem(ClassID, dr("ClassID"))
        ' Times.Text = dr("Times")
        'Common.SetListItem(Me.TPeriod, dr("TPeriod"))
        'Me.TrainDemain.Text = dr("TrainDemain").ToString
        'Me.TrainTarget.Text = dr("TrainTarget").ToString

        'Me.TeacherDesc.Text = dr("TeacherDesc").ToString
        'Me.TeacherDesc2.Text = dr("TeacherDesc2").ToString
        'Me.Domain.Text = dr("Domain").ToString

        CapAll.Text = dr("CapAll").ToString
        CostDesc.Text = dr("CostDesc").ToString

        'Me.TrainMode.Text = dr("TrainMode").ToString
        'Me.Content.Text = dr("Content").ToString

        RecDesc.Text = dr("RecDesc").ToString
        LearnDesc.Text = dr("LearnDesc").ToString
        ActDesc.Text = dr("ActDesc").ToString
        ResultDesc.Text = dr("ResultDesc").ToString
        OtherDesc.Text = dr("OtherDesc").ToString

        '是否為iCAP課程 / 是, 請填寫/否/ 課程相關說明
        Dim sISiCAPCOUR As String = Convert.ToString(dr("ISiCAPCOUR"))
        RB_ISiCAPCOUR_Y.Checked = If(sISiCAPCOUR = "Y", True, False)
        RB_ISiCAPCOUR_N.Checked = If(sISiCAPCOUR = "N", True, False)
        iCAPCOURDESC.Text = Convert.ToString(dr("iCAPCOURDESC")) '課程相關說明
        Recruit.Text = Convert.ToString(dr("Recruit")) '招訓方式
        Selmethod.Text = Convert.ToString(dr("Selmethod")) '遴選方式
        Inspire.Text = Convert.ToString(dr("Inspire")) '學員激勵辦法

        'PLAN_PLANINFO.AppliedResult
        Select Case Convert.ToString(dr("AppliedResult"))
            Case "O" '(OLD資料)
                Label1.Text = cst_Msg_2
                Common.SetListItem(Me.VerSeqNo_ch1, "")
                VerReason_ch1.Text = ""
                Exit Sub
            Case "N"
                'Me.VerReason_ch1.Text = Get_Plan_VerRecord_VerReason(GetVerType(sm.UserInfo.LID.ToString), 1)
                'Common.SetListItem(Me.VerSeqNo_ch1, "N")
                If VerReason_ch1.Text = "" Then
                    VerReason_ch1.Text = Cst_不通過
                End If
            Case "Y"
                'Common.SetListItem(Me.VerSeqNo_ch1, "Y")
                VerReason_ch1.Text = ""
                If sm.UserInfo.RID <> "A" Then '非署(局)的單位: 署(局) = sm.UserInfo.RID:"A"
                    VerSeqNo_ch1.Enabled = False
                    VerReason_ch1.Enabled = False
                    bt_addrow.Enabled = False
                    TIMS.Tooltip(VerSeqNo_ch1, "已審核通過")
                    TIMS.Tooltip(VerReason_ch1, "已審核通過")
                    TIMS.Tooltip(bt_addrow, "已審核通過")
                Else
                    TIMS.Tooltip(VerSeqNo_ch1, "已審核通過，該登入機構可修改")
                    TIMS.Tooltip(VerReason_ch1, "已審核通過，該登入機構可修改")
                    TIMS.Tooltip(bt_addrow, "已審核通過，該登入機構可修改")
                End If
            Case "R"
                If VerReason_ch1.Text = "" Then VerReason_ch1.Text = Cst_退件修正

            Case Else
                Common.SetListItem(Me.VerSeqNo_ch1, "")
                VerReason_ch1.Text = ""

        End Select

        Label1.Text = ""
        Common.SetListItem(Me.VerSeqNo_ch1, Convert.ToString(dr("AppliedResult")))
        Dim s_VerType As String = TIMS.GetVerType(sm.UserInfo.LID.ToString)
        Dim str_VerReason_ch1 As String = ""
        Select Case s_VerType
            Case "S" '第2層級，可參考第1層級
                str_VerReason_ch1 = TIMS.Get_Plan_VerRecord_VerReason(GPlanID, GComIDNO, GSeqNO, s_VerType, 1, objconn)
                If str_VerReason_ch1 = "" Then str_VerReason_ch1 = TIMS.Get_Plan_VerRecord_VerReason(GPlanID, GComIDNO, GSeqNO, "F", 1, objconn)
            Case Else
                str_VerReason_ch1 = TIMS.Get_Plan_VerRecord_VerReason(GPlanID, GComIDNO, GSeqNO, s_VerType, 1, objconn)
        End Select
        VerReason_ch1.Text = str_VerReason_ch1 'TIMS.Get_Plan_VerRecord_VerReason(GPlanID, GComIDNO, GSeqNO, s_VerType, 1, objconn)
    End Sub

    '建立可選教師列表
    Sub SHOW_PLAN_TEACHER12(ByVal rqRID As String)
        'If g_flagNG Then
        '    sm.LastErrorMessage = cst_errmsg3
        '    Exit Sub
        'End If
        'Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        'Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        'Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        'If rqPlanID = "" Then Exit Sub
        'If rqComIDNO = "" Then Exit Sub
        'If rqSeqNO = "" Then Exit Sub
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'Dim rqRID As String = Convert.ToString(sm.UserInfo.RID)
        'If RIDValue.Value <> "" Then rqRID = RIDValue.Value

        'ByVal TechTYPE As String
        'Dim dtT As New DataTable
        'If upt_PlanX.Value = "" Then Exit Sub '無有效值離開
        'tmpPCS = upt_PlanX.Value  '有儲存資料過了
        'PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
        'ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
        'SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
        'Dim rqPlanID As String = PlanID_value 'TIMS.GetMyValue2(htSS, "rqPlanID")
        'Dim rqComIDNO As String = ComIDNO_value 'TIMS.GetMyValue2(htSS, "rqComIDNO")
        'Dim rqSeqNO As String = SeqNO_value 'TIMS.GetMyValue2(htSS, "rqSeqNO")
        'Dim sRIDn As String = sUtl_GetRIDn()
        'If sRIDn <> "" Then rqRID = sRIDn
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("PlanID", GPlanID)
        parms.Add("ComIDNO", GComIDNO)
        parms.Add("SeqNO", GSeqNO)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.TechID" & vbCrLf '教師ID
        sql &= " ,a.TeachCName" & vbCrLf '教師姓名 
        sql &= " ,a.DegreeID" & vbCrLf '學歷
        sql &= " ,c.Name DegreeName" & vbCrLf '學歷

        '專業領域
        sql &= " ,REPLACE(ISNULL(a.Specialty1, ' '),',',' ')" & vbCrLf
        sql &= "  +REPLACE(ISNULL(a.Specialty2, ' '),',',' ')" & vbCrLf
        sql &= "  +REPLACE(ISNULL(a.Specialty3, ' '),',',' ')" & vbCrLf
        sql &= "  +REPLACE(ISNULL(a.Specialty4, ' '),',',' ')" & vbCrLf
        sql &= "  +REPLACE(ISNULL(a.Specialty5, ' '),',',' ') major" & vbCrLf '專業領域

        '專業證照-相關證照
        sql &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1 + '、' + a.ProLicense2" & vbCrLf
        sql &= " ELSE a.ProLicense END ProLicense" & vbCrLf

        sql &= " ,dbo.FN_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'A', a.TechID) TeacherDesc " 'TechTYPE: A:師資/B:助教
        sql &= " FROM dbo.TEACH_TEACHERINFO a" & vbCrLf
        sql &= " JOIN (" & vbCrLf
        sql &= " SELECT DISTINCT TechID, planid, comidno, seqno" & vbCrLf
        sql &= " FROM dbo.PLAN_TRAINDESC" & vbCrLf
        sql &= " WHERE TECHID IS NOT NULL" & vbCrLf
        'sql &= " AND PlanID = '" & GPlanID & "' AND ComIDNO = '" & GComIDNO & "' AND SeqNo = '" & GSeqNO & "'" & vbCrLf
        sql &= " AND PlanID=@PlanID and ComIDNO=@ComIDNO and SeqNO=@SeqNO" & vbCrLf
        sql &= " ) b ON a.TechID = b.TechID" & vbCrLf
        sql &= " LEFT JOIN dbo.KEY_DEGREE c ON a.DegreeID = c.DegreeID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.WorkStatus = '1'" & vbCrLf
        sql &= " AND a.RID = '" & rqRID & "'" & vbCrLf
        sql &= " ORDER BY a.TechID" & vbCrLf
        Dim dtT As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        i_gSeqno = 0
        tbDataGrid21.Visible = False
        If dtT.Rows.Count > 0 Then
            tbDataGrid21.Visible = True
            DataGrid21.DataSource = dtT
            DataGrid21.DataBind()
        End If
        'If TechTYPE = "A" Then
        '    'Dim dtT As DataTable = Nothing
        '    dtT = DbAccess.GetDataTable(sql, objconn)
        'End If

        Dim parms2 As New Hashtable
        parms2.Clear()
        parms2.Add("PlanID", GPlanID)
        parms2.Add("ComIDNO", GComIDNO)
        parms2.Add("SeqNO", GSeqNO)

        sql = "" & vbCrLf
        sql &= " SELECT a.TechID" & vbCrLf '教師ID
        sql &= " ,a.TeachCName" & vbCrLf '教師姓名 
        sql &= " ,a.DegreeID" & vbCrLf '學歷
        sql &= " ,c.Name DegreeName" & vbCrLf '學歷

        '專業領域
        sql &= " ,REPLACE(ISNULL(a.Specialty1, ' '),',',' ')" & vbCrLf
        sql &= "  +REPLACE(ISNULL(a.Specialty2, ' '),',',' ')" & vbCrLf
        sql &= "  +REPLACE(ISNULL(a.Specialty3, ' '),',',' ')" & vbCrLf
        sql &= "  +REPLACE(ISNULL(a.Specialty4, ' '),',',' ')" & vbCrLf
        sql &= "  +REPLACE(ISNULL(a.Specialty5, ' '),',',' ') major" & vbCrLf '專業領域

        '專業證照-相關證照
        sql &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1 + '、' + a.ProLicense2" & vbCrLf
        sql &= " ELSE a.ProLicense END ProLicense" & vbCrLf

        sql &= " ,dbo.FN_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'B', a.TechID) TeacherDesc " 'TechTYPE: A:師資/B:助教
        sql &= " FROM dbo.TEACH_TEACHERINFO a" & vbCrLf
        sql &= " JOIN (" & vbCrLf
        sql &= " SELECT DISTINCT TECHID2 TechID, planid, comidno, seqno" & vbCrLf
        sql &= " FROM dbo.PLAN_TRAINDESC" & vbCrLf
        sql &= " WHERE TECHID2 IS NOT NULL" & vbCrLf
        'sql &= " AND PlanID = '" & GPlanID & "' AND ComIDNO = '" & GComIDNO & "' AND SeqNo = '" & GSeqNO & "'" & vbCrLf
        sql &= " AND PlanID=@PlanID and ComIDNO=@ComIDNO and SeqNO=@SeqNO" & vbCrLf
        sql &= " ) b ON a.TechID = b.TechID" & vbCrLf
        sql &= " LEFT JOIN dbo.KEY_DEGREE c ON a.DegreeID = c.DegreeID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.WorkStatus = '1'" & vbCrLf
        sql &= " AND a.RID = '" & rqRID & "'" & vbCrLf
        sql &= " ORDER BY a.TechID" & vbCrLf
        Dim dtT2 As DataTable = DbAccess.GetDataTable(sql, objconn, parms2)
        i_gSeqno = 0
        tbDataGrid22.Visible = False
        If dtT2.Rows.Count > 0 Then
            tbDataGrid22.Visible = True
            DataGrid22.DataSource = dtT2
            DataGrid22.DataBind()
        End If
        'If TechTYPE = "B" Then
        '    'Dim dtT As DataTable = Nothing
        '    dtT = DbAccess.GetDataTable(sql, objconn)
        'End If
        'Return dtT
    End Sub

    '#Region "班級轉入動作--產投使用"

    '異動報名時間。(計算)(預設值)--班級轉入
    'Public Shared Function ChangSEnterDate(ByRef htCC As Hashtable) As Boolean
    '    Dim blnRst As Boolean = True    '可報名(回傳值)
    '    Dim Period As Integer = 0       '可報名期間 (依天)
    '    Dim Period2 As Integer = 0      '可報名期間 (依月)
    '    Dim rqSDate As String = TIMS.GetMyValue2(htCC, "STDate")
    '    If Trim(rqSDate) = "" Then Return blnRst

    '    '開訓日期一個月前
    '    Dim SDateM1 As String = Common.FormatDate(DateAdd(DateInterval.Month, -1, CDate(rqSDate)))
    '    Period2 = DateDiff(DateInterval.Day, CDate(Now), CDate(SDateM1))
    '    Dim SEnterDate As String = ""
    '    Dim FEnterDate As String = ""

    '    '受訓期間超過14天或未超過14天(不分)
    '    If Period2 > 0 Then
    '        '超過一個月 (則報名起日為 開訓日前一個月(30天),報名迄日為開訓前三日)
    '        SEnterDate = Common.FormatDate(SDateM1) & " 12:00" '中午12點
    '        FEnterDate = Common.FormatDate(DateAdd(DateInterval.Day, -3, CDate(rqSDate))) & " 18:00" '下午18點(晚上6點)
    '    Else
    '        '未滿一個月 (則報名起日為 開班轉入日的次日,報名迄日為開訓前三日)
    '        SEnterDate = Common.FormatDate(DateAdd(DateInterval.Day, 1, CDate(Now))) & " 12:00" '中午12點
    '        FEnterDate = Common.FormatDate(DateAdd(DateInterval.Day, -3, CDate(rqSDate))) & " 18:00" '下午18點(晚上6點)
    '    End If
    '    TIMS.SetMyValue2(htCC, "SEnterDate", SEnterDate)
    '    TIMS.SetMyValue2(htCC, "FEnterDate", FEnterDate)

    '    '上架日期(預設值)
    '    'https://jira.turbotech.com.tw/browse/TIMSC-43
    '    '2:.檢核():上架日期不能晚於報名起日, 預設日期值為當日
    '    'If OnShellDate.Text = "" Then '上架日期
    '    '    OnShellDate.Text = Common.FormatDate(DateAdd(DateInterval.Day, 0, CDate(Now)))
    '    '    Common.SetListItem(OnShellDate_HR, 12) '中午12點
    '    '    Common.SetListItem(OnShellDate_MI, 0)
    '    'End If

    '    '報名日期順序有誤
    '    'If DateDiff(DateInterval.Day, CDate(FEnterDate.Text), CDate(SEnterDate.Text)) > 0 Then
    '    '    Dim tmpSEnterDate As String = SEnterDate.Text
    '    '    SEnterDate.Text = FEnterDate.Text
    '    '    FEnterDate.Text = tmpSEnterDate
    '    'End If

    '    '可報名期間
    '    Period = DateDiff(DateInterval.Day, CDate(Now), CDate(rqSDate))
    '    If Period <= 3 Then
    '        '小於、等於 開訓前三天 
    '        blnRst = False '不可報名
    '    End If
    '    Return blnRst
    'End Function

    ''' <summary>
    ''' 儲存前檢核: false:異常/true:正常--班級轉入
    ''' </summary>
    ''' <returns></returns>
    Function xChk_CC1(ByRef V_Errmsg As String) As Boolean
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        '檢核登入者的計畫 異常為False
        If Not TIMS.ChkTPlanID28(sm) Then
            V_Errmsg = cst_errmsg4 'Common.MessageBox(Me, cst_errmsg4)
            Return False 'Exit Sub
        End If

        'Dim sErrorMsg As String = ""
        'Dim rst As Boolean = ChkSTDate(sErrorMsg)
        'If sErrorMsg <> "" Then
        '    Common.MessageBox(Me, sErrorMsg)
        '    Return False 'Exit Sub
        'End If
        'Dim drPP As DataRow = TIMS.GetPPInfo(GPlanID, GComIDNO, GSeqNO, objconn)
        'Hid_clsid.Value = TIMS.ClearSQM(Hid_clsid.Value) '班別代碼
        'Hid_CyclType.Value = TIMS.ClearSQM(Hid_CyclType.Value) '期別
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'rq_OCID = TIMS.ClearSQM(rq_OCID)

        '2006/03/ add conn by matt
        '檢查班級代碼是否被使用過
        'Dim sql As String = ""
        'If rq_OCID <> "" Then
        '    sql = "SELECT 'X' FROM CLASS_CLASSINFO WHERE CLSID='" & clsid.Value & "' and PlanID='" & PlanID.Value & "' and CyclType='" & CyclType.Text & "' and RID='" & RIDValue.Value & "' and OCID!='" & rq_OCID & "'"
        'Else
        '    sql = "SELECT 'X' FROM CLASS_CLASSINFO WHERE CLSID='" & clsid.Value & "' and PlanID='" & PlanID.Value & "' and CyclType='" & CyclType.Text & "' and RID='" & RIDValue.Value & "'"
        'End If
        'Dim dtX As DataTable = DbAccess.GetDataTable(sql, objconn)
        'If dtX.Rows.Count > 0 Then
        '    Common.MessageBox(Me, "新增開班資料重複(該機構在當年度計畫有相同的班別代碼與期別!!)")
        '    Return False 'Exit Sub
        'End If

        Dim sql As String = ""
        sql = ""
        sql &= " SELECT 'X' "
        sql &= " FROM CLASS_CLASSINFO "
        sql &= " WHERE 1=1 "
        sql &= " AND PlanID=@PlanID"
        sql &= " AND ComIDNO=@ComIDNO"
        sql &= " AND SeqNO=@SeqNO"
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("PlanID", GPlanID)
        parms.Add("ComIDNO", GComIDNO)
        parms.Add("SeqNO", GSeqNO)
        Dim dtCC As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dtCC.Rows.Count > 0 Then
            V_Errmsg = "新增開班資料重複(已有轉班資料!!!)" 'Common.MessageBox(Me, "新增開班資料重複(已有轉班資料!!!)")
            Return False 'Exit Sub
        End If
        Return True
    End Function

    ''' <summary>儲存(CLASS_CLASSINFO) --班級轉入</summary>
    Public Sub SaveData1()
        'Dim rq_OCID As String = TIMS.GetMyValue2(htCC, "OCID")
        If Hid_SENTERDATE.Value = "" Then Exit Sub
        If Hid_FENTERDATE.Value = "" Then Exit Sub
        Dim dr1 As DataRow = TIMS.GetPPInfo(GPlanID, GComIDNO, GSeqNO, objconn)
        If dr1 Is Nothing Then Exit Sub

        Dim vRIDValue As String = Convert.ToString(dr1("RID"))
        Dim vRelship As String = TIMS.GET_RelshipforRID(vRIDValue, objconn)

        Dim htPV As New Hashtable 'htPV.Clear()
        htPV.Add("RID", Convert.ToString(dr1("RID")))
        htPV.Add("TMID", Convert.ToString(dr1("TMID")))
        If sm.UserInfo.LID = 0 Then
            Dim drPN As DataRow = TIMS.GetPlanID1(GPlanID, objconn)
            If drPN Is Nothing Then Exit Sub
            htPV.Add("TPLANID", drPN("TPLANID"))
            htPV.Add("DISTID", drPN("DISTID"))
            htPV.Add("YEARS", drPN("YEARS"))
        Else
            htPV.Add("TPLANID", sm.UserInfo.TPlanID)
            htPV.Add("DISTID", sm.UserInfo.DistID)
            htPV.Add("YEARS", sm.UserInfo.Years)
        End If
        htPV.Add("CJOB_UNKEY", Convert.ToString(dr1("CJOB_UNKEY")))
        htPV.Add("CLASSNAME", Convert.ToString(dr1("CLASSNAME")))

        Dim vRID As String = TIMS.GetMyValue2(htPV, "RID")
        Dim vTMID As String = TIMS.GetMyValue2(htPV, "TMID")
        'Dim vRelship As String = TIMS.GetMyValue2(htPV, "Relship")
        Dim vTPLANID As String = TIMS.GetMyValue2(htPV, "TPLANID") '28
        Dim vDISTID As String = TIMS.GetMyValue2(htPV, "DISTID") '001
        Dim vYEARS As String = TIMS.GetMyValue2(htPV, "YEARS") 'YYYY
        Dim vCJOB_UNKEY As String = TIMS.GetMyValue2(htPV, "CJOB_UNKEY")
        Dim vCLASSNAME As String = TIMS.GetMyValue2(htPV, "CLASSNAME")

        'Dim hPMS As New Hashtable From {{"PlanID", GPlanID}, {"ComIDNO", GComIDNO}, {"SeqNo", GSeqNO}}
        'Dim sql As String = ""
        'sql &= " SELECT * FROM PLAN_TEACHER"
        'sql &= " WHERE TechTYPE='A'" 'TechTYPE: A:師資/B:助教
        'sql &= " AND PlanID=@PlanID and ComIDNO=@ComIDNO and SeqNo=@SeqNo"
        'Dim dtPt As DataTable = DbAccess.GetDataTable(sql, objconn, hPMS)

        Dim vClassCount As String = "01" '預設值為01
        vClassCount = TIMS.FmtCyclType(vClassCount)

        'Call TIMS.OpenDbConn(tConn) 'Dim sql As String = ""
        Dim iOCID_New As Integer = -1
        Dim sql As String = ""
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim tTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            '2006/03/ add conn by matt
            iOCID_New = DbAccess.GetNewId(tTrans, "CLASS_CLASSINFO_OCID_SEQ,CLASS_CLASSINFO,OCID") 'fix ora-00001 違反必須唯一的限制條件
            Dim vHtClsid As Hashtable = TIMS.Get_ClassIDG28(htPV, tTrans)
            Dim vCLSID As String = TIMS.GetMyValue2(vHtClsid, "CLSID")
            'Dim vCLSID As String = TIMS.GetMyValue2(vHtClsid, "CLSID")
            If vCLSID = "" Then
                Dim strEx As String = "CLSID 取得資料為空異常!"
                Throw New Exception(strEx)
            End If

            Dim dr As DataRow = Nothing
            Dim dt As DataTable = Nothing
            Dim da As SqlDataAdapter = Nothing
            sql = "SELECT * FROM CLASS_CLASSINFO WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, tTrans)
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = iOCID_New
            dr("CLSID") = vCLSID 'vClsid.Value
            dr("PlanID") = dr1("PlanID")
            dr("Years") = Right(dr1("PlanYear"), 2)
            Dim vCyclType As String = TIMS.FmtCyclType(dr1("CyclType"))
            dr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
            dr("ClassNum") = vClassCount
            dr("RID") = dr1("RID")
            dr("ClassCName") = dr1("ClassName")
            'dr("CJOB_UNKEY") = dr9("CJOB_UNKEY")  '通俗職類
            'dr("ClassEngName") = If(ClassEngName <> "", ClassEngName, "") 'ClassEngName
            dr("Content") = dr1("Content")
            'dr("Purpose") = "一、學科：" & dr1("PurScience") & vbCrLf & "二、術科：" & dr1("PurTech")
            '2007/9/26 修改成將訓練目標帶入即可--Charles
            dr("Purpose") = dr1("PurScience")
            dr("TPropertyID") = "1" '在職VALUE
            dr("TMID") = dr1("TMID")
            dr("CJOB_UNKEY") = dr1("CJOB_UNKEY") '通俗職類
            dr("STDate") = dr1("STDate")
            dr("CheckInDate") = TIMS.Cdate2(dr1("STDate"))
            dr("FTDate") = dr1("FDDate")
            'TPeriod = TIMS.Get_Plan_VerReport(dr1("PlanID"), dr1("ComIDNO"), dr1("SeqNo"), "TPeriod", objconn)
            'If TPeriod = "" Then dr("TPeriod") = Convert.DBNull Else dr("TPeriod") = TPeriod
            dr("TPeriod") = Convert.DBNull
            dr("TaddressZip") = dr1("TaddressZip")
            dr("TaddressZIP6W") = dr1("TaddressZIP6W")
            dr("TAddress") = dr1("TAddress")
            dr("THours") = dr1("THours")
            dr("TNum") = dr1("TNum")
            dr("Relship") = vRelship
            dr("ComIDNO") = dr1("ComIDNO")
            dr("SeqNO") = dr1("SeqNO")
            dr("SEnterDate") = TIMS.Cdate2(Hid_SENTERDATE.Value)
            dr("FEnterDate") = TIMS.Cdate2(Hid_FENTERDATE.Value)
            'If vsOnShellDate <> "" Then dr("OnShellDate") = vsOnShellDate'上架日期
            'dr("SEnterDate") = SEnterDate.Text
            'dr("FEnterDate") = FEnterDate.Text
            'dr("CheckInDate") = If(CheckInDate.Text <> "", CheckInDate.Text, STDate.Text)
            dr("ExamDate") = Convert.DBNull '甄試時段
            dr("ExamPeriod") = Convert.DBNull '甄試時段
            Dim vTDeadline As String = TIMS.Get_TDeadline(CDate(dr1("STDate")), CDate(dr1("FDDate")))
            dr("TDeadline") = vTDeadline 'TDeadline.SelectedValue
            'dr("STDate") = STDate.Text
            'dr("FTDate") = FTDate.Text
            'dr("TaddressZip") = city_code.Value
            'dr("TAddress") = TAddress.Text
            'dr("THours") = THours.Text
            'dr("TNum") = Tnum.Text
            'dr("TPeriod") = TPeriod.SelectedValue
            dr("NotOpen") = "N"
            dr("NORID") = Convert.DBNull
            dr("OtherReason") = Convert.DBNull
            dr("IsApplic") = "N"
            dr("Relship") = vRelship
            dr("ComIDNO") = GComIDNO '.Value '(PCS)
            dr("SeqNO") = GSeqNO '.Value '(PCS)
            dr("IsCalculate") = "Y"
            dr("IsSuccess") = "Y"
            'TechName.Text = TIMS.ClearSQM(TechName.Text)
            dr("IsFullDate") = "N" '產學訓預設值為否
            'dr("CTName") = Left("X:" & TechName.Text, 40)
            'dr("CTName") = Left(TechName.Text, 40)
            dr("CTName") = " "
            dr("IsBusiness") = "N" 'IIf(IsBusiness.Checked = True, "Y", "N")
            'dr("EnterpriseName") = IIf(EnterpriseName.Text <> "", EnterpriseName.Text, Convert.DBNull)
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, tTrans)
            'If rq_OCID <> "" Then iOCID_New = Val(rq_OCID)

            Dim hUPMS As New Hashtable
            hUPMS.Add("PlanID", GPlanID)
            hUPMS.Add("ComIDNO", GComIDNO)
            hUPMS.Add("SeqNo", GSeqNO)
            hUPMS.Add("MODIFYACCT", sm.UserInfo.UserID)
            Dim usSql As String = ""
            usSql &= " UPDATE PLAN_PLANINFO" & vbCrLf
            usSql &= " SET TransFlag='Y',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql &= " WHERE PlanID=@PlanID and ComIDNO=@ComIDNO and SeqNo=@SeqNo" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql, tTrans, hUPMS)

            '假如有學員的話,要更新學員學號--------------------Strat
            'If TBclass.Value <> "" Then
            '    If OldClassID.Value <> TBclass.Value Then  'TBclass_id.Text Then
            '        TBclass_id.Text = TBclass.Value.ToString
            '        sql = "SELECT SOCID,StudentID FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & iOCID_New & "'"
            '        dt = DbAccess.GetDataTable(sql, da, trans)
            '        If dt.Rows.Count <> 0 Then
            '            For Each dr In dt.Rows
            '                dr("StudentID") = Replace(dr("StudentID"), OldClassID.Value, TBclass_id.Text)
            '            Next
            '            DbAccess.UpdateDataTable(dt, da, trans)
            '        End If
            '    End If
            'End If
            '假如有學員的話,要更新學員學號--------------------End
            'DbAccess.RollbackTrans(trans)
            DbAccess.CommitTrans(tTrans)

        Catch ex As Exception
            DbAccess.RollbackTrans(tTrans)
            Common.MessageBox(Me, "儲存失敗!!")
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Call TIMS.CloseDbConn(tConn)
            'DbAccess.RollbackTrans(trans)
            'Throw ex
        End Try

        If iOCID_New > 0 Then
            '儲存 班級申請老師(CLASS_TEACHER)
            Call SAVE_CLASS_TEACHER(iOCID_New, objconn)
            'Call TIMS.CloseDbConn(tConn)
        End If

        '重複 為 true
        Dim Double_flag As Boolean = False 'false 沒有重複。
        Dim iDouble As Integer = 0
        Do
            iDouble += 1
            Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
            '刪除重複轉班資料。
            Try
                '至少判斷1次是否有重複轉班
                Double_flag = TIMS.sUtl_DeleteDoubleClassInfo(GPlanID, GComIDNO, GSeqNO, objconn)
            Catch ex As Exception
                Double_flag = False '只要有1次失敗就算了吧。
            End Try
            '判斷5次也太 ...
            If iDouble >= 5 Then Double_flag = False
        Loop Until Not Double_flag '直到沒有重複。

        'If Not ViewState("ClassSearchStr") Is Nothing Then
        '    Session("ClassSearchStr") = ViewState("ClassSearchStr")
        'End If
        'Common.RespWrite(Me, "<script>alert('儲存成功!');location.href='TC_04_002.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    '#End Region

    Private Sub DataGrid21_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid21.ItemDataBound

        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim ProLicense As Label = e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEA As HtmlInputButton = e.Item.FindControl("btn_TCTYPEA")
                'Dim rqRID As String = sm.UserInfo.RID
                'sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=A&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                'btn_TCTYPEA.Attributes("onclick") = sWOScript1

                HidTechID.Value = Convert.ToString(drv("TechID"))
                i_gSeqno += 1
                seqno.Text = i_gSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                ProLicense.Text = Convert.ToString(drv("ProLicense"))
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                'TeacherDesc.ReadOnly = False
                'btn_TCTYPEA.Visible = True

                'Select Case rqProcessType 'ProcessType @Insert/Update/View
                '    Case cst_ptView '查詢功能不提供儲存
                '        TeacherDesc.ReadOnly = True
                '        btn_TCTYPEA.Visible = False
                'End Select

                'Dim flag_can_save As Boolean = True
                'If RIDValue.Value = "" Then flag_can_save = False '不同單位 不提供儲存
                'If sm.UserInfo.RID <> RIDValue.Value Then flag_can_save = False '不同單位 不提供儲存
                ''不同單位 不提供儲存
                'If Not flag_can_save Then
                '    TeacherDesc.ReadOnly = True
                '    btn_TCTYPEA.Visible = False
                'End If
                '不提供儲存
                TeacherDesc.ReadOnly = True
                btn_TCTYPEA.Visible = False
        End Select
    End Sub

    Private Sub DataGrid22_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid22.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim ProLicense As Label = e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEB As HtmlInputButton = e.Item.FindControl("btn_TCTYPEB")
                'Dim rqRID As String = sm.UserInfo.RID
                'sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=B&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                'btn_TCTYPEB.Attributes("onclick") = sWOScript1
                HidTechID.Value = Convert.ToString(drv("TechID"))
                i_gSeqno += 1
                seqno.Text = i_gSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                ProLicense.Text = Convert.ToString(drv("ProLicense"))
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                'TeacherDesc.ReadOnly = False
                'btn_TCTYPEB.Visible = True

                'Dim flag_can_save As Boolean = True
                'If RIDValue.Value = "" Then flag_can_save = False '不同單位 不提供儲存
                'If sm.UserInfo.RID <> RIDValue.Value Then flag_can_save = False '不同單位 不提供儲存
                ''不同單位 不提供儲存
                'If Not flag_can_save Then
                '    TeacherDesc.ReadOnly = True
                '    btn_TCTYPEB.Visible = False
                'End If
                '不提供儲存
                TeacherDesc.ReadOnly = True
                btn_TCTYPEB.Visible = False
                'Select Case rqProcessType 'ProcessType @Insert/Update/View
                '    Case cst_ptView '查詢功能不提供儲存
                '        TeacherDesc.ReadOnly = True
                '        btn_TCTYPEB.Visible = False
                'End Select
        End Select

    End Sub

    Sub SAVE_CLASS_TEACHER(ByVal iOCID As Integer, ByVal tConn As SqlConnection)

        Dim dParms As New Hashtable From {{"OCID", iOCID}}
        Dim dSql As String = "DELETE CLASS_TEACHER WHERE OCID =@OCID"
        DbAccess.ExecuteNonQuery(dSql, tConn, dParms)

        Dim iSqlc As String = ""
        iSqlc &= " INSERT INTO CLASS_TEACHER (CTRID ,OCID,TECHID,MODIFYACCT,MODIFYDATE,TECHTYPE,TEACHERDESC)" & vbCrLf
        iSqlc &= " VALUES (@CTRID ,@OCID,@TECHID,@MODIFYACCT,GETDATE(),@TECHTYPE,@TEACHERDESC )" & vbCrLf

        'Dim sParms As New Hashtable
        Dim sSql1 As String = ""
        sSql1 = " SELECT 1 FROM CLASS_TEACHER WHERE OCID=@OCID AND TECHID=@TECHID AND TECHTYPE=@TECHTYPE" & vbCrLf

        Const cst_iMaxLen_TeacherDesc As Integer = 500
        '更新師資表 'TechTYPE: A:師資/B:助教
        Const cst_tTECHTYPE_A As String = "A"
        Const cst_tTECHTYPE_B As String = "B"

        '更新師資表 cst_tTECHTYPE_A
        For Each eItem As DataGridItem In DataGrid21.Items
            Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            'Dim seqno As Label = eItem.FindControl("seqno")
            'Dim TeachCName As Label = eItem.FindControl("TeachCName")
            'Dim DegreeName As Label = eItem.FindControl("DegreeName")
            'Dim major As Label = eItem.FindControl("major")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
            'Dim btn_TCTYPEA As HtmlInputButton = eItem.FindControl("btn_TCTYPEA")
            If HidTechID.Value <> "" Then
                Dim sParms As New Hashtable
                sParms.Add("OCID", iOCID)
                sParms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                sParms.Add("TECHTYPE", cst_tTECHTYPE_A) 'TechTYPE: A:師資/B:助教
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms)
                If dr1 Is Nothing Then
                    Dim iCTRID As Integer = DbAccess.GetNewId(tConn, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                    Dim parms As New Hashtable
                    parms.Add("CTRID", iCTRID)
                    parms.Add("OCID", iOCID)
                    parms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                    parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    parms.Add("TECHTYPE", cst_tTECHTYPE_A)
                    parms.Add("TEACHERDESC", tTEACHERDESC)
                    DbAccess.ExecuteNonQuery(iSqlc, tConn, parms)
                End If
            End If
        Next

        '更新師資表(助教) cst_tTECHTYPE_B
        For Each eItem As DataGridItem In DataGrid22.Items
            Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            'Dim seqno As Label = eItem.FindControl("seqno")
            'Dim TeachCName As Label = eItem.FindControl("TeachCName")
            'Dim DegreeName As Label = eItem.FindControl("DegreeName")
            'Dim major As Label = eItem.FindControl("major")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
            'Dim btn_TCTYPEA As HtmlInputButton = eItem.FindControl("btn_TCTYPEA")
            If HidTechID.Value <> "" Then
                Dim sParms As New Hashtable
                sParms.Add("OCID", iOCID)
                sParms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                sParms.Add("TECHTYPE", cst_tTECHTYPE_B) 'TechTYPE: A:師資/B:助教
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms)
                If dr1 Is Nothing Then
                    Dim iCTRID As Integer = DbAccess.GetNewId(tConn, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                    Dim parms As New Hashtable
                    parms.Add("CTRID", iCTRID)
                    parms.Add("OCID", iOCID)
                    parms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                    parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    parms.Add("TECHTYPE", cst_tTECHTYPE_B)
                    parms.Add("TEACHERDESC", tTEACHERDESC)
                    DbAccess.ExecuteNonQuery(iSqlc, tConn, parms)
                End If
            End If
        Next

        '更新師資表 -End
    End Sub

End Class
