Public Class TC_01_001_add
    Inherits AuthBasePage

    Const cst_s_type_plankind As String = "plankind"
    Const cst_s_type_FlexTurnoutKind As String = "FlexTurnoutKind"
    Const cst_s_type_FlexTurnoutKind_N As String = "FlexTurnoutKind_N"
    Const cst_FlexTurnoutKind_CLOSE As String = "CLOSE"
    Const cst_FlexTurnoutKind_OPEN As String = "OPEN"

    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()
    Dim objconn As SqlConnection

    Function Get_ListCT_OBJ(ByRef obj As ListControl, ByVal s_type As String) As ListControl
        If obj Is Nothing Then Return obj
        obj.Items.Clear()

        With obj.Items
            Select Case s_type
                Case cst_s_type_plankind
                    .Insert(0, New ListItem("自辦", 1))
                    .Insert(1, New ListItem("委外", 2))
                Case cst_s_type_FlexTurnoutKind
                    '20080610 Andy  彈性調整出缺勤(開放 「婚假是否列入缺曠課時數」 選取功能)
                    .Insert(0, New ListItem("否", 0))
                    .Insert(1, New ListItem("是", 1))
                Case cst_s_type_FlexTurnoutKind_N
                    .Insert(0, New ListItem("否", 0))
            End Select
        End With

        Return obj
    End Function



    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("ID_PLAN")
        TIMS.sUtl_SetMaxLen(dt, "SPONSOR", main_center)
        TIMS.sUtl_SetMaxLen(dt, "COSPONSOR", sub_center)
        TIMS.sUtl_SetMaxLen(dt, "SUBTITLE", SubTitle)
        TIMS.sUtl_SetMaxLen(dt, "PCOMMENT", PComment)
    End Sub

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

        Hid_rqeditid.Value = TIMS.ClearSQM(Request("editid"))

        If Not Page.IsPostBack Then
            DistValue = TIMS.Get_DistID(DistValue)
            PlanKind = Get_ListCT_OBJ(PlanKind, cst_s_type_plankind)
            '20080610 Andy  彈性調整出缺勤(開放 「婚假是否列入缺曠課時數」 選取功能)
            FlexTurnoutKind = Get_ListCT_OBJ(FlexTurnoutKind, cst_s_type_FlexTurnoutKind)

            yearlist_add = TIMS.GetSyear(yearlist_add)
            Call Show_KeyPlan()
            Hid_rqeditid.Value = TIMS.ClearSQM(Hid_rqeditid.Value)
            If Hid_rqeditid.Value <> "" Then Call sShowData1(Hid_rqeditid.Value)
        End If

        If Not Session("_search") Is Nothing Then Me.ViewState("_search") = Session("_search") 'Session("_search") = Nothing
    End Sub

    ''' <summary>
    ''' SHOW ID_PLAN
    ''' </summary>
    ''' <param name="vPlanid"></param>
    Sub sShowData1(ByVal vPlanid As String)
        vPlanid = TIMS.ClearSQM(vPlanid)
        If vPlanid = "" Then Exit Sub

        Me.DistValue.Enabled = False '修改鎖轄區分署(轄區中心)
        'Dim selreader As SqlDataReader
        Dim selsql As String = " SELECT * FROM ID_PLAN WHERE PLANID =@PLANID " '& vPlanid 'Hid_rqeditid.Value '@PlanID
        Dim s_parms As New Hashtable
        s_parms.Clear()
        s_parms.Add("PLANID", vPlanid)
        Dim dr1 As DataRow = DbAccess.GetOneRow(selsql, objconn, s_parms)
        If dr1 Is Nothing Then Exit Sub

        'selreader = DbAccess.GetReader(selsql, objconn)
        'objconn.Close()
        Common.SetListItem(DistValue, Convert.ToString(dr1("DistID")))
        Common.SetListItem(yearlist_add, Convert.ToString(dr1("years")))
        Common.SetListItem(planlist_add, Convert.ToString(dr1("TPlanID")))

        seqno.Text = Convert.ToString(dr1("seq"))
        main_center.Text = Convert.ToString(dr1("Sponsor"))
        sub_center.Text = Convert.ToString(dr1("Cosponsor"))
        '調整[時效起、迄日]欄位要顯示西元日期or民國日期，by:20181018
        start_date.Text = If(flag_ROC, TIMS.Cdate17(dr1("SDate")), TIMS.Cdate3(dr1("SDate")))
        end_date.Text = If(flag_ROC, TIMS.Cdate17(dr1("EDate")), TIMS.Cdate3(dr1("EDate")))

        Common.SetListItem(PlanKind, Convert.ToString(dr1("PlanKind")))
        hidPlanKind.Value = Convert.ToString(dr1("PlanKind"))

        PComment.Text = Convert.ToString(dr1("PComment"))
        SubTitle.Text = Convert.ToString(dr1("SubTitle"))

        '是否下放轄區分署(轄區中心)決定【e網報名審核發送Email】
        If Not TIMS.CheckReusable(Convert.ToString(dr1("TPlanID")), objconn) Then
            '沒有設定權
            Tr18.Style("display") = "none"
        Else
            '有設定權
            Tr18.Style("display") = ""
            If Convert.ToString(dr1("EmailSend")) <> "" Then
                Common.SetListItem(R18, Convert.ToString(dr1("EmailSend")))
            Else
                Hid_rqeditid.Value = TIMS.ClearSQM(Hid_rqeditid.Value)
                Dim flagES As Boolean = TIMS.CheckEmailSend(Me, Convert.ToString(dr1("TPlanID")), Hid_rqeditid.Value, objconn)
                Dim vR18 As String = If(flagES, "Y", "N")
                Common.SetListItem(R18, vR18)
            End If
        End If

        '進入修改模式年度與計畫不可修改
        Me.yearlist_add.Enabled = False
        Me.planlist_add.Enabled = False

        '---------------20080610 Andy  彈性調整出缺勤(開放 「婚假是否列入缺曠課時數」 選取功能) ,"FlexTurnoutKind"=null，開放 ；1,不開放
        Hid_FlexTurnoutKind_OPEN_CLOSE.Value = cst_FlexTurnoutKind_CLOSE
        Select Case Convert.ToString(dr1("Plankind"))
            Case "2" '委外
                FlexTurnoutKind = Get_ListCT_OBJ(FlexTurnoutKind, cst_s_type_FlexTurnoutKind_N)
                Common.SetListItem(FlexTurnoutKind, "0")
                'Hid_FlexTurnoutKind_OPEN_CLOSE.Value = cst_FlexTurnoutKind_CLOSE
            Case "1" '自辦
                FlexTurnoutKind = Get_ListCT_OBJ(FlexTurnoutKind, cst_s_type_FlexTurnoutKind)
                Common.SetListItem(FlexTurnoutKind, "0")
                '彈性調整是否開放
                Hid_FlexTurnoutKind_OPEN_CLOSE.Value = If(Convert.ToString(dr1("FlexTurnoutKind")) <> "", cst_FlexTurnoutKind_OPEN, cst_FlexTurnoutKind_CLOSE)
            Case Else
                FlexTurnoutKind = Get_ListCT_OBJ(FlexTurnoutKind, cst_s_type_FlexTurnoutKind)
                Common.SetListItem(FlexTurnoutKind, "0")
                'Hid_FlexTurnoutKind_OPEN_CLOSE.Value = cst_FlexTurnoutKind_CLOSE
        End Select

    End Sub

    Public Shared Function GET_SEQNO_MAX1(ByVal oConn As SqlConnection, ByVal YEARS As String, ByVal TPlanID As String, ByVal DistID As String, ByRef errMsg As String) As String
        Dim rst As String = "001"
        Dim sql As String = ""
        sql = "Select MAX(SEQ)+1 MAX FROM ID_PLAN WHERE YEARS =@YEARS  And TPlanID =@TPlanID And DistID =@DistID"
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("YEARS", YEARS)
        parms.Add("TPlanID", TPlanID)
        parms.Add("DistID", DistID)
        Dim dr As DataRow = DbAccess.GetOneRow(sql, oConn, parms)
        If Not IsDBNull(dr("max")) Then
            If Int(dr("max")) > 999 Then
                errMsg = "目前序號超過999，不能新增資料!"
                Return ""
                'Common.MessageBox(Me, "目前序號超過999，不能新增資料!")
                'Exit Sub
            ElseIf Int(dr("max")) < 10 Then
                rst = "00" & dr("max")
            ElseIf Int(dr("max")) < 100 Then
                rst = "0" & dr("max")
            Else
                rst = dr("max")
            End If
        End If
        Return rst
    End Function

    Function Checkdata1() As Boolean
        Dim rst As Boolean = False 'false:異常/true:正常
        Dim v_yearlist_add As String = TIMS.GetListValue(yearlist_add) '.ClearSQM(yearlist_add.SelectedValue)
        Dim v_planlist_add As String = TIMS.GetListValue(planlist_add) 'ClearSQM(planlist_add.SelectedValue)
        Dim v_DistValue As String = TIMS.GetListValue(DistValue) 'ClearSQM(DistValue.SelectedValue)
        Dim v_PlanKind As String = TIMS.GetListValue(PlanKind) '.SelectedValue
        Dim v_FlexTurnoutKind As String = TIMS.GetListValue(FlexTurnoutKind) '.SelectedValue

        '#Region "儲存"
        Dim sErrMsg As String = ""
        If v_yearlist_add = "" Then sErrMsg &= "請選擇年度" & vbCrLf
        If v_planlist_add = "" Then sErrMsg &= "請選擇訓練計畫" & vbCrLf
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        If start_date.Text = "" Then sErrMsg &= "請輸入起日" & vbCrLf
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        If end_date.Text = "" Then sErrMsg &= "請輸入迄日" & vbCrLf
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return rst '  Exit Sub
        End If

        Dim myStartDate As String = ""
        Dim myEndDate As String = ""
        Try
            '-------------------- 檢核(西元/民國)日期格式是否正確，by:20181001  start
            If flag_ROC Then
                myStartDate = TIMS.Cdate3(TIMS.Cdate18(start_date.Text))
                myEndDate = TIMS.Cdate3(TIMS.Cdate18(end_date.Text))
                If Not TIMS.IsDate7(start_date.Text) Then sErrMsg &= "起日-日期格式不正確" & vbCrLf
                If Not TIMS.IsDate7(end_date.Text) Then sErrMsg &= "迄日-日期格式不正確" & vbCrLf
            Else
                myStartDate = TIMS.Cdate3(start_date.Text)
                myEndDate = TIMS.Cdate3(end_date.Text)
                If Not TIMS.IsDate1(start_date.Text) Then sErrMsg &= "起日-日期格式不正確" & vbCrLf
                If Not TIMS.IsDate1(end_date.Text) Then sErrMsg &= "迄日-日期格式不正確" & vbCrLf
            End If
            '-------------------- 檢核(西元/民國)日期格式是否正確，by:20181001  end
        Catch ex As Exception
            sErrMsg &= "起日或迄日日期格式不正確" & vbCrLf
        End Try
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return rst '  Exit Sub
        End If

        'sErrMsg &= "迄日不得小於起日或日期格式不正確" & vbCrLf
        If DateDiff(DateInterval.Day, CDate(myStartDate), CDate(myEndDate)) < 0 Then sErrMsg &= "迄日不得小於起日" & vbCrLf
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return rst '  Exit Sub
        End If

        If v_PlanKind = "" Then sErrMsg &= "請選擇計畫總類" & vbCrLf
        If v_FlexTurnoutKind = "" Then sErrMsg &= "請選擇彈性調整出缺勤" & vbCrLf
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return rst '  Exit Sub
        End If

        hidPlanKind.Value = TIMS.ClearSQM(hidPlanKind.Value)
        Hid_rqeditid.Value = TIMS.ClearSQM(Hid_rqeditid.Value)
        If Hid_rqeditid.Value <> "" AndAlso hidPlanKind.Value <> "" AndAlso v_PlanKind <> "" Then
            '修改自辦、委外計畫，先判斷該計畫是否已有使用經費 計價種類
            If v_PlanKind <> hidPlanKind.Value Then
                '偵測是否已使用計價種類。
                If TIMS.ChkPlanCostItem(Me.Hid_rqeditid.Value, hidPlanKind.Value) Then
                    Common.MessageBox(Me, "該經費計價方式 已有資料不可任意修改 計畫種類!!")
                    Return rst 'Exit Sub
                End If
            End If
        End If

        If Hid_rqeditid.Value <> "" Then
            Dim sqlstr As String = "SELECT * FROM ID_PLAN WHERE PLANID = " & Hid_rqeditid.Value
            Dim sqldr As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
            If sqldr Is Nothing Then
                Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                Return rst 'Exit Sub
            End If
        End If
        If sErrMsg <> "" Then Return rst
        rst = True '無錯誤訊息
        Return rst
    End Function

    Function sSaveData1() As Boolean
        Dim rst As Boolean = True 'false:異常/true:正常
        Dim tYEARS As String = TIMS.GetListValue(yearlist_add) '.ClearSQM(yearlist_add.SelectedValue)
        Dim tTPlanID As String = TIMS.GetListValue(planlist_add) 'ClearSQM(planlist_add.SelectedValue)
        Dim tDistID As String = TIMS.GetListValue(DistValue) 'ClearSQM(DistValue.SelectedValue)
        Dim v_PlanKind As String = TIMS.GetListValue(PlanKind) '.SelectedValue
        Dim v_FlexTurnoutKind As String = TIMS.GetListValue(FlexTurnoutKind) '.SelectedValue
        Dim v_R18 As String = TIMS.GetListValue(R18) '.SelectedValue

        main_center.Text = TIMS.ClearSQM(main_center.Text)
        sub_center.Text = TIMS.ClearSQM(sub_center.Text)
        SubTitle.Text = TIMS.ClearSQM(SubTitle.Text) '備註顯示
        PComment.Text = TIMS.ClearSQM(PComment.Text) '計畫說明

        seqno.Text = TIMS.ClearSQM(seqno.Text)
        Dim seq_NO As String = seqno.Text '原值

        Hid_rqeditid.Value = TIMS.ClearSQM(Hid_rqeditid.Value)
        If Hid_rqeditid.Value = "" Then
            '產生新值
            '新增使用 '同一年度裡不能有相同訓練計畫及序號!!!
            If tDistID = "" Then tDistID = sm.UserInfo.DistID
            Dim errMsg As String = ""
            seq_NO = GET_SEQNO_MAX1(objconn, tYEARS, tTPlanID, tDistID, errMsg)
            If errMsg <> "" Then
                seq_NO = ""
                Common.MessageBox(Me, errMsg)
                Return False 'Exit Sub
            End If
        End If

        Dim myStartDate As String = ""
        Dim myEndDate As String = ""
        myStartDate = If(flag_ROC, TIMS.Cdate3(TIMS.Cdate18(start_date.Text)), TIMS.Cdate3(start_date.Text))
        myEndDate = If(flag_ROC, TIMS.Cdate3(TIMS.Cdate18(end_date.Text)), TIMS.Cdate3(end_date.Text))

        Dim parms As New Hashtable
        Dim sql As String = ""
        Dim iPLANID As Integer = 0
        If Hid_rqeditid.Value = "" Then
            sql = ""
            sql &= " INSERT INTO ID_PLAN (PLANID ,YEARS ,DISTID ,TPLANID ,SEQ ,SPONSOR ,COSPONSOR ,SDATE ,EDATE ,PLANKIND ,MODIFYACCT ,MODIFYDATE ,FLEXTURNOUTKIND ,EMAILSEND ,SUBTITLE ,PCOMMENT)" & vbCrLf
            sql &= " VALUES (@PLANID ,@YEARS ,@DISTID ,@TPLANID ,@SEQ ,@SPONSOR ,@COSPONSOR ,@SDATE ,@EDATE ,@PLANKIND ,@MODIFYACCT ,GETDATE() ,@FLEXTURNOUTKIND ,@EMAILSEND ,@SUBTITLE ,@PCOMMENT)" & vbCrLf
            'Dim parms As New Hashtable
            iPLANID = DbAccess.GetNewId(objconn, "ID_PLAN_PLANID_SEQ,ID_PLAN,PLANID")
            parms.Clear()
            parms.Add("PLANID", iPLANID)
            parms.Add("YEARS", tYEARS)
            parms.Add("DISTID", tDistID)
            parms.Add("TPLANID", tTPlanID)
            parms.Add("SEQ", seq_NO)
            parms.Add("SPONSOR", main_center.Text)
            parms.Add("COSPONSOR", sub_center.Text)
            parms.Add("SDATE", CDate(myStartDate))
            parms.Add("EDATE", CDate(myEndDate))
            parms.Add("PLANKIND", v_PlanKind)
            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            'parms.Add("MODIFYDATE", MODIFYDATE.Text)
            '彈性調整出缺勤(開放 「婚假是否列入缺曠課時數」 選取功能
            parms.Add("FLEXTURNOUTKIND", If(v_FlexTurnoutKind = "1", v_FlexTurnoutKind, Convert.DBNull))
            parms.Add("EMAILSEND", If(v_R18 <> "", v_R18, Convert.DBNull))
            parms.Add("SUBTITLE", SubTitle.Text) '備註顯示
            parms.Add("PCOMMENT", PComment.Text) '計畫說明
            DbAccess.ExecuteNonQuery(sql, objconn, parms)

        Else
            sql = ""
            sql &= " UPDATE ID_PLAN" & vbCrLf
            sql &= " SET SPONSOR = @SPONSOR" & vbCrLf
            sql &= " ,COSPONSOR = @COSPONSOR" & vbCrLf
            sql &= " ,SDATE = @SDATE" & vbCrLf
            sql &= " ,EDATE = @EDATE" & vbCrLf
            sql &= " ,PLANKIND = @PLANKIND" & vbCrLf
            sql &= " ,MODIFYACCT = @MODIFYACCT" & vbCrLf
            sql &= " ,MODIFYDATE = GETDATE()" & vbCrLf
            sql &= " ,FLEXTURNOUTKIND = @FLEXTURNOUTKIND" & vbCrLf
            sql &= " ,EMAILSEND = @EMAILSEND" & vbCrLf
            sql &= " ,SUBTITLE = @SUBTITLE" & vbCrLf
            sql &= " ,PCOMMENT = @PCOMMENT" & vbCrLf
            sql &= " WHERE PLANID = @PLANID" & vbCrLf
            sql &= " AND YEARS = @YEARS" & vbCrLf
            sql &= " AND DISTID = @DISTID" & vbCrLf
            sql &= " AND TPLANID = @TPLANID" & vbCrLf
            sql &= " AND SEQ = @SEQ" & vbCrLf

            iPLANID = Val(Hid_rqeditid.Value)
            parms.Clear()
            parms.Add("SPONSOR", main_center.Text)
            parms.Add("COSPONSOR", sub_center.Text)
            parms.Add("SDATE", CDate(myStartDate))
            parms.Add("EDATE", CDate(myEndDate))
            parms.Add("PLANKIND", v_PlanKind)
            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            'parms.Add("MODIFYDATE", MODIFYDATE.Text)
            '彈性調整出缺勤(開放 「婚假是否列入缺曠課時數」 選取功能
            parms.Add("FLEXTURNOUTKIND", If(v_FlexTurnoutKind = "1", v_FlexTurnoutKind, Convert.DBNull))
            parms.Add("EMAILSEND", If(v_R18 <> "", v_R18, Convert.DBNull))
            parms.Add("SUBTITLE", SubTitle.Text) '備註顯示
            parms.Add("PCOMMENT", PComment.Text) '計畫說明
            parms.Add("PLANID", iPLANID)
            parms.Add("YEARS", tYEARS)
            parms.Add("DISTID", tDistID)
            parms.Add("TPLANID", tTPlanID)
            parms.Add("SEQ", seq_NO)
            DbAccess.ExecuteNonQuery(sql, objconn, parms)
        End If
        Return rst
    End Function

    '儲存
    Private Sub bt_addrow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_addrow.Click
        If Not Checkdata1() Then Exit Sub  '檢核 false 有異常'(內含錯誤訊息彈跳)
        If Not sSaveData1() Then Exit Sub  '儲存有誤 '(內含錯誤訊息彈跳)
        If ViewState("_search") IsNot Nothing AndAlso Session("_search") Is Nothing Then Session("_search") = Me.ViewState("_search")
        TIMS.Utl_Redirect1(Me, "TC_01_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")))
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ViewState("_search") IsNot Nothing AndAlso Session("_search") Is Nothing Then Session("_search") = Me.ViewState("_search")
        TIMS.Utl_Redirect1(Me, "TC_01_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")))
    End Sub

    Private Sub PlanKind_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PlanKind.SelectedIndexChanged
        Dim v_PlanKind As String = TIMS.GetListValue(PlanKind) '.SelectedValue
        Select Case v_PlanKind
            Case "1" '自辦
                FlexTurnoutKind = Get_ListCT_OBJ(FlexTurnoutKind, cst_s_type_FlexTurnoutKind)
                Select Case Hid_FlexTurnoutKind_OPEN_CLOSE.Value
                    Case cst_FlexTurnoutKind_OPEN
                        Common.SetListItem(FlexTurnoutKind, "1")
                    Case cst_FlexTurnoutKind_CLOSE
                        Common.SetListItem(FlexTurnoutKind, "0")
                End Select
            Case "2" '委外
                FlexTurnoutKind = Get_ListCT_OBJ(FlexTurnoutKind, cst_s_type_FlexTurnoutKind_N)
                Common.SetListItem(FlexTurnoutKind, "0")
        End Select
    End Sub

    '年度選擇後重新搜尋計畫 SQL
    Sub Show_KeyPlan()
        Dim vYears As String = sm.UserInfo.Years
        Dim v_yearlist_add As String = TIMS.GetListValue(yearlist_add) '.ClearSQM(yearlist_add.SelectedValue)
        If v_yearlist_add <> "" Then vYears = v_yearlist_add '(將vYears 值置換為選擇值)

        '含不啟用的計畫
        Dim sqlstr As String = ""
        sqlstr &= " SELECT TPlanID, PlanName + CASE WHEN Clsyear IS NULL OR Clsyear > '" & vYears & "' THEN '' ELSE '…已停用' + Clsyear END PlanName"
        sqlstr &= " FROM Key_Plan"
        sqlstr &= " WHERE 1=1" & vbCrLf
        'vYears = v_yearlist_add
        If v_yearlist_add <> "" Then sqlstr &= " AND (Clsyear IS NULL OR Clsyear > '" & vYears & "')" & vbCrLf
        sqlstr &= " ORDER BY TPlanID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn)

        With planlist_add
            .DataSource = dt
            .DataTextField = "PlanName"
            .DataValueField = "TPlanID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        If sm.UserInfo.LID <> 0 Then
            With planlist_add.Items
                .Remove(planlist_add.Items.FindByValue("08"))
                .Remove(planlist_add.Items.FindByValue("09"))
                .Remove(planlist_add.Items.FindByValue("19"))
                .Remove(planlist_add.Items.FindByValue("20"))
            End With
        End If
    End Sub

    Private Sub yearlist_add_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles yearlist_add.SelectedIndexChanged
        Call Show_KeyPlan()
    End Sub

    Private Sub planlist_add_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles planlist_add.SelectedIndexChanged
        Tr18.Style("display") = "none" '沒有設定權
        Dim v_planlist_add As String = TIMS.GetListValue(planlist_add) 'ClearSQM(planlist_add.SelectedValue)
        Dim flag_Reusable As Boolean = TIMS.CheckReusable(v_planlist_add, objconn)

        If flag_Reusable Then
            'Tr18.Style("display") = "inline" '有設定權
            Tr18.Style("display") = "" '有設定權
            Dim flagES As Boolean = TIMS.CheckEmailSend(Me, v_planlist_add, "", objconn)
            Dim v_R18_YN As String = "N"
            If flagES Then v_R18_YN = "Y"

            Common.SetListItem(R18, v_R18_YN)
        End If
    End Sub

End Class