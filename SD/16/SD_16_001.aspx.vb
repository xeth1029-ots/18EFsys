Partial Class SD_16_001
    Inherits AuthBasePage

    'Const cst_c36 As String = "36"
    'Const cst_c37 As String = "37"
    'Const cst_c99 As String = "99" '其他
    'Const cst_cN36 As String = "第36點"
    'Const cst_cN37 As String = "第37點"
    'Const cst_cN99 As String = "其他"

    'SELECT SBTERMS FROM STUD_BLACKLIST where rownum <=10 --SBTERMS
    Const Cst_stradd As String = "學員處分功能新增"
    Const Cst_strview As String = "學員處分功能檢視"
    Const Cst_strUPDATE As String = "黑名單資料修改"
    Const Cst_strsearch As String = "學員處分功能查詢"
    Const Cst_事由位置 As Integer = 8
    Const Cst_功能欄位 As Integer = 9

    Dim sMemo As String = "" '(查詢原因)
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = dg_Sch

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            Call SCreate1()
        End If

    End Sub

    '第1次載入
    Sub SCreate1()
        '作業顯示模式：0:其他 1:模糊顯示 2:正常顯示
        rblWorkMode.Enabled = False
        Common.SetListItem(rblWorkMode, TIMS.cst_wmdip1)
        TIMS.Tooltip(rblWorkMode, "全計畫-學員處分功能-身分證號-隱碼顯示")

        '取出鍵詞-查詢原因-INQUIRY
        Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        '階層代碼 0:署 1:中心 2:委訓(待確認是否還有??)
        Hid_LID.Value = sm.UserInfo.LID

        ddlTPlanIDSch = TIMS.Get_TPlan(ddlTPlanIDSch, , , , , objconn)
        Common.SetListItem(ddlTPlanIDSch, sm.UserInfo.TPlanID)

        ddlTPlanID = TIMS.Get_TPlan(ddlTPlanID, , , , , objconn)
        Common.SetListItem(ddlTPlanID, sm.UserInfo.TPlanID)

        DistID = TIMS.Get_DistID(DistID)
        ddl_DistID = TIMS.Get_DistID(ddl_DistID)
        Common.SetListItem(ddl_DistID, sm.UserInfo.DistID)

        'ddl_DistID.SelectedValue = sm.UserInfo.DistID
        If sm.UserInfo.LID <> 0 Then ddl_DistID.Enabled = False

        Years = TIMS.Get_Years(Years)
        Years.Items.Insert(0, New ListItem("==請選擇==", ""))
        '處份綠由
        UTL_SBTERMS(ddlSBTERMS, ddl_SBYears, sm.UserInfo.TPlanID)

        tb_Sch.Visible = False

        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        Dim s_javascript_btn6 As String = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
        Button6.Attributes("onclick") = s_javascript_btn6
        btn_Save.Attributes("onclick") = "javascript:return checkSave()"

    End Sub

    '顯示狀況
    Sub SUtl_PanelList(ByVal iType As Integer)
        'iType:1 搜尋 2:'新增/修改 3:'檢視
        Panel1.Visible = False '搜尋
        Panel2.Visible = False '新增/修改
        Panel3.Visible = False
        Select Case iType
            Case 1
                Panel1.Visible = True '搜尋
            Case 2
                Panel2.Visible = True '新增/修改
            Case 3
                Panel3.Visible = True '檢視
        End Select
        'Panel
    End Sub

    '查詢鈕
    Private Sub Btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch)

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Call SSearch1("")
        'Sch_Mark.Value = "1"
    End Sub

    '查詢原因
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        '計畫別, ddlTPlanIDSch,原處分分署, DistID,處分年度, Years,身分證號碼, IDNO,學員姓名, Name,
        Dim V_ddlTPlanIDSch As String = TIMS.GetListValue(ddlTPlanIDSch)
        Dim V_DistID As String = TIMS.GetListValue(DistID)
        Dim V_Years As String = TIMS.GetListValue(Years)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Name.Text = TIMS.ClearSQM(Name.Text)

        If V_ddlTPlanIDSch <> "" Then RstMemo &= String.Concat("&計畫別=", V_ddlTPlanIDSch)
        If V_DistID <> "" Then RstMemo &= String.Concat("&原處分分署=", V_DistID)
        If V_Years <> "" Then RstMemo &= String.Concat("&處分年度=", V_Years)
        If IDNO.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", IDNO.Text)
        If Name.Text <> "" Then RstMemo &= String.Concat("&學員姓名=", Name.Text)
        Return RstMemo
    End Function

    '查詢sub
    Sub SSearch1(ISEXP1 As String)
        'iType :0:一般查詢/1:匯出(查詢結果，將身分證字號中間6碼調整為隱碼顯示)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        HidSBSN.Value = ""
        HidvsType.Value = "" 'HidvsType I:新增/V:檢視/E:編輯

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        Name.Text = TIMS.ClearSQM(Name.Text)

        Dim v_ddlTPlanIDSch As String = TIMS.GetListValue(ddlTPlanIDSch)
        Dim v_DistID As String = TIMS.GetListValue(DistID)
        Dim v_Years As String = TIMS.GetListValue(Years)

        'Dim dt As DataTable
        Dim sql As String = ""
        sql = " select distinct a.SBSN" & vbCrLf
        sql &= " ,dbo.NVL(b.name,'-') NAME" & vbCrLf
        'sql &= If(iType = 1, ",dbo.FN_GET_MASK1(a.IDNO) IDNO", ",a.IDNO") & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK,a.IDNO" & vbCrLf
        sql &= " ,CONVERT(varchar, a.SBSDATE, 111) SBSdate" & vbCrLf
        sql &= " ,a.SBYears" & vbCrLf
        sql &= " ,a.SBComment" & vbCrLf
        sql &= " ,c.Name DistName" & vbCrLf
        sql &= " ,a.DistID" & vbCrLf
        sql &= " ,a.SBTERMS" & vbCrLf
        sql &= " ,a.TPlanID" & vbCrLf
        sql &= " ,kp.PlanName" & vbCrLf
        sql &= " FROM STUD_BLACKLIST a" & vbCrLf
        sql &= " JOIN KEY_Plan kp on kp.TPlanID=a.TPlanID" & vbCrLf
        sql &= " LEFT JOIN STUD_STUDENTINFO b on b.IDNO=a.IDNO" & vbCrLf
        sql &= " LEFT JOIN ID_District c on a.Distid = c.Distid " & vbCrLf
        sql &= " where a.Avail='Y'" & vbCrLf
        If v_ddlTPlanIDSch <> "" Then sql &= " and a.TPlanID ='" & v_ddlTPlanIDSch & "'" & vbCrLf
        '身分證條件
        If IDNO.Text <> "" Then sql += "and a.IDNO='" & IDNO.Text & "'" & vbCrLf

        If v_DistID <> "" Then sql += "and a.DistID ='" & v_DistID & "'" & vbCrLf

        If v_Years <> "" Then sql += "and DATEPART(YEAR, a.SBSDATE) ='" & v_Years & "'" & vbCrLf

        If Name.Text <> "" Then sql += "and b.name like '%" & Name.Text & "%'" & vbCrLf

        sql &= " order by a.DistID,SBSdate "
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        '作業顯示模式：0:其他 1:模糊顯示 2:正常顯示
        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode)
        ViewState(TIMS.gcst_rblWorkMode) = v_rblWorkMode
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "PLANNAME,DISTNAME,NAME,IDNO,SBSDATE,SBYEARS,SBCOMMENT")
        If ISEXP1 = "Y" Then
            Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, v_rblWorkMode, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)
        Else
            Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, v_rblWorkMode, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)
        End If

        Call SUtl_PanelList(1)
        tb_Sch.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count = 0 Then Return
        'If dt.Rows.Count > 0 Then End If

        tb_Sch.Visible = True
        msg.Text = ""

        dg_Sch.DataKeyField = "SBSN"
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub Dg_Sch_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim sCmdArg As String = e.CommandArgument
        Dim SBSN As String = TIMS.GetMyValue(sCmdArg, "SBSN") ', Convert.ToString(drv("SBSN")))
        If SBSN = "" Then Exit Sub

        Dim sql As String = ""
        Select Case e.CommandName
            Case "view" '檢視
                Call SUtl_PanelList(3) '檢視
                Call Show_SGDetail(SBSN, "V") 'I:新增/V:檢視/E:編輯
                lbl_title.Text = Cst_strview '"黑名單資料檢視"

            Case "edit" '修改
                Call SUtl_PanelList(2) '修改
                Call Show_SGDetail(SBSN, "E") 'I:新增/V:檢視/E:編輯
                lbl_title.Text = Cst_strUPDATE '"黑名單資料修改"

            Case "del" '刪除(非資料刪除只做註記,使用者無法看到資料)
                Dim dParms As New Hashtable From {{"SBSN", SBSN}, {"MODIFYACCT", sm.UserInfo.UserID}}
                Dim SQL_U As String = ""
                SQL_U &= " UPDATE STUD_BLACKLIST"
                SQL_U &= " SET avail='N' ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE() WHERE SBSN=@SBSN"
                DbAccess.ExecuteNonQuery(SQL_U, objconn, dParms)

                Common.MessageBox(Me, "刪除成功")

                Call SSearch1("")
        End Select
    End Sub

    Private Sub Dg_Sch_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound
        'Case ListItemType.Header, ListItemType.Footer
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtView As LinkButton = e.Item.FindControl("lbtView")
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim lbtDel As LinkButton = e.Item.FindControl("lbtDel")
                Dim labIDNO As Label = e.Item.FindControl("labIDNO")
                Dim labSBTERMS As Label = e.Item.FindControl("labSBTERMS")
                'e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                If dg_Sch.AllowPaging Then
                    '有分頁執行/無分頁情況為匯出
                    If Len(e.Item.Cells(Cst_事由位置).Text) > 19 Then '事由內容超過19字後顯示...
                        e.Item.Cells(Cst_事由位置).Text = Mid(e.Item.Cells(Cst_事由位置).Text, 1, 18) + "..."
                    End If
                End If

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SBSN", Convert.ToString(drv("SBSN")))

                lbtView.Visible = False
                lbtEdit.Visible = True
                lbtDel.Visible = True
                lbtEdit.CommandArgument = sCmdArg ' drv("SBSN")
                lbtView.CommandArgument = ""
                Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
                If Not flagS1 Then
                    If sm.UserInfo.DistID <> drv("DistID") Then '如果不是原轄區中心
                        lbtView.Visible = True
                        lbtEdit.Visible = False
                        lbtDel.Visible = False
                        lbtEdit.CommandArgument = ""
                        lbtView.CommandArgument = sCmdArg 'drv("SBSN")
                    End If
                End If
                '身分證號碼 '作業顯示模式：0:其他 1:模糊顯示 2:正常顯示 
                labIDNO.Text = Convert.ToString(If(ViewState(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip2, drv("IDNO"), drv("IDNO_MK")))
                '處分緣由
                labSBTERMS.Text = If(Convert.ToString(drv("SBTERMS")) <> "", TIMS.Get_SBTERMSName(TIMS.Get_SBTERM(), drv("SBTERMS")), "")

                lbtDel.CommandArgument = sCmdArg 'drv("SBSN")
                lbtDel.Attributes("onclick") = "return confirm('確定要刪除第" & e.Item.Cells(0).Text & "筆紀錄?');"
        End Select

    End Sub

    '單1資料顯示。
    Sub Show_SGDetail(ByVal siSBSN As String, ByVal TYPE As String)
        HidvsType.Value = "" 'I:新增/V:檢視/E:編輯
        siSBSN = TIMS.ClearSQM(siSBSN)
        If siSBSN = "" Then Exit Sub

        HidSBSN.Value = siSBSN
        HidvsType.Value = TYPE 'I:新增/V:檢視/E:編輯

        'SBSN 條件
        'TYPE 執行動作


        Dim pms1 As New Hashtable From {{"SBSN", TIMS.CINT1(siSBSN)}}
        Dim sql As String = ""
        sql &= " select c.DistID" & vbCrLf
        sql &= " ,c.Name DistName" & vbCrLf
        sql &= " ,oo.Orgname" & vbCrLf
        sql &= " ,case when ke.JobID is not null then '[' + ke.JobID + ']' + ke.JobName end TrainName" & vbCrLf
        sql &= " ,case when cc.ocid is not null then dbo.FN_GET_CLASSCNAME(cc.ClasscName,cc.CyclType) end Classname" & vbCrLf
        sql &= " ,ISNULL(cc.RID,a.RID) RID" & vbCrLf
        sql &= " ,cc.ocid,cc.TMID,a.IDNO,a.SBNum,a.SBYears,a.SBSDATE,a.SBComment,a.SBTERMS,a.TPlanID,kp.PlanName" & vbCrLf
        sql &= " from STUD_BLACKLIST a" & vbCrLf
        sql &= " JOIN KEY_Plan kp on kp.TPlanID=a.TPlanID" & vbCrLf
        sql &= " LEFT JOIN STUD_STUDENTINFO b on b.IDNO = a.IDNO" & vbCrLf
        sql &= " LEFT JOIN ID_District c on c.Distid=a.Distid" & vbCrLf
        sql &= " LEFT JOIN Class_ClassInfo cc on cc.ocid = a.ocid" & vbCrLf
        sql &= " LEFT JOIN Key_TrainType ke on ke.TMID = cc.TMID" & vbCrLf
        sql &= " LEFT JOIN Auth_Relship ar on ar.rid = a.rid" & vbCrLf
        sql &= " LEFT JOIN Org_OrgInfo oo on oo.orgid = ar.orgid" & vbCrLf
        sql &= " WHERE a.SBSN=@SBSN" & vbCrLf 'SBSN條件
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, pms1)
        If dr Is Nothing Then Exit Sub

        Select Case TYPE 'V 檢視/E 編輯/I 新增
            Case "V"
                lbl_PlanName.Text = Convert.ToString(dr("PlanName"))
                lbl_DistID.Text = Convert.ToString(dr("DistName"))
                lbl_RID.Text = Convert.ToString(dr("Orgname"))
                lbl_ClassName.Text = Convert.ToString(dr("TrainName")) & Convert.ToString(dr("Classname"))
                lbl_idno.Text = dr("idno").ToString
                lbl_No.Text = Convert.ToString(dr("SBNum"))

                Common.SetListItem(ddlSBTERMS, Convert.ToString(dr("SBTERMS")))
                Dim v_ddlSBTERMS As String = TIMS.GetListText(ddlSBTERMS)
                lbl_SBTERMS.Text = v_ddlSBTERMS
                lbl_SBYears.Text = Convert.ToString(dr("SBYears"))
                If Convert.ToString(dr("SBSDATE")) <> "" Then
                    lbl_year.Text = Year(dr("SBSDATE"))
                    lbl_month.Text = Month(dr("SBSDATE"))
                    lbl_day.Text = Day(dr("SBSDATE"))
                End If
                'lbl_SBCommect.Text = dr("SBComment").ToString
                '採換行顯示功能
                dr("SBComment") = TIMS.ClearSQM(dr("SBComment"))
                Const Cst_CutLen As Integer = 40
                lbl_SBCommect.Text = TIMS.Utl_WORDWRAP(Convert.ToString(dr("SBComment")), Cst_CutLen)

                HidSBSN.Value = siSBSN
                HidvsType.Value = "V"'V 檢視/E 編輯/I 新增
            Case "E"
                Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
                ddlTPlanID.Enabled = True
                If Not flagS1 Then ddlTPlanID.Enabled = False

                Common.SetListItem(ddlTPlanID, Convert.ToString(dr("TPlanID")))
                'ddl_DistID.SelectedValue = dr("DistID")
                Common.SetListItem(ddl_DistID, Convert.ToString(dr("DistID")))
                center.Text = Convert.ToString(dr("Orgname"))

                RIDValue.Value = Convert.ToString(dr("RID"))
                TMID1.Text = Convert.ToString(dr("TrainName"))
                OCID1.Text = Convert.ToString(dr("Classname"))
                TMIDValue1.Value = Convert.ToString(dr("TMID"))
                OCIDValue1.Value = Convert.ToString(dr("OCID"))

                txt_idno.Text = Convert.ToString(dr("IDNO"))
                txt_No.Text = Convert.ToString(dr("SBNum"))
                Common.SetListItem(ddlSBTERMS, dr("SBTERMS"))
                txt_SBSdate.Text = Common.FormatDate(Convert.ToString(dr("SBSDATE")))
                Common.SetListItem(ddl_SBYears, Convert.ToString(dr("SBYears")))

                If Convert.ToString(dr("SBYears")) <> "" Then
                End If
                txt_SBComment.Text = TIMS.ClearSQM(dr("SBComment"))

                HidSBSN.Value = siSBSN
                HidvsType.Value = "E" 'V 檢視/E 編輯/I 新增
        End Select

    End Sub

    '新增鈕
    Private Sub Btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
        Call SUtl_PanelList(2) '新增
        Call Clear_value1()

        lbl_title.Text = Cst_stradd ' "黑名單資料新增"
        HidSBSN.Value = ""
        HidvsType.Value = "I" 'I:新增/V:檢視/E:編輯

        Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        ddlTPlanID.Enabled = True
        If Not flagS1 Then ddlTPlanID.Enabled = False '新增鎖定。
        Common.SetListItem(ddlTPlanID, sm.UserInfo.TPlanID)

        '20100208 按新增時代查詢之 身分證號
        txt_idno.Text = IDNO.Text
    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="vsType"></param>
    Sub SAVE_STUD_BLACKLIST1(ByRef vsType As String)

        Select Case vsType'Convert.ToString(HidvsType.Value)'I:新增/V:檢視/E:編輯
            Case "I"
                Dim I_sql As String = ""
                I_sql &= " insert into STUD_BLACKLIST(SBSN,idno,sbsdate,sbyears,sbcomment,avail,modifyacct,modifydate,SBNum,OCID,DistID,RID,SBTERMS,TPlanID)" & vbCrLf
                I_sql &= " VALUES (@SBSN,@idno,@sbsdate,@sbyears,@sbcomment,@avail,@modifyacct,getdate() ,@SBNum,@OCID,@DistID,@RID,@SBTERMS,@TPlanID)" & vbCrLf
                Dim iCmd As New SqlCommand(I_sql, objconn)
                'STUD_BLACKLIST_SBSN_SEQ
                Dim iSBSN As Integer = DbAccess.GetNewId(objconn, "STUD_BLACKLIST_SBSN_SEQ,STUD_BLACKLIST,SBSN")
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("SBSN", SqlDbType.Int).Value = iSBSN
                    .Parameters.Add("idno", SqlDbType.VarChar).Value = txt_idno.Text
                    .Parameters.Add("sbsdate", SqlDbType.DateTime).Value = TIMS.Cdate2(txt_SBSdate.Text)
                    .Parameters.Add("sbyears", SqlDbType.VarChar).Value = ddl_SBYears.SelectedValue
                    .Parameters.Add("sbcomment", SqlDbType.NVarChar).Value = txt_SBComment.Text
                    .Parameters.Add("avail", SqlDbType.VarChar).Value = "Y"
                    .Parameters.Add("modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                    .Parameters.Add("SBNum", SqlDbType.VarChar).Value = txt_No.Text
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = If(OCIDValue1.Value <> "", OCIDValue1.Value, Convert.DBNull)
                    .Parameters.Add("DistID", SqlDbType.VarChar).Value = ddl_DistID.SelectedValue
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
                    .Parameters.Add("SBTERMS", SqlDbType.VarChar).Value = ddlSBTERMS.SelectedValue
                    .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = ddlTPlanID.SelectedValue
                    .ExecuteNonQuery()
                End With

                'txt_idno.Text = UCase(Mid(txt_idno.Text, 1, 1)) + Mid(txt_idno.Text, 2, 9)
                'sql = "insert into STUD_BLACKLIST(idno,sbsdate,sbyears,sbcomment,avail,modifyacct,modifydate,SBNum,OCID,DistID,RID)" & vbCrLf
                'sql += "values(upper('" & TIMS.ChangeIDNO(txt_idno.Text) & "'),convert(datetime, '" & txt_SBSdate.Text & "', 111)," & ddl_SBYears.SelectedValue & ",'" & txt_SBComment.Text & "','Y','" & sm.UserInfo.UserID & "',getdate(),'" & txt_No.Text & "',"
                '" & OCIDValue1.Value & "','" & ddl_DistID.SelectedValue & "','" & RIDValue.Value & "')"
                'DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "儲存成功")

                Call SSearch1("")
            Case "E"
                If Convert.ToString(HidSBSN.Value) = "" Then Exit Sub
                'txt_idno.Text = UCase(Mid(txt_idno.Text, 1, 1)) + Mid(txt_idno.Text, 2, 9)
                Dim U_sql As String = ""
                U_sql &= " update STUD_BLACKLIST" & vbCrLf
                U_sql &= " set idno =@idno,sbsdate=@sbsdate,sbyears=@sbyears,sbcomment=@sbcomment,avail =@avail" & vbCrLf
                U_sql &= " ,SBNum=@SBNum,OCID=@OCID,DistID=@DistID,RID=@RID,SBTERMS=@SBTERMS,TPlanID=@TPlanID" & vbCrLf
                U_sql &= " ,modifyacct=@modifyacct,modifydate=getdate()" & vbCrLf
                U_sql &= " where SBSN=@SBSN" & vbCrLf
                Dim uCmd As New SqlCommand(U_sql, objconn)
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("idno", SqlDbType.VarChar).Value = txt_idno.Text
                    .Parameters.Add("sbsdate", SqlDbType.DateTime).Value = TIMS.Cdate2(txt_SBSdate.Text)
                    .Parameters.Add("sbyears", SqlDbType.VarChar).Value = ddl_SBYears.SelectedValue
                    .Parameters.Add("sbcomment", SqlDbType.NVarChar).Value = txt_SBComment.Text
                    .Parameters.Add("avail", SqlDbType.VarChar).Value = "Y"
                    .Parameters.Add("modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                    .Parameters.Add("SBNum", SqlDbType.VarChar).Value = txt_No.Text
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = If(OCIDValue1.Value <> "", OCIDValue1.Value, Convert.DBNull)
                    .Parameters.Add("DistID", SqlDbType.VarChar).Value = ddl_DistID.SelectedValue
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
                    .Parameters.Add("SBTERMS", SqlDbType.VarChar).Value = ddlSBTERMS.SelectedValue
                    .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = ddlTPlanID.SelectedValue
                    .Parameters.Add("SBSN", SqlDbType.VarChar).Value = HidSBSN.Value
                    .ExecuteNonQuery()
                End With

                'DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "修改成功")

                Call SSearch1("")
        End Select
    End Sub

    '儲存鈕
    Private Sub Btn_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click
        'Dim sql As String = ""
        'If txt_idno.Text <> "" Then
        '    txt_idno.Text = UCase(Mid(txt_idno.Text, 1, 1)) & Mid(txt_idno.Text, 2, 9)
        'End If
        txt_idno.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(txt_idno.Text))
        txt_SBComment.Text = TIMS.ClearSQM(txt_SBComment.Text)
        HidvsType.Value = TIMS.ClearSQM(HidvsType.Value) 'I:新增/V:檢視/E:編輯

        Dim sErrmsg As String = ""
        Dim flagIDNO As Boolean = False '驗證異常(1:國民身分證)
        Dim flagPermit As Boolean = False '驗證異常('2:居留證)

        '1:國民身分證 -檢查
        Dim flagIdno1 As Boolean = TIMS.CheckIDNO(txt_idno.Text)
        '2:居留證 4:居留證2021 -檢查
        Dim flagPermit2 As Boolean = TIMS.CheckIDNO2(txt_idno.Text, 2)
        Dim flagPermit4 As Boolean = TIMS.CheckIDNO2(txt_idno.Text, 4)
        If Not flagIdno1 AndAlso Not flagPermit2 AndAlso Not flagPermit4 Then
            sErrmsg &= "身分證號碼或居留證號碼 輸入錯誤!" & vbCrLf
        End If

        Select Case ddlSBTERMS.SelectedValue
            Case "57"
                If ddl_SBYears.SelectedValue <> "1" Then
                    sErrmsg &= "五十七 處分確定日起一年內不予補助!!(請選擇1年)" & vbCrLf
                End If
            Case "58"
                If ddl_SBYears.SelectedValue <> "2" Then
                    sErrmsg &= "五十八 處分確定日起二年內不予補助!!(請選擇2年)" & vbCrLf
                End If
        End Select

        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        Dim htSS As New Hashtable From {
            {"IDNO", txt_idno.Text},
            {"SBSdate", txt_SBSdate.Text},
            {"Type", Convert.ToString(HidvsType.Value)}, 'I:新增/V:檢視/E:編輯
            {"SBSN", Convert.ToString(HidSBSN.Value)}
        }
        '檢查資料
        sErrmsg = Check_Blacklist(htSS, objconn)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        'HidvsType.Value I:新增/V:檢視/E:編輯
        SAVE_STUD_BLACKLIST1(HidvsType.Value)
    End Sub

    ''' <summary>
    ''' 檢查身分證號-處分起日資料是否已存在
    ''' </summary>
    ''' <param name="htSS"></param>
    ''' <param name="objconn"></param>
    ''' <returns></returns>
    Public Shared Function Check_Blacklist(ByRef htSS As Hashtable, ByVal objconn As SqlConnection) As String
        Dim rst As String = ""
        Dim IDNO As String = TIMS.GetMyValue2(htSS, "IDNO")
        Dim TPlanID As String = TIMS.GetMyValue2(htSS, "TPlanID")
        Dim SBSdate As String = TIMS.GetMyValue2(htSS, "SBSdate") 'yyyy/MM/dd
        Dim Type As String = TIMS.GetMyValue2(htSS, "Type") 'yyyy/MM/dd
        Dim SBSN As String = TIMS.GetMyValue2(htSS, "SBSN") 'yyyy/MM/dd

        'txt_idno.Text = TIMS.ClearSQM(txt_idno.Text)
        'txt_idno.Text = TIMS.ChangeIDNO(txt_idno.Text)
        IDNO = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO))
        Dim parms As New Hashtable From {{"IDNO", IDNO}, {"TPlanID", TPlanID}}
        Dim sql As String = ""
        sql &= " select a.IDNO ,CONVERT(varchar, a.SBSdate, 111) SBSdate" & vbCrLf 'yyyy/MM/dd
        sql &= " from STUD_BLACKLIST a" & vbCrLf
        sql &= " where a.Avail='Y' and a.IDNO=@IDNO and a.TPlanID=@TPlanID" & vbCrLf
        Select Case Type
            Case "E"
                If SBSN <> "" Then
                    sql &= " and a.SBSN !=@SBSN" & vbCrLf
                    parms.Add("SBSN", SBSN)
                End If
        End Select

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        For i As Int16 = 0 To dt.Rows.Count - 1
            If dt.Rows(i)("SBSdate") = SBSdate Then
                rst = "此學員 " + SBSdate + " 處分資料已存在!"
                Exit For
                'Return msg
                'Exit Function
            End If
        Next
        Return rst
    End Function

    '清除編輯值。
    Sub Clear_value1()
        ddlTPlanID.Enabled = False '新增/修改鎖定(預設)。
        'ddlTPlanID.SelectedIndex = -1
        Common.SetListItem(ddlTPlanID, sm.UserInfo.TPlanID)

        txt_idno.Text = ""

        txt_No.Text = ""
        ddlSBTERMS.SelectedIndex = -1
        txt_SBSdate.Text = ""
        ddl_SBYears.SelectedIndex = 0
        txt_SBComment.Text = ""
        center.Text = ""
        RIDValue.Value = ""
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""

        HidSBSN.Value = ""
        'HidvsType.Value = ""
    End Sub

    '儲存離開鈕
    Private Sub Btn_lev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev.Click
        Call SUtl_PanelList(1) '搜尋
        Call Clear_value1()
        lbl_title.Text = Cst_strsearch ' "黑名單資料查詢"
    End Sub

    '檢視離開鈕
    Private Sub Btn_lev2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev2.Click
        Call SUtl_PanelList(1) '搜尋
        Call Clear_value1()
        lbl_title.Text = Cst_strsearch ' "黑名單資料查詢"
    End Sub

    '匯出
    Sub SExport1()
        'Const Cst_功能欄位 As Integer = 9
        'Const Cst_xlsFileName As String = "學員處分資料匯出.xls"
        Dim oDataGrid1 As DataGrid = dg_Sch

        oDataGrid1.AllowPaging = False
        oDataGrid1.EnableViewState = False  '把ViewState給關了

        Call SSearch1("Y")

        Dim sFileName1 As String = "學員處分資料匯出"

        Dim strSTYLE As String = ""
        ''套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        oDataGrid1.AllowPaging = False
        oDataGrid1.Columns(Cst_功能欄位).Visible = False
        oDataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim parmsExp As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)},
            {"FileName", sFileName1},
            {"strSTYLE", strSTYLE},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)

        oDataGrid1.AllowPaging = True
        oDataGrid1.Columns(Cst_功能欄位).Visible = True
        'Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '匯出
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        HidSBSN.Value = ""

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Call SExport1()
    End Sub

    Private Sub BtnGETvalue2_Click(sender As Object, e As System.EventArgs) Handles BtnGETvalue2.Click
        'BtnGETvalue2
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub
    Protected Sub DdlTPlanID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTPlanID.SelectedIndexChanged
        Dim v_ddlTPlanID As String = TIMS.GetListValue(ddlTPlanID)
        If v_ddlTPlanID = "" Then Return

        UTL_SBTERMS(ddlSBTERMS, ddl_SBYears, v_ddlTPlanID)
    End Sub

    Public Shared Sub UTL_SBTERMS(ByRef ddlSBTERMS As DropDownList, ByRef ddl_SBYears As DropDownList, ByRef v_TPlanID As String)
        ddlSBTERMS.Items.Clear()
        If v_TPlanID = "" Then Return

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(v_TPlanID) > -1 Then
            Select Case v_TPlanID'Convert.ToString(sm.UserInfo.TPlanID)
                Case "28"
                    ddlSBTERMS = TIMS.Get_BTERMS(ddlSBTERMS, 2)
                Case "54"
                    ddlSBTERMS = TIMS.Get_BTERMS(ddlSBTERMS, 254)
            End Select
            ddl_SBYears = Get_SBYears(ddl_SBYears, 2)
        Else
            '其它或未定義
            ddlSBTERMS = TIMS.Get_BTERMS(ddlSBTERMS, 99)
            ddl_SBYears = Get_SBYears(ddl_SBYears, 2)
        End If
        'If TIMS.Cst_TPlanID68.IndexOf(v_TPlanID) > -1 Then
        '    ddlSBTERMS = TIMS.Get_BTERMS(ddlSBTERMS, 42)
        '    ddl_SBYears = Get_SBYears(ddl_SBYears, 42)
        'End If
    End Sub

    Public Shared Function Get_SBYears(ByVal obj As ListControl, ByVal iType As Integer) As ListControl
        obj.Items.Clear()
        Select Case iType
            Case 2 '學員處分功能(SD_16_001)  '28:產業人才投資方案
                obj.Items.Add(New ListItem("請選擇", "-1"))
                For i As Integer = 0 To 3
                    obj.Items.Add(New ListItem(CStr(i), CStr(i)))
                Next
            Case 42 '學員處分功能(SD_16_001)  '68:照顧服務員自訓自用訓練計畫
                obj.Items.Add(New ListItem("請選擇", "-1"))
                For i As Integer = 0 To 1
                    obj.Items.Add(New ListItem(CStr(i), CStr(i)))
                Next
        End Select
        Return obj
    End Function

End Class