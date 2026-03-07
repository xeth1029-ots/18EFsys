Partial Class TC_04_002
    Inherits AuthBasePage

    '若是產學訓計畫，則跳到 TC_04_002.aspx 執行審核動作
    Dim gflag_ERROR_STOP As Boolean = False

    'DataGrid2
    Const Cst_Index As Integer = 0
    Const Cst_PlanID As Integer = 1
    Const Cst_ComIDNO As Integer = 9
    Const Cst_SeqNo As Integer = 2
    Const Cst_ParentName As Integer = 4 '訓練機構
    Const Cst_CyclType As Integer = 5 '班別名稱/期別

    Const Col_Phone As Integer = 13
    Const Col_資格初審 As Integer = 14
    Const Col_審核 As Integer = 15
    Const Col_原因 As Integer = 16

    'Dim SqlCmd As String
    'Dim dsA As SqlDataAdapter
    'Dim sqlAdapter As SqlDataAdapter
    'Dim sqlTable As DataTable
    'Dim Sqlstr As String
    'Dim DistID As String
    'Dim PlanID As String
    'Dim UserID As String
    'Const cst_TC04002Addaspx As String = "../04/TC_04_002_Add.aspx" 'OLD
    Const cst_TC04002Add2aspx As String = "../04/TC_04_002_Add2.aspx" 'NEW
    Const cst_TC03006Addaspx As String = "../03/TC_03_006.aspx"

    Dim Auth_Relship As DataTable
    Const Cst_EmptySelValue As String = "==請選擇=="
    Const cst_errmsg4 As String = "使用者登入計畫有誤，不提供儲存!!"

    '產業人才投資方案，審核計畫專用
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            LabTMID.Text = "訓練業別"
        End If

        'DistID = sm.UserInfo.DistID
        'PlanID = sm.UserInfo.PlanID
        'UserID = sm.UserInfo.UserID
        Pagecontroler1.PageDataGrid = dgPlan
        PageControler2.PageDataGrid = DataGrid2

        '依登入者RID、TPlanID、Years、PlanID  取得 Auth_Relship.RID,OrgName
        Auth_Relship = TIMS.sUtl_GetAuthRelship(Me, objconn)

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    FunDr = FunDrArray(0)
        '    If FunDr("Sech") = 1 Then
        '        btnQuery.Enabled = True
        '    Else
        '        btnQuery.Enabled = False
        '        TIMS.Tooltip(btnQuery, "沒有搜尋權限")
        '    End If
        '    If FunDr("Adds") = 1 Then
        '        bntAdd.Enabled = True
        '        'If sm.UserInfo.Years = 2006 Then Button2.Visible = True Else Button2.Visible = False
        '    Else
        '        bntAdd.Enabled = False
        '        TIMS.Tooltip(btnQuery, "沒有新增權限")
        '    End If

        'End If

        If Not IsPostBack Then
            '有不區分
            OrgKind2 = TIMS.Get_RblSearchPlan(Me, OrgKind2)
            Common.SetListItem(OrgKind2, "A")

            '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)
            If tr_AppStage_TP28.Visible Then
                AppStage2 = TIMS.Get_AppStage2_NotCase(AppStage2)
                Common.SetListItem(AppStage2, "")
            End If

            'DistrictList = TIMS.Get_DistID(DistrictList)
            Call TIMS.Get_DISTCBL(DistrictList, objconn)
            TRA.Visible = False
            '取得訓練計畫
            TPlanid.Value = sm.UserInfo.TPlanID 'DbAccess.ExecuteScalar(Sqlstr, objconn)
            '(加強操作便利性)2005/4/1-melody
            RIDValue.Value = sm.UserInfo.RID
            'Sqlstr = "select orgname from Auth_Relship a join Org_orginfo b on  a.orgid=b.orgid where a.RID='" & sm.UserInfo.RID & "'"
            center.Text = sm.UserInfo.OrgName 'DbAccess.ExecuteScalar(Sqlstr, objconn)

            DataGridTable1.Visible = False
            DataGridTable2.Visible = False
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            Button1.Attributes("onclick") = "return SavaData();"

            trDistType.Visible = False
            If sm.UserInfo.DistID = "000" Then
                '署
                trDistType.Visible = True
                DistType.Enabled = True
                DistrictList.Enabled = True
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                '選擇全部轄區
                DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"
                'DistType.Attributes("onclick") = "SetDistType('DistType','DistrictList','center','Org');"
            Else
                '委訓／分署
                'DistType: 搜尋型態: 0:依轄區 1:依訓練機構
                Common.SetListItem(DistType, "1") '0/1
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistType.Enabled = False
                DistrictList.Enabled = False
                TIMS.Tooltip(DistType, TIMS.cst_ErrorMsg16, True)
                TIMS.Tooltip(DistrictList, TIMS.cst_ErrorMsg16, True)
            End If
        End If


        '帶入查詢參數
        If Not IsPostBack Then
            If Not Session("search") Is Nothing Then
                Dim MyValue As String = ""

                MyValue = TIMS.GetMyValue(Session("search"), "prg")
                If MyValue = "TC_04_002" Then
                    'DistType: 搜尋型態: 0:依轄區 1:依訓練機構
                    Common.SetListItem(DistType, TIMS.GetMyValue(Session("search"), "DistType"))
                    Common.SetListItem(DistrictList, TIMS.GetMyValue(Session("search"), "DistrictList"))
                    DistHidden.Value = TIMS.GetMyValue(Session("search"), "DistHidden")
                    RIDValue.Value = TIMS.GetMyValue(Session("search"), "RIDValue")

                    center.Text = TIMS.GetMyValue(Session("search"), "center")
                    TB_career_id.Text = TIMS.GetMyValue(Session("search"), "TB_career_id")
                    TPlanid.Value = TIMS.GetMyValue(Session("search"), "TPlanid")
                    trainValue.Value = TIMS.GetMyValue(Session("search"), "trainValue")
                    jobValue.Value = TIMS.GetMyValue(Session("search"), "jobValue")

                    txtCJOB_NAME.Text = TIMS.GetMyValue(Session("search"), "txtCJOB_NAME")
                    cjobValue.Value = TIMS.GetMyValue(Session("search"), "cjobValue")

                    ClassName.Text = TIMS.GetMyValue(Session("search"), "ClassName")
                    CyclType.Text = TIMS.GetMyValue(Session("search"), "CyclType")
                    UNIT_SDATE.Text = TIMS.GetMyValue(Session("search"), "UNIT_SDATE")
                    UNIT_EDATE.Text = TIMS.GetMyValue(Session("search"), "UNIT_EDATE")
                    start_date.Text = TIMS.GetMyValue(Session("search"), "start_date")
                    end_date.Text = TIMS.GetMyValue(Session("search"), "end_date")
                    Common.SetListItem(OrgKind2, TIMS.GetMyValue(Session("search"), "OrgKind2"))
                    Common.SetListItem(AppStage2, TIMS.GetMyValue(Session("search"), "AppStage2"))
                    Common.SetListItem(PlanMode, TIMS.GetMyValue(Session("search"), "PlanMode"))
                    Common.SetListItem(AdvanceMode, TIMS.GetMyValue(Session("search"), "AdvanceMode"))
                    'btnQuery_Click(sender, e)
                    Call SSearch1()
                End If

                Session("search") = Nothing
            End If
        End If


    End Sub

    Sub SSearch1()
        'Const Cst_F = "F" '分署(中心):初審 'Const Cst_S = "S" '署(局):複審
        UNIT_SDATE.Text = TIMS.ClearSQM(UNIT_SDATE.Text)
        UNIT_EDATE.Text = TIMS.ClearSQM(UNIT_EDATE.Text)

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)

        UNIT_SDATE.Text = TIMS.Cdate3(UNIT_SDATE.Text)
        UNIT_EDATE.Text = TIMS.Cdate3(UNIT_EDATE.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(UNIT_SDATE.Text, UNIT_EDATE.Text) Then
            Dim T_DATE1 As String = UNIT_SDATE.Text
            UNIT_SDATE.Text = UNIT_EDATE.Text
            UNIT_EDATE.Text = T_DATE1
        End If
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(start_date.Text, end_date.Text) Then
            Dim T_DATE1 As String = start_date.Text
            start_date.Text = end_date.Text
            end_date.Text = T_DATE1
        End If

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        'Dim vRelship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        If sm.UserInfo.OrgLevel <= 1 Then
            '署(局):0／分署(中心):1
            dgPlan.Columns(Col_Phone).Visible = False
            dgPlan.Columns(Col_資格初審).Visible = True
            dgPlan.Columns(Col_審核).Visible = True
            dgPlan.Columns(Col_原因).Visible = True
            bntAdd.Visible = True
        Else
            '委訓
            dgPlan.Columns(Col_Phone).Visible = True
            dgPlan.Columns(Col_資格初審).Visible = False
            dgPlan.Columns(Col_審核).Visible = False
            dgPlan.Columns(Col_原因).Visible = False
            bntAdd.Visible = False
        End If


        Dim v_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        Dim v_AdvanceMode As String = TIMS.GetListValue(AdvanceMode)
        Dim v_PlanMode As String = TIMS.GetListValue(PlanMode)
        '依申請階段
        Dim v_AppStage2 As String = "" 'TIMS.GetListValue(AppStage2)
        If tr_AppStage_TP28.Visible Then v_AppStage2 = TIMS.GetListValue(AppStage2)

        Dim sql As String = ""
        sql &= " Select P1.PlanYear,P1.PlanID,P1.ComIDNO,P1.SeqNo,P1.ClassName,P1.CyclType"
        sql &= " ,dbo.FN_GET_CLASSCNAME(P1.CLASSNAME ,P1.CYCLTYPE) CLASSNAME2"
        ',dbo.FN_GET_CLASSCNAME(P1.CLASSNAME,P1.CYCLTYPE) 
        '+(CASE WHEN P1.RESULTBUTTON IN ('Y','R') THEN '(未送出)' ELSE '' END) CLASSCNAME
        sql &= " ,P1.AppliedDate,P1.STdate,P1.FDdate,P1.AppliedResult"
        sql &= " ,P1.RESULTBUTTON" & vbCrLf 'NULL:已送出不可修改 Y:還原可修改 R:退件修改
        sql &= " ,case when P1.TransFlag='Y' then '是' else '否' end TransFlag"
        sql &= " ,ISNULL(P1.TransFlag,'N') TransFlag2"
        sql &= " ,P1.TNum,P1.THours,P1.DefGovCost,P1.DefStdCost,P1.ProcID,P1.PointYN"
        sql &= " ,Case when ISNULL(P1.PointYN,'N') ='Y' then '學分班' else '非學分班' end Point"
        sql &= " ,P1.SciPlaceID,P1.TechPlaceID,P1.TPlanID,P1.TMID,P1.CapDegree,P1.ClassCate,A1.relship,O1.OrgName,O2.Address"
        'sql += " ,O2.ContactName" 'X 'sql += " ,O2.Phone" 'X
        'sql += " ,ISNULL(P1.ContactName,'<FONT COLOR=''RED''>無資料</FONT>') AS ContactName" '請使用班級計畫時的聯絡人
        'sql += " ,ISNULL(P1.ContactPhone,'<FONT COLOR=''RED''>無資料</FONT>') AS Phone" '請使用班級計畫時的聯絡人電話
        sql &= " ,P1.ContactName AS ContactName" '請使用班級計畫時的聯絡人
        sql &= " ,P1.ContactPhone AS Phone" '請使用班級計畫時的聯絡人電話
        sql &= " ,pvr.PVID ,pvr.FirResult ,pvr.SecResult" & vbCrLf
        sql &= " ,Case when pvr.SecResult='Y' then '通過' else pvrc1.VerReason end Reason_all" & vbCrLf
        sql &= " FROM dbo.PLAN_PLANINFO P1" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN P2 ON P1.PlanID=P2.PlanID" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP A1 ON P1.RID=A1.RID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO O1 ON A1.OrgID=O1.OrgID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGPLANINFO O2 ON A1.RSID=O2.RSID" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_VERREPORT pvr ON P1.PlanID = pvr.PlanID AND  P1.ComIDNO = pvr.ComIDNO AND P1.SeqNO = pvr.SeqNo" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_VERRECORD pvrc1 ON pvrc1.PlanID = pvr.PlanID AND pvrc1.ComIDNO = pvr.ComIDNO AND pvrc1.SeqNO = pvr.SeqNo AND (pvrc1.VerSeqNo = 1)" & vbCrLf
        sql &= " WHERE P1.IsApprPaper='Y'" & vbCrLf '正式
        'v_PlanMode  '(S/Y/R)
        If Not v_PlanMode = "R" Then sql &= " AND pvr.IsApprPaper='Y'" & vbCrLf '正式 

        sql &= " AND P2.PlanKind=2" & vbCrLf '計畫種類:1.自辦／2.委外
        '依登入年度
        sql &= " AND P2.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        '依登入計畫
        sql &= " AND P2.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        'DistType: 搜尋型態: 0:依轄區 1:依訓練機構
        Dim v_DistType As String = TIMS.GetListValue(DistType)
        Select Case v_DistType 'DistType.SelectedValue
            Case "0"
                '依轄區
                Select Case sm.UserInfo.LID
                    Case 0 '署-非轄區
                        If RIDValue.Value.Length = 1 Then
                            Dim sDISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                            If sDISTID <> "000" Then sql &= " AND P2.DistID='" & sDISTID & "'"
                            If sDISTID = "" Then sql &= " AND 1!=1"
                        ElseIf RIDValue.Value.Length > 1 Then
                            sql &= " AND P1.RID ='" & RIDValue.Value & "'"
                        End If
                    Case 1 '轄區
                        If RIDValue.Value.Length > 1 AndAlso RIDValue.Value <> sm.UserInfo.RID Then
                            sql &= " AND P1.RID='" & RIDValue.Value & "'"
                        End If
                        sql &= " AND P2.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                        sql &= " AND P2.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
                    Case 2 '委訓限定
                        sql &= " AND P1.RID ='" & sm.UserInfo.RID & "'"
                        sql &= " AND P2.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                        sql &= " AND P2.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
                End Select
                'Dim sDistValue As String = Get_DistIDValue()
                Dim sDistValue As String = TIMS.GetChkBoxListValue(DistrictList)
                If sDistValue <> "" Then
                    sDistValue = TIMS.CombiSQM2IN(sDistValue)
                    sql &= " AND P2.DistID IN (" & sDistValue & ")" & vbCrLf
                End If

            Case Else '1:依訓練機構
                '依訓練機構
                Select Case sm.UserInfo.LID
                    Case 0 '署-非轄區
                        If RIDValue.Value.Length = 1 Then
                            Dim sDISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                            If sDISTID <> "000" Then sql &= " AND P2.DistID ='" & sDISTID & "'"
                            If sDISTID = "" Then sql &= " AND 1!=1"
                        ElseIf RIDValue.Value.Length > 1 Then
                            sql &= " AND P1.RID ='" & RIDValue.Value & "'"
                        End If
                    Case 1 '轄區
                        If RIDValue.Value.Length > 0 AndAlso RIDValue.Value <> sm.UserInfo.RID Then
                            sql &= " AND P1.RID ='" & RIDValue.Value & "'"
                        End If
                        sql &= " AND P2.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                        sql &= " AND P2.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
                    Case 2 '委訓限定
                        sql &= " AND P1.RID ='" & sm.UserInfo.RID & "'"
                        sql &= " AND P2.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                        sql &= " AND P2.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
                End Select
        End Select

        DataGridTable1.Visible = True
        DataGridTable2.Visible = False
        '(S/Y/R)
        Select Case v_PlanMode'PlanMode.SelectedValue
            Case "S" '審核中的
                sql &= " AND pvr.SecResult IS NULL" & vbCrLf
                '產投不判斷 P1.AppliedResult 依 pvr.SecResult 為準
                'sql += " AND P1.AppliedResult IS NULL" & vbCrLf " AND P1.ResultButton IS NULL" & vbCrLf 'NULL:已送出不可修改 Y:還原可修改
            Case "Y" '已通過
                sql &= " AND pvr.SecResult='Y'" & vbCrLf
                Select Case v_AdvanceMode'AdvanceMode.SelectedValue
                    Case "S" '查詢審核狀態
                    Case "C" '取消審核
                        DataGridTable1.Visible = False
                        DataGridTable2.Visible = True
                End Select
            Case "R" '退件修正
                sql &= " AND pvr.SecResult IN ('R','N')" & vbCrLf
        End Select

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'Me.LabTMID.Text = "訓練業別"
            If jobValue.Value <> "" Then
                sql &= " AND (P1.TMID = " & jobValue.Value & vbCrLf
                sql &= " OR P1.TMID IN ( select TMID from Key_TrainType where parent IN (" & vbCrLf '職類別
                sql &= " select TMID from Key_TrainType where parent IN (" & vbCrLf '業別
                sql &= " select TMID from Key_TrainType where busid ='G')" & vbCrLf '產業人才投資方案類
                sql &= " AND TMID=" & jobValue.Value & " )))" & vbCrLf
            End If
        Else
            If trainValue.Value <> "" Then sql &= " AND P1.TMID = " & trainValue.Value & vbCrLf
        End If

        '通俗職類
        If txtCJOB_NAME.Text <> "" Then sql &= " AND P1.CJOB_UNKEY = " & cjobValue.Value & "" & vbCrLf

        If UNIT_SDATE.Text <> "" Then sql = sql & " AND P1.AppliedDate >= " & TIMS.To_date(UNIT_SDATE.Text) & vbCrLf

        If UNIT_EDATE.Text <> "" Then sql = sql & " AND P1.AppliedDate <= " & TIMS.To_date(UNIT_EDATE.Text) & vbCrLf

        If start_date.Text <> "" Then sql &= " AND P1.STDate >= " & TIMS.To_date(start_date.Text) & vbCrLf

        If end_date.Text <> "" Then sql &= " AND P1.STDate <= " & TIMS.To_date(end_date.Text) & vbCrLf

        If ClassName.Text <> "" Then sql &= " AND P1.ClassName like '%" & ClassName.Text & "%'"

        If CyclType.Text <> "" Then
            If CyclType.Text.Length < 2 Then CyclType.Text = "0" & CInt(Val(CyclType.Text))
            sql &= " AND P1.CyclType='" & CyclType.Text & "'"
        End If

        Select Case v_OrgKind2'OrgKind2.SelectedValue
            Case "G", "W" 'sql &= " AND O1.OrgKind2='" & OrgKind2.SelectedValue & "'"
                sql &= " AND O1.OrgKind2='" & v_OrgKind2 & "'"
        End Select
        '依申請階段
        If v_AppStage2 <> "" Then sql &= " AND P1.AppStage= '" & v_AppStage2 & "'" & vbCrLf

        '檢送資料-未檢送 未檢送資料
        If CB_DataNotSent_SCH.Checked Then
            sql &= " AND P1.DataNotSent='Y'" & vbCrLf
        Else
            sql &= " AND P1.DataNotSent IS NULL" & vbCrLf
        End If

        'Button1.Visible = False
        If DataGridTable1.Visible Then
            '查詢審核狀態
            sql &= " ORDER BY O1.OrgName,P1.STDate,P1.ClassName"

            Dim sqlAdapter As New SqlDataAdapter(sql, objconn)
            Dim sqlTable As New DataTable
            sqlAdapter.Fill(sqlTable)

            dgPlan.AllowPaging = False
            Pagecontroler1.Visible = False

            dgPlan.DataSource = sqlTable
            dgPlan.DataBind()
            'PageControler1.SqlDataCreate(sql)
            'PageControler1.Visible = True
        Else
            'Button1.Visible = True
            TIMS.Tooltip(Button1, "儲存 取消審核使用")
            '取消審核
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
            PageControler2.Visible = False
            Button1.Enabled = False
            If dt.Rows.Count > 0 Then
                PageControler2.Visible = True
                Button1.Enabled = True

                'PageControler2.SqlDataCreate(sql, "OrgName,STDate,ClassName")
                PageControler2.PageDataTable = dt
                PageControler2.Sort = "OrgName,STDate,ClassName"
                PageControler2.ControlerLoad()
            Else
                Button1.Enabled = False
                TIMS.Tooltip(Button1, "查無資料 不提供儲存功能")

                dt = DbAccess.GetDataTable(sql, objconn)
                DataGrid2.DataSource = dt
                DataGrid2.DataBind()
            End If

        End If
        'TIMS.CloseDbConn(objconn)
    End Sub

    '(查詢)若有改正審核搜尋條件，請一併更正檢查 主頁訊息功能。
    Private Sub BtnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call SSearch1()
    End Sub

    ''' <summary>
    ''' 審核儲存鈕(Plan_VerRecord)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BntAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntAdd.Click
        Select Case TIMS.GetVerType(sm.UserInfo.LID)
            Case "F", "S"
            Case Else
                Common.MessageBox(Me, "該登入者無審核權限!!")
                Exit Sub
        End Select

        'PlanMode:S:審核中/Y:已通過/R:退件修正(含不通過的)
        Dim v_PlanMode As String = TIMS.GetListValue(PlanMode)
        Dim V_Errmsg As String = ""
        Dim blnSaveDataOk1 As Boolean = False
        'blnSaveDataOk1 = False
        For Each myitem As DataGridItem In dgPlan.Items
            Dim objApply1 As DropDownList = myitem.FindControl("AppliedResult1")
            Dim objReason As HtmlControls.HtmlTextArea = myitem.FindControl("Reason")
            'Dim vAppliedResult As String = TIMS.ClearSQM(objApply1.SelectedValue) 'Y/N/R
            Dim v_AppliedResult As String = TIMS.GetListValue(objApply1) '.SelectedValue) 'Y/N/R 'obj:AppliedResult1
            Dim KPlanID As String = "" & TIMS.ClearSQM(myitem.Cells(Cst_PlanID).Text)
            Dim KComIDNO As String = "" & TIMS.ClearSQM(myitem.Cells(Cst_ComIDNO).Text)
            Dim KSeqNo As String = "" & TIMS.ClearSQM(myitem.Cells(Cst_SeqNo).Text)
            'If objApply.SelectedIndex <> 0 Or objApply1.SelectedIndex <> 0 Then
            If KPlanID = "" Then blnSaveDataOk1 = False
            If KComIDNO = "" Then blnSaveDataOk1 = False
            If KSeqNo = "" Then blnSaveDataOk1 = False
            If KPlanID = "" Then Exit Sub
            If KComIDNO = "" Then Exit Sub
            If KSeqNo = "" Then Exit Sub

            If objApply1.Enabled Then
                If objApply1.SelectedIndex <> 0 Then
                    If v_AppliedResult = "Y" Then
                        '檢核計畫與班級轉入
                        Dim hPCS As New Hashtable From {
                            {"PlanID", KPlanID},
                            {"ComIDNO", KComIDNO},
                            {"SeqNO", KSeqNo},
                            {"AppliedResult", v_AppliedResult}
                        }
                        Dim flag_SaveCC As Boolean = False 'true:可執行轉入/false:不可執行
                        flag_SaveCC = xChk_CC1(hPCS, V_Errmsg)
                        '審核有誤不執行轉入
                        If Not flag_SaveCC Then
                            If V_Errmsg = "" Then V_Errmsg = "資料中 審核有誤不執行轉入!"
                            Exit For
                        End If
                        'If Not flag_SaveCC Then Exit For

                        Dim drPP As DataRow = TIMS.GetPPInfo(KPlanID, KComIDNO, KSeqNo, objconn)
                        '檢核報名日期 (若OK 轉出OUT SEnterDate/FEnterDate)
                        Dim vSTDate As String = TIMS.Cdate3(drPP("STDate"))
                        Dim vSEnterDate As String = "" 'TIMS.GetMyValue2(htCC, "SEnterDate")
                        Dim vFEnterDate As String = "" 'TIMS.GetMyValue2(htCC, "FEnterDate") 'Dim flag_chkSEnDate As Boolean = False 'false:異常
                        Call TIMS.ChangeSEnterDate(vSTDate, vSEnterDate, vFEnterDate)
                        Dim flag_chkSEnDate As Boolean = If(vSEnterDate = "" OrElse vFEnterDate = "", False, True)  'false:異常
                        If Not flag_chkSEnDate Then
                            '報名時間有誤不執行轉入 'V_Errmsg = "資料中開訓時間計算報名時間有誤不執行轉入!"
                            V_Errmsg = "資料中開訓時間計算報名時間有誤 不可執行審核作業!"
                            Exit For
                            'Return rst_errmsg '"審核有誤不執行轉入!" 'Exit Sub
                        End If

                        'PlanMode:S:審核中/Y:已通過/R:退件修正(含不通過的)
                        If v_PlanMode = "Y" OrElse v_AppliedResult = "Y" Then
                            '小於、等於 開訓前三天 -不可報名
                            Dim flag_chkSEnDate3 As Boolean = TIMS.ChkEnterDayS3(vSTDate)
                            If Not flag_chkSEnDate3 Then
                                V_Errmsg = "班級審核日距離開訓日為3日(含)內，不可執行審核作業!"
                                Exit For
                            End If
                        End If
                    End If
                    blnSaveDataOk1 = True
                    Exit For
                End If
            End If
        Next
        If Not blnSaveDataOk1 Then
            Common.MessageBox(Me, "查無儲存資料，請重新確認!!")
            Exit Sub
        End If
        If V_Errmsg <> "" Then
            Common.MessageBox(Me, V_Errmsg)
            Exit Sub
        End If

        'PlanMode:S:審核中/Y:已通過/R:退件修正(含不通過的)
        'Dim v_PlanMode As String = TIMS.GetListValue(PlanMode)
        'Dim strErrmsg As String = ""
        Dim rowi As Integer = 0
        gflag_ERROR_STOP = False
        'rowi = 0
        Try
            'Dim objReason As HtmlControls.HtmlTextArea
            'Dim dtVerRecord As DataTable
            'Dim daVerRecord As New SqlDataAdapter
            'Dim sql1 As String = "Select * From Plan_VerRecord where 1<>1"
            'Dim dr As DataRow
            'dtVerRecord = DbAccess.GetDataTable(sql1, daVerRecord, objconn)
            For Each myitem As DataGridItem In dgPlan.Items
                rowi += 1
                'objApply = myitem.FindControl("AppliedResult")
                Dim objApply1 As DropDownList = myitem.FindControl("AppliedResult1")
                Dim objReason As HtmlControls.HtmlTextArea = myitem.FindControl("Reason")

                If objApply1.Enabled Then
                    If objApply1.SelectedIndex <> 0 Then
                        'Dim vAppliedResult As String = TIMS.ClearSQM(objApply1.SelectedValue) 'Y/N/R
                        Dim v_AppliedResult As String = TIMS.GetListValue(objApply1) '.SelectedValue) 'Y/N/R 'obj:AppliedResult1
                        Dim KPlanID As String = "" & TIMS.ClearSQM(myitem.Cells(Cst_PlanID).Text)
                        Dim KComIDNO As String = "" & TIMS.ClearSQM(myitem.Cells(Cst_ComIDNO).Text)
                        Dim KSeqNo As String = "" & TIMS.ClearSQM(myitem.Cells(Cst_SeqNo).Text)
                        If KPlanID = "" Then Exit Sub
                        If KComIDNO = "" Then Exit Sub
                        If KSeqNo = "" Then Exit Sub

                        'Dim drPP As DataRow = TIMS.GetPPInfo(KPlanID, KComIDNO, KSeqNo, objconn)
                        ''檢核報名日期 (若OK 轉出OUT SEnterDate/FEnterDate)
                        'Dim vSTDate As String = TIMS.cdate3(drPP("STDate"))
                        ''小於、等於 開訓前三天 -不可報名
                        'Dim flag_chkSEnDate3 As Boolean = TIMS.ChkEnterDayS3(vSTDate)
                        'If Not flag_chkSEnDate3 Then
                        '    gflag_ERROR_STOP = True
                        '    '報名時間有誤不執行轉入
                        '    strErrmsg = "班級審核日距離開訓日為3日(含)內，不可執行審核作業!"
                        '    Common.MessageBox(Me, strErrmsg)
                        '    Exit For
                        '    'Return rst_errmsg '"審核有誤不執行轉入!" 'Exit Sub
                        'End If

                        If v_AppliedResult = "Y" Then
                            TIMS.Plan_VerRecord_Update(KPlanID, KComIDNO, KSeqNo, objconn)
                        Else
                            'AppliedResult <> "Y" 
                            Dim UserID As String = sm.UserInfo.UserID
                            TIMS.Plan_VerRecord_Update(KPlanID, KComIDNO, KSeqNo, UserID, TIMS.GetVerType(sm.UserInfo.LID.ToString), "1", objReason.Value, objconn) '批次處理，錯誤原因刪除，新增於 dtVerRecord
                        End If
                        'AppliedResult: Y/N/R
                        Call TIMS.PLAN_VERREPROT_UPDATE(Me, KPlanID, KComIDNO, KSeqNo, v_AppliedResult, objconn)

                        '執行班級轉入-----------------start
                        If v_AppliedResult = "Y" Then
                            Dim hPCS As New Hashtable 'hPCS.Clear()
                            hPCS.Add("PlanID", KPlanID)
                            hPCS.Add("ComIDNO", KComIDNO)
                            hPCS.Add("SeqNO", KSeqNo)
                            hPCS.Add("AppliedResult", v_AppliedResult)
                            hPCS.Add("PlanMode", v_PlanMode)
                            Call SaveEnterCCInfoDr(hPCS)
                        End If

                    End If
                End If
            Next
            'DbAccess.UpdateDataTable(dtVerRecord, daVerRecord) '批次處理，錯誤原因存入'dtVerRecord
        Catch ex As Exception
            Dim exMessage1 As String = ex.Message

            Dim strErrmsg As String = ""
            strErrmsg = ""
            strErrmsg &= "/* ex.ToString */" & vbCrLf
            strErrmsg &= ex.ToString & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)

            strErrmsg = ""
            strErrmsg &= "審核作業-未完成!!" & vbCrLf
            strErrmsg &= "Message:" & exMessage1 & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            Exit Sub
        End Try

        If Not gflag_ERROR_STOP Then
            Dim rqMID As String = TIMS.Get_MRqID(Me)
            'Response.Redirect("../04/TC_04_002.aspx?ID=" & Request("ID") & "")
            Common.RespWrite(Me, "<script>alert('審核作業完成!!');</script>")
            Common.RespWrite(Me, "<script>location.href='../04/TC_04_002.aspx?ID=" & rqMID & "'</script>")
        End If
    End Sub

    Function SaveEnterCCInfoDr(ByRef hPCS As Hashtable) As String
        Dim rst_errmsg As String = ""
        Dim vPlanID As String = TIMS.GetMyValue2(hPCS, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue2(hPCS, "ComIDNO")
        Dim vSeqNO As String = TIMS.GetMyValue2(hPCS, "SeqNO")
        Dim vAppliedResult As String = TIMS.GetMyValue2(hPCS, "AppliedResult")
        Dim v_PlanMode As String = TIMS.GetMyValue2(hPCS, "PlanMode")

        Dim flag_SaveCC As Boolean = False 'true:可執行轉入/false:不可執行
        Select Case vAppliedResult
            Case "Y"
                flag_SaveCC = True 'true:可執行轉入/false:不可執行
        End Select
        If Not flag_SaveCC Then '審核有誤不執行轉入
            '審核不通過 不執行轉入
            'Common.MessageBox(Me, "審核不通過 不執行轉入!")
            gflag_ERROR_STOP = True
            rst_errmsg = "審核不通過 不執行轉入!" 'Exit Sub
            Return rst_errmsg 'Exit Sub
        End If

        '檢核計畫與班級轉入
        flag_SaveCC = xChk_CC1(hPCS, rst_errmsg)
        If Not flag_SaveCC Then
            '審核有誤不執行轉入
            gflag_ERROR_STOP = True
            If rst_errmsg = "" Then rst_errmsg = "審核有誤不執行轉入!"
            Return rst_errmsg '"審核有誤不執行轉入!" 'Exit Sub
            'Exit Sub
        End If

        Dim drPP As DataRow = TIMS.GetPPInfo(vPlanID, vComIDNO, vSeqNO, objconn)
        '檢核報名日期 (若OK 轉出OUT SEnterDate/FEnterDate)
        Dim vSTDate As String = TIMS.Cdate3(drPP("STDate"))
        Dim vSEnterDate As String = "" 'TIMS.GetMyValue2(htCC, "SEnterDate")
        Dim vFEnterDate As String = "" 'TIMS.GetMyValue2(htCC, "FEnterDate") 'Dim flag_chkSEnDate As Boolean = False 'false:異常
        Call TIMS.ChangeSEnterDate(vSTDate, vSEnterDate, vFEnterDate)

        Dim flag_chkSEnDate As Boolean = If(vSEnterDate = "" OrElse vFEnterDate = "", False, True)  'false:異常
        If Not flag_chkSEnDate Then
            '報名時間有誤不執行轉入
            rst_errmsg = "開訓時間計算報名時間有誤 不可執行審核作業!"
            Return rst_errmsg '
            'Return rst_errmsg '"審核有誤不執行轉入!" 'Exit Sub
        End If

        'Dim htCC As New Hashtable
        'htCC.Clear()
        'htCC.Add("STDate", TIMS.cdate3(drPP("STDate")))
        ''檢核報名日期 (若OK 轉出OUT SEnterDate/FEnterDate)
        'Dim flag_chkSEnDate As Boolean = ChangSEnterDate(htCC)
        'Dim vSEnterDate As String = TIMS.GetMyValue2(htCC, "SEnterDate")
        'Dim vFEnterDate As String = TIMS.GetMyValue2(htCC, "FEnterDate")
        'If vSEnterDate = "" Then flag_chkSEnDate = False
        'If vFEnterDate = "" Then flag_chkSEnDate = False

        'PlanMode:S:審核中/Y:已通過/R:退件修正(含不通過的)
        If v_PlanMode = "Y" OrElse vAppliedResult = "Y" Then
            '小於、等於 開訓前三天 -不可報名
            Dim flag_chkSEnDate3 As Boolean = TIMS.ChkEnterDayS3(vSTDate)
            If Not flag_chkSEnDate3 Then
                '報名時間有誤不執行轉入
                rst_errmsg = "班級審核日距離開訓日為3日(含)內，不可執行審核作業!"
                'rst_errmsg = "開訓時間 小於、等於 開訓前三天 不可報名(不執行轉入)!"
                Return rst_errmsg '"審核有誤不執行轉入!" 'Exit Sub
            End If
        End If

        TIMS.SetMyValue2(hPCS, "SEnterDate", vSEnterDate)
        TIMS.SetMyValue2(hPCS, "FEnterDate", vFEnterDate)
        'TIMS.SetMyValue2(hPCS, "TechName", vTechName)
        Call SaveData1(hPCS) 'CLASS_CLASSINFO

        Return rst_errmsg
    End Function

    Private Sub DgPlan_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgPlan.ItemCommand
        GetSearchStr()

        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Select Case e.CommandName
            Case "Edit"
                TIMS.Utl_Redirect1(Me, cst_TC04002Add2aspx & "?ID=" & rqMID & "&todo=1&" & e.CommandArgument)
            Case "Add"
                TIMS.Utl_Redirect1(Me, cst_TC04002Add2aspx & "?ID=" & rqMID & "&todo=1&" & e.CommandArgument)
            Case Else
                '點選班別名稱
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '企訓專用
                    TIMS.Utl_Redirect1(Me, cst_TC03006Addaspx & "?ID=" & rqMID & "&todo=1&" & e.CommandArgument)
                End If
        End Select
    End Sub

    Private Sub DgPlan_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgPlan.ItemDataBound
        'Const Cst_資格複審 = "資格複審"
        'Const Cst_資格初審 = "資格初審"
        'Const Cst_複審狀況 = "複審狀況"
        'Const Cst_初審狀況 = "初審狀況"
        Const Cst_資格審核 As String = "資格審核"
        Const Cst_審核狀況 As String = "審核狀況"
        Dim drA As DataRowView = e.Item.DataItem
        'Dim objControl As WebControls.DropDownList
        Dim objReason As HtmlControls.HtmlTextArea = e.Item.FindControl("Reason")
        Dim LinkButton1 As LinkButton = e.Item.FindControl("LinkButton1")
        Dim Label12 As Label = e.Item.FindControl("Label12")
        Dim Label15 As Label = e.Item.FindControl("Label15")
        Dim BtnEdit As Button = e.Item.FindControl("BtnEdit")
        Dim BtnAdd As Button = e.Item.FindControl("BtnAdd")

        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim SelectAll1 As DropDownList = e.Item.FindControl("SelectAll1")
                'SelectAll1.Attributes("onchange") = "ChangeAll1(this.selectedIndex);"
                SelectAll1.Attributes.Add("onChange", "SelectAll_J();")
                '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
                Dim v_PlanMode As String = TIMS.GetListValue(PlanMode) 'PlanMode.SelectedValue
                If v_PlanMode = "R" Then
                    If sm.UserInfo.OrgLevel <= 1 Then '署(局):0／分署(中心):1
                        Label12.Text = Cst_資格審核
                        Label15.Text = Cst_審核狀況
                    Else
                        Label12.Text = ""
                    End If
                    SelectAll1.Enabled = False
                    bntAdd.Visible = False
                    TIMS.Tooltip(SelectAll1, "該功能不可全選")
                Else
                    If sm.UserInfo.OrgLevel <= 1 Then '署(局):0／分署(中心):1
                        Label12.Text = Cst_資格審核
                        Label15.Text = Cst_審核狀況
                        SelectAll1.Enabled = True
                        bntAdd.Visible = True
                    Else
                        Label12.Text = ""
                        SelectAll1.Enabled = False
                        TIMS.Tooltip(SelectAll1, "該功能不可全選")
                        bntAdd.Visible = False
                    End If
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem '帶入一行資料
                Dim labNumber1 As Label = e.Item.FindControl("labNumber1")
                'Dim Hid_PCS As HiddenField = e.Item.FindControl("Hid_PCS")
                labNumber1.Text = e.Item.ItemIndex + 1 + dgPlan.CurrentPageIndex * dgPlan.PageSize
                'e.Item.Cells(Cst_Index).Text = e.Item.ItemIndex + 1 + dgPlan.CurrentPageIndex * dgPlan.PageSize
                'Hid_PCS.Value = "PLANID=" & Convert.ToString(drv("PLANID")) & "&COMIDNO=" & Convert.ToString(drv("COMIDNO")) & "&SEQNO=" & Convert.ToString(drv("SEQNO"))
                Dim AppliedResult1 As DropDownList = e.Item.FindControl("AppliedResult1")
                '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU

                Dim lab_RESULTBUTTON As Label = e.Item.FindControl("lab_RESULTBUTTON")
                'Public Const cst_ResultButton_尚未送出_待送審 As String = "Y" '修改後可送出
                'Public Const cst_ResultButton_尚未送出_未送出 As String = "R" '還不可送出
                Dim sg_RESULTBUTTON As String = ""
                Dim tt_RESULTBUTTON As String = ""
                Select Case Convert.ToString(drv("RESULTBUTTON"))
                    Case TIMS.cst_ResultButton_尚未送出_待送審
                        sg_RESULTBUTTON = "#"
                        tt_RESULTBUTTON = "尚未送出-待送審"
                    Case TIMS.cst_ResultButton_尚未送出_未送出
                        sg_RESULTBUTTON = "*"
                        tt_RESULTBUTTON = "尚未送出-未送出"
                End Select
                lab_RESULTBUTTON.ForeColor = Color.Red
                lab_RESULTBUTTON.Text = sg_RESULTBUTTON
                If (tt_RESULTBUTTON <> "") Then TIMS.Tooltip(lab_RESULTBUTTON, tt_RESULTBUTTON, True)

                If sm.UserInfo.OrgLevel <= 1 Then '署(局):0／分署(中心):1
                    AppliedResult1.Enabled = True
                    Common.SetListItem(AppliedResult1, drv("SecResult").ToString)
                    '970509 Andy 產業人才投資方案,班級轉班已完成,不可取消審核
                    ' If (PlanMode.SelectedValue <> "R") Then
                    '署(局)還有分署(中心)
                    '且為產投
                    If (sm.UserInfo.OrgLevel <= 1 AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then
                        If Convert.ToString(drv("TransFlag2")) = "Y" Then
                            Dim v_AppliedResult1 As String = TIMS.GetListValue(AppliedResult1) 'AppliedResult1.SelectedValue
                            If v_AppliedResult1 <> "" AndAlso v_AppliedResult1 <> Cst_EmptySelValue Then
                                AppliedResult1.Enabled = False  '審核狀況
                                TIMS.Tooltip(AppliedResult1, "班級轉班已完成,不可取消審核")
                            End If
                        End If

                        'If e.Item.Cells(31).Text = "是" Then   'Cells(31)轉班
                        '    If AppliedResult1.SelectedValue <> "" _
                        '        AndAlso AppliedResult1.SelectedValue <> Cst_EmptySelValue Then
                        '        AppliedResult1.Enabled = False      '審核狀況
                        '        TIMS.Tooltip(AppliedResult1, "班級轉班已完成,不可取消審核")
                        '    End If
                        'End If
                    End If
                    'ElseIf sm.UserInfo.OrgLevel = 1 Then '分署(中心)
                    '    AppliedResult1.Enabled = True
                    '    Common.SetListItem(AppliedResult1, drv("FirResult").ToString)
                Else
                    AppliedResult1.Enabled = False
                    TIMS.Tooltip(AppliedResult1, "班級轉班 權限不足")
                    Common.SetListItem(AppliedResult1, drv("AppliedResult").ToString)
                End If
                'LinkButton1 = e.Item.FindControl("LinkButton1")
                LinkButton1.Text = Convert.ToString(drv("CLASSNAME2")) '.ToString
                'If drv("CyclType").ToString <> "" Then
                '    If Int(drv("CyclType")) <> 0 Then
                '        LinkButton1.Text += "第" & drv("CyclType").ToString & "期"
                '    End If
                'End If

                '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
                Select Case sm.UserInfo.LID.ToString
                    Case "0", "1" '署(職訓局)0 '分署(中心)1
                        BtnAdd.Visible = False
                        BtnEdit.Visible = True

                        BtnAdd.Enabled = False
                        BtnEdit.Enabled = True

                        '970509 Andy 產業人才投資方案,班級轉班已完成,不可取消審核
                        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                            If Convert.ToString(drv("TransFlag2")) = "Y" Then
                                '有選擇審核內容
                                Dim v_AppliedResult1 As String = TIMS.GetListValue(AppliedResult1) 'AppliedResult1.SelectedValue
                                If v_AppliedResult1 <> "" AndAlso v_AppliedResult1 <> Cst_EmptySelValue Then
                                    BtnEdit.Enabled = False  '資格複審
                                    BtnAdd.Enabled = False
                                    TIMS.Tooltip(BtnEdit, "班級轉班已完成,不可取消審核")
                                    TIMS.Tooltip(BtnAdd, "班級轉班已完成,不可取消審核")
                                End If
                            End If

                            'If e.Item.Cells(31).Text = "是" Then   'Cells(31)轉班
                            '    '有選擇審核內容
                            '    If AppliedResult1.SelectedValue <> "" _
                            '        AndAlso AppliedResult1.SelectedValue <> Cst_EmptySelValue Then
                            '        BtnEdit.Enabled = False  '資格複審
                            '        BtnAdd.Enabled = False
                            '        TIMS.Tooltip(BtnEdit, "班級轉班已完成,不可取消審核")
                            '        TIMS.Tooltip(BtnAdd, "班級轉班已完成,不可取消審核")
                            '    End If
                            'End If
                        End If
                End Select

                Dim str_CommandArgument As String = ""
                str_CommandArgument = "PlanYear=" & drv("PlanYear")
                str_CommandArgument += "&TPlanID=" & drv("TPlanID")
                str_CommandArgument += "&PlanID=" & drv("PlanID")
                str_CommandArgument += "&ComIDNO=" & drv("ComIDNO")
                str_CommandArgument += "&SeqNO=" & drv("SeqNO")
                str_CommandArgument += "&TMID=" & drv("TMID")
                str_CommandArgument += "&TNum=" & drv("TNum")
                str_CommandArgument += "&THours=" & drv("THours")
                str_CommandArgument += "&ProcID=" & drv("ProcID")
                str_CommandArgument += "&PointYN=" & drv("PointYN")
                str_CommandArgument += "&ClassCate=" & drv("ClassCate")
                str_CommandArgument += "&CapDegree=" & drv("CapDegree")
                str_CommandArgument += "&DefGovCost=" & drv("DefGovCost")
                str_CommandArgument += "&DefStdCost=" & drv("DefStdCost")
                str_CommandArgument += "&STDate=" & drv("STDate")
                str_CommandArgument += "&FDDate=" & drv("FDDate")
                str_CommandArgument += "&ClassName=" & HttpUtility.UrlEncode(drv("ClassName"))

                LinkButton1.CommandArgument = str_CommandArgument
                BtnEdit.CommandArgument = str_CommandArgument & "&CmdStatus=Edit"
                BtnAdd.CommandArgument = str_CommandArgument & "&CmdStatus=Add"
                Select Case drv("AppliedResult").ToString
                    Case "N"
                        objReason.Value = If(Convert.ToString(drv("Reason_all")) <> "", Convert.ToString(drv("Reason_all")), "不通過")
                    Case "Y"
                        objReason.Value = "通過"
                    Case Else
                        objReason.Value = Convert.ToString(drv("Reason_all"))
                End Select

            Case ListItemType.Footer
                If dgPlan.Items.Count = 0 Then
                    dgPlan.ShowFooter = True
                    Dim mycell As New TableCell
                    mycell.ColumnSpan = e.Item.Cells.Count
                    mycell.Text = "目前沒有任何資料!"
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(mycell)
                    e.Item.HorizontalAlign = HorizontalAlign.Center
                Else
                    dgPlan.ShowFooter = False
                End If
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim AppliedResult2 As DropDownList = e.Item.FindControl("AppliedResult2")
                Dim KeyValue As HtmlInputHidden = e.Item.FindControl("KeyValue")
                Dim KPlanID As HtmlInputHidden = e.Item.FindControl("KPlanID")
                Dim KComIDNO As HtmlInputHidden = e.Item.FindControl("KComIDNO")
                Dim KSeqNo As HtmlInputHidden = e.Item.FindControl("KSeqNo")

                '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
                '970509 Andy 產業人才投資方案,班級轉班已完成,不可取消審核
                If (PlanMode.SelectedValue <> "R") Then
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        If Convert.ToString(drv("TransFlag2")) = "Y" Then '轉班
                            AppliedResult2.Enabled = False  '取消審核
                            TIMS.Tooltip(AppliedResult2, "班級轉班已完成,不可取消審核")
                        End If
                        'If e.Item.Cells(6).Text = "是" Then 'Cells(6)轉班
                        '    AppliedResult2.Enabled = False  '取消審核
                        '    TIMS.Tooltip(AppliedResult2, "班級轉班已完成,不可取消審核")
                        'End If
                    End If
                End If

                e.Item.Cells(Cst_Index).Text = e.Item.ItemIndex + 1 + DataGrid2.CurrentPageIndex * DataGrid2.PageSize

                If Split(drv("relship"), "/").Length >= 3 Then
                    Dim ParenrRID As String = Split(drv("relship"), "/")(Split(drv("relship"), "/").Length - 3)
                    Dim ParentName As String = ""
                    If Auth_Relship.Select("RID='" & ParenrRID & "'").Length <> 0 Then
                        ParentName = Auth_Relship.Select("RID='" & ParenrRID & "'")(0)("OrgName")
                    End If
                    If ParentName <> "" Then
                        e.Item.Cells(Cst_ParentName).Text = "<font color='Blue'>" & ParentName & "</font>-" & drv("OrgName")
                    End If
                End If

                'LinkButton1.Text = Convert.ToString(drv("CLASSNAME2")) '.ToString
                'If IsNumeric(drv("CyclType")) Then
                '    If Int(drv("CyclType")) <> 0 Then
                '        e.Item.Cells(Cst_CyclType).Text += "第" & Int(drv("CyclType")) & "期"
                '    End If
                'End If
                KeyValue.Value = "PlanID='" & drv("PlanID").ToString & "' AND ComIDNO='" & drv("ComIDNO").ToString & "' AND SeqNo='" & drv("SeqNo").ToString & "'"
                KPlanID.Value = "" & drv("PlanID").ToString
                KComIDNO.Value = "" & drv("ComIDNO").ToString
                KSeqNo.Value = "" & drv("SeqNo").ToString

            Case ListItemType.Footer
                If DataGrid2.Items.Count = 0 Then
                    DataGrid2.ShowFooter = True
                    Dim mycell As New TableCell

                    mycell.ColumnSpan = e.Item.Cells.Count
                    mycell.Text = "目前沒有任何資料!"
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(mycell)
                    e.Item.HorizontalAlign = HorizontalAlign.Center
                Else
                    DataGrid2.ShowFooter = False
                End If
        End Select
    End Sub

    '取消審核儲存鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Select Case TIMS.GetVerType(sm.UserInfo.LID)
            Case "F", "S"
            Case Else
                Common.MessageBox(Me, "該登入者無取消審核權限!!")
                Exit Sub
        End Select

        For Each Item As DataGridItem In DataGrid2.Items
            Dim AppliedResult2 As DropDownList = Item.FindControl("AppliedResult2")
            Dim KeyValue As HtmlInputHidden = Item.FindControl("KeyValue")
            Dim KPlanID As HtmlInputHidden = Item.FindControl("KPlanID")
            Dim KComIDNO As HtmlInputHidden = Item.FindControl("KComIDNO")
            Dim KSeqNo As HtmlInputHidden = Item.FindControl("KSeqNo")

            If AppliedResult2.SelectedIndex <> 0 Then
                TIMS.Plan_VerRecord_Update(KPlanID.Value, KComIDNO.Value, KSeqNo.Value, objconn)
                'AppliedResult2@VALUE:取消審核(O)
                TIMS.PLAN_VERREPROT_UPDATE(Me, KPlanID.Value, KComIDNO.Value, KSeqNo.Value, "O", objconn)
            End If
        Next

        Common.MessageBox(Me, "儲存成功")
        'btnQuery_Click(sender, e)
        Call SSearch1()
    End Sub

    'Function Get_DistIDValue() As String
    '    'Dim rst As String = ""
    '    Dim value As String = ""
    '    For i As Integer = 0 To DistrictList.Items.Count - 1
    '        If DistrictList.Items.Item(i).Selected Then
    '            If value <> "" Then value += ","
    '            value += "'" & DistrictList.Items.Item(i).Value & "'"
    '        End If
    '    Next
    '    'rst = value
    '    Return value 'rst
    'End Function

    Sub GetSearchStr()
        'KeepSearch
        Dim v_DistType As String = TIMS.GetListValue(DistType)
        Dim v_DistrictList As String = TIMS.GetListValue(DistrictList)
        Dim v_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        '依申請階段
        Dim v_AppStage2 As String = TIMS.GetListValue(AppStage2)
        Dim v_PlanMode As String = TIMS.GetListValue(PlanMode)
        Dim v_AdvanceMode As String = TIMS.GetListValue(AdvanceMode)

        Dim str_KeepSearch As String = ""
        str_KeepSearch = "ks=1"
        str_KeepSearch &= "&prg=TC_04_002"
        str_KeepSearch &= "&DistType=" & v_DistType '.SelectedValue
        str_KeepSearch &= "&DistrictList=" & v_DistrictList '.SelectedValue
        str_KeepSearch &= "&DistHidden=" & TIMS.ClearSQM(DistHidden.Value)
        str_KeepSearch &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        str_KeepSearch &= "&center=" & TIMS.ClearSQM(center.Text)
        str_KeepSearch &= "&TB_career_id=" & TIMS.ClearSQM(TB_career_id.Text)
        str_KeepSearch &= "&TPlanid=" & TIMS.ClearSQM(TPlanid.Value)
        str_KeepSearch &= "&trainValue=" & TIMS.ClearSQM(trainValue.Value)
        str_KeepSearch &= "&jobValue=" & TIMS.ClearSQM(jobValue.Value)
        str_KeepSearch &= "&txtCJOB_NAME=" & TIMS.ClearSQM(txtCJOB_NAME.Text)
        str_KeepSearch &= "&cjobValue=" & TIMS.ClearSQM(cjobValue.Value)
        str_KeepSearch &= "&ClassName=" & TIMS.ClearSQM(ClassName.Text)
        str_KeepSearch &= "&CyclType=" & TIMS.ClearSQM(CyclType.Text)
        str_KeepSearch &= "&UNIT_SDATE=" & TIMS.ClearSQM(UNIT_SDATE.Text)
        str_KeepSearch &= "&UNIT_EDATE=" & TIMS.ClearSQM(UNIT_EDATE.Text)
        str_KeepSearch &= "&start_date=" & TIMS.ClearSQM(start_date.Text)
        str_KeepSearch &= "&end_date=" & TIMS.ClearSQM(end_date.Text)
        str_KeepSearch &= "&OrgKind2=" & v_OrgKind2 '.SelectedValue
        str_KeepSearch &= "&AppStage2=" & v_AppStage2 '.SelectedValue
        str_KeepSearch &= "&PlanMode=" & v_PlanMode '.SelectedValue
        str_KeepSearch &= "&AdvanceMode=" & v_AdvanceMode 'AdvanceMode.SelectedValue
        str_KeepSearch &= "&btnQuery=1"

        Session("search") = str_KeepSearch

        'HttpUtility.UrlEncode(

        'Session("_Search") = "PlanYear=" & yearlist.SelectedValue
        'Session("_Search") += "&TxtPageSize=" & TxtPageSize.Text
        'Session("_Search") += "&PageIndex=" & DG_Org.CurrentPageIndex + 1
        'Session("_Search") += "&Button1=" & DG_Org.Visible
    End Sub

    'Protected Sub DataGrid2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid2.SelectedIndexChanged

    'End Sub

#Region "班級轉入動作--產投使用"

    ''' <summary>
    ''' 儲存前檢核: false:異常/true:正常--班級轉入
    ''' </summary>
    ''' <returns></returns>
    Function xChk_CC1(ByRef hPCS As Hashtable, ByRef rErrMsg As String) As Boolean
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        '檢核登入者的計畫 異常為False
        If Not TIMS.ChkTPlanID28(sm) Then
            rErrMsg = cst_errmsg4
            'Common.MessageBox(Me, cst_errmsg4)
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
        '    sql = "SELECT 'X' FROM CLASS_CLASSINFO WHERE CLSID='" & clsid.Value & "' AND PlanID='" & PlanID.Value & "' AND CyclType='" & CyclType.Text & "' AND RID='" & RIDValue.Value & "' AND OCID!='" & rq_OCID & "'"
        'Else
        '    sql = "SELECT 'X' FROM CLASS_CLASSINFO WHERE CLSID='" & clsid.Value & "' AND PlanID='" & PlanID.Value & "' AND CyclType='" & CyclType.Text & "' AND RID='" & RIDValue.Value & "'"
        'End If
        'Dim dtX As DataTable = DbAccess.GetDataTable(sql, objconn)
        'If dtX.Rows.Count > 0 Then
        '    Common.MessageBox(Me, "新增開班資料重複(該機構在當年度計畫有相同的班別代碼與期別!!)")
        '    Return False 'Exit Sub
        'End If

        Dim vPlanID As String = TIMS.GetMyValue2(hPCS, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue2(hPCS, "ComIDNO")
        Dim vSeqNO As String = TIMS.GetMyValue2(hPCS, "SeqNO")

        Dim sPMS As New Hashtable From {{"PlanID", vPlanID}, {"ComIDNO", vComIDNO}, {"SeqNO", vSeqNO}}
        Dim sql As String = ""
        sql &= " SELECT 'X' "
        sql &= " FROM dbo.CLASS_CLASSINFO"
        sql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO"
        Dim dtCC As DataTable = DbAccess.GetDataTable(sql, objconn, sPMS)
        If dtCC.Rows.Count > 0 Then
            'Common.MessageBox(Me, "新增開班資料重複(已有轉班資料!!!)")
            rErrMsg = "新增開班資料重複(已有轉班資料!!!)"
            Return False 'Exit Sub
        End If
        Return True
    End Function

    ''' <summary>儲存(CLASS_CLASSINFO) --班級轉入</summary>
    Public Sub SaveData1(ByRef hPCS As Hashtable)
        If hPCS Is Nothing Then Return

        'Dim rq_OCID As String = TIMS.GetMyValue2(htCC, "OCID")
        'in:
        Dim vPlanID As String = TIMS.GetMyValue2(hPCS, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue2(hPCS, "ComIDNO")
        Dim vSeqNO As String = TIMS.GetMyValue2(hPCS, "SeqNO")
        'Dim vTechName As String = TIMS.GetMyValue2(hPCS, "TechName")
        Dim vSEnterDate As String = TIMS.GetMyValue2(hPCS, "SEnterDate")
        Dim vFEnterDate As String = TIMS.GetMyValue2(hPCS, "FEnterDate")

        Dim dr1 As DataRow = TIMS.GetPPInfo(vPlanID, vComIDNO, vSeqNO, objconn)
        If dr1 Is Nothing OrElse vSEnterDate = "" OrElse vFEnterDate = "" Then Return 'Exit Sub

        Dim vRIDValue As String = Convert.ToString(dr1("RID"))
        Dim vRelship As String = TIMS.GET_RelshipforRID(vRIDValue, objconn)
        'RIDValue.Value

        'htPV.Clear()
        Dim htPV As New Hashtable 'out
        htPV.Add("RID", Convert.ToString(dr1("RID")))
        htPV.Add("TMID", Convert.ToString(dr1("TMID")))
        htPV.Add("TPLANID", sm.UserInfo.TPlanID)
        htPV.Add("DISTID", sm.UserInfo.DistID)
        htPV.Add("YEARS", sm.UserInfo.Years)
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

        Dim pms_pt As New Hashtable() From {{"PlanID", vPlanID}, {"ComIDNO", vComIDNO}, {"SeqNo", vSeqNO}}
        Dim sql_pt As String = ""
        sql_pt &= " SELECT * FROM PLAN_TEACHER"
        sql_pt &= " WHERE TechTYPE='A'" 'TechTYPE: A:師資/B:助教
        sql_pt &= " AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dtPt As DataTable = DbAccess.GetDataTable(sql_pt, objconn, pms_pt)

        Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        'Call TIMS.OpenDbConn(tConn)
        'Dim sql As String = ""
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim tTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            '2006/03/ add conn by matt
            Dim iOCID_New As Integer = 0
            iOCID_New = DbAccess.GetNewId(tTrans, "CLASS_CLASSINFO_OCID_SEQ,CLASS_CLASSINFO,OCID") 'fix ora-00001 違反必須唯一的限制條件

            Dim vHtClsid As Hashtable = TIMS.Get_ClassIDG28(htPV, tTrans)
            Dim vCLSID As String = TIMS.GetMyValue2(vHtClsid, "CLSID")
            'Dim vCLSID As String = TIMS.GetMyValue2(vHtClsid, "CLSID")

            Dim dr As DataRow = Nothing
            Dim dt As DataTable = Nothing
            Dim da As SqlDataAdapter = Nothing
            Dim sql_c As String = "SELECT * FROM CLASS_CLASSINFO WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql_c, da, tTrans)
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = iOCID_New

            dr("CLSID") = vCLSID 'vClsid.Value
            dr("PlanID") = dr1("PlanID")
            dr("Years") = Right(dr1("PlanYear"), 2)
            Dim vCyclType As String = Convert.ToString(dr1("CyclType")) ' If vCyclType = "" Then vCyclType = "01"
            dr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
            dr("ClassNum") = "01"

            dr("RID") = dr1("RID")
            dr("ClassCName") = dr1("ClassName")
            'dr("CJOB_UNKEY") = dr9("CJOB_UNKEY")  '通俗職類
            'If ClassEngName <> "" Then
            '    dr("ClassEngName") = ClassEngName
            'Else
            '    dr("ClassEngName") = ""
            'End If
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

            dr("SEnterDate") = TIMS.Cdate2(vSEnterDate)
            dr("FEnterDate") = TIMS.Cdate2(vFEnterDate)
            '上架日期
            'If vsOnShellDate <> "" Then dr("OnShellDate") = vsOnShellDate
            'dr("SEnterDate") = SEnterDate.Text
            'dr("FEnterDate") = FEnterDate.Text
            'If CheckInDate.Text <> "" Then
            '    dr("CheckInDate") = CheckInDate.Text
            'Else
            '    dr("CheckInDate") = STDate.Text
            'End If
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
            'If TPeriod.SelectedValue <> "" Then
            '    dr("TPeriod") = TPeriod.SelectedValue
            'Else
            '    dr("TPeriod") = TPeriod.SelectedValue
            'End If
            dr("NotOpen") = "N"
            dr("NORID") = Convert.DBNull
            dr("OtherReason") = Convert.DBNull
            dr("IsApplic") = "N"

            dr("Relship") = vRelship
            dr("ComIDNO") = vComIDNO '.Value '(PCS)
            dr("SeqNO") = vSeqNO '.Value '(PCS)

            dr("IsCalculate") = "Y"
            dr("IsSuccess") = "Y"
            'TechName.Text = TIMS.ClearSQM(TechName.Text)
            dr("IsFullDate") = "N" '產學訓預設值為否
            'dr("CTName") = Left("X:" & TechName.Text, 40) 
            'dr("CTName") = Left(vTechName, 40)
            dr("CTName") = " "

            dr("IsBusiness") = "N" 'IIf(IsBusiness.Checked = True, "Y", "N")
            'dr("EnterpriseName") = IIf(EnterpriseName.Text <> "", EnterpriseName.Text, Convert.DBNull)

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, tTrans)
            'If rq_OCID <> "" Then iOCID_New = Val(rq_OCID)

            Dim sql_p As String = "SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & vPlanID & "' AND ComIDNO='" & vComIDNO & "' AND SeqNo='" & vSeqNO & "'"
            dt = DbAccess.GetDataTable(sql_p, da, tTrans)
            dr = dt.Rows(0)
            'dr("CredPoint") = 0 'CredPoint.Text
            'dr("RoomName") = Convert.DBNull 'RoomName.Text
            'dr("FactMode") = Convert.DBNull 'FactMode.SelectedValue
            'dr("FactModeOther") = Convert.DBNull '
            'dr("ConNum") = ConNum.Text
            'dr("ContactName") = ContactName.Text
            'dr("ContactPhone") = ContactPhone.Text
            'dr("ContactEmail") = ContactEmail.Text
            'dr("ContactFax") = ContactFax.Text
            'dr("ClassCate") = ClassCate.SelectedValue
            dr("TransFlag") = "Y"
            '正常後關閉****************
            'dr("IsBusiness") = "N" 'IIf(IsBusiness.Checked = True, "Y", "N")
            'dr("EnterpriseName") = IIf(EnterpriseName.Text <> "", EnterpriseName.Text, Convert.DBNull) '企業包班
            '**************************
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, tTrans)

            'sql = "SELECT * FROM PLAN_ONCLASS WHERE 1<>1"
            'dt = DbAccess.GetDataTable(sql, da, trans)
            'Dim TempDataTable As DataTable = Session("Plan_OnClass")
            'dt = TempDataTable.Copy
            'DbAccess.UpdateDataTable(dt, da, trans)

            '儲存 班級申請老師(CLASS_TEACHER)
            Call SAVE_CLASS_TEACHER(iOCID_New, dtPt, tTrans)

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
            'Call TIMS.CloseDbConn(tConn)
        Catch ex As Exception
            Dim exMessage1 As String = ex.Message

            Dim strErrmsg As String = ""
            strErrmsg &= "/* ex.ToString */" & vbCrLf
            strErrmsg &= ex.ToString & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)

            DbAccess.RollbackTrans(tTrans)
            DbAccess.CloseDbConn(tConn)

            strErrmsg = ""
            strErrmsg &= "儲存失敗-未完成!!" & vbCrLf
            strErrmsg &= "Message:" & exMessage1 & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            Exit Sub
            'Call TIMS.CloseDbConn(tConn)
            'DbAccess.RollbackTrans(trans)
            'Throw ex
        End Try

        '重複 為 true
        Dim Double_flag As Boolean = False 'false 沒有重複。
        Dim iDouble As Integer = 0
        Do
            iDouble += 1
            Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
            '刪除重複轉班資料。
            Try
                '至少判斷1次是否有重複轉班
                Double_flag = TIMS.sUtl_DeleteDoubleClassInfo(vPlanID, vComIDNO, vSeqNO, objconn)
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

    ''' <summary>
    ''' 儲存 班級申請老師(CLASS_TEACHER)--班級轉入
    ''' </summary>
    ''' <param name="iOCID_New"></param>
    ''' <param name="dtPt"></param>
    ''' <param name="trans"></param>
    Public Sub SAVE_CLASS_TEACHER(ByVal iOCID_New As Integer, ByVal dtPt As DataTable, ByRef trans As SqlTransaction)
        '更新師資表---------------------------------------Start
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = "DELETE CLASS_TEACHER WHERE OCID='" & iOCID_New & "'"
        DbAccess.ExecuteNonQuery(sql, trans)
        If dtPt.Rows.Count > 0 Then
            sql = "SELECT * FROM CLASS_TEACHER WHERE OCID='" & iOCID_New & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)
            'Dim i As Integer
            For i As Integer = 0 To dtPt.Rows.Count - 1
                If dt.Select("TechID='" & dtPt.Rows(i).Item("TechID") & "'").Length = 0 Then
                    Dim iCTRID As Integer = DbAccess.GetNewId(trans, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("CTRID") = iCTRID
                    dr("OCID") = iOCID_New
                    dr("TechID") = dtPt.Rows(i).Item("TechID")
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
            Next
            DbAccess.UpdateDataTable(dt, da, trans)
            'Else
            '    CTName.Value = TIMS.ClearSQM(CTName.Value)
            '    If CTName.Value <> "" Then
            '        sql = "SELECT * FROM CLASS_TEACHER WHERE OCID='" & iOCID_New & "'"
            '        dt = DbAccess.GetDataTable(sql, da, trans)
            '        For i As Integer = 0 To Split(CTName.Value, ",").Length - 1
            '            If dt.Select("TechID='" & Split(CTName.Value, ",")(i) & "'").Length = 0 Then
            '                Dim iCTRID As Integer = DbAccess.GetNewId(trans, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
            '                dr = dt.NewRow
            '                dt.Rows.Add(dr)
            '                dr("CTRID") = iCTRID
            '                dr("OCID") = iOCID_New
            '                dr("TechID") = Split(CTName.Value, ",")(i)
            '                dr("ModifyAcct") = sm.UserInfo.UserID
            '                dr("ModifyDate") = Now
            '            End If
            '        Next
            '        DbAccess.UpdateDataTable(dt, da, trans)
            '    End If
        End If
        '更新師資表---------------------------------------End

    End Sub

    Protected Sub DistType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DistType.SelectedIndexChanged
        trDist.Visible = True
        trOrg.Visible = False
        Dim v_DistType As String = TIMS.GetListValue(DistType)
        If v_DistType = "1" Then
            trDist.Visible = False
            trOrg.Visible = True
        End If
        DataGridTable1.Visible = False
        DataGridTable2.Visible = False
    End Sub

    Private Sub PlanMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlanMode.SelectedIndexChanged
        Select Case PlanMode.SelectedValue
            Case "Y" '已通過
                TRA.Visible = True
            Case Else
                TRA.Visible = False
        End Select
        DataGridTable1.Visible = False
        DataGridTable2.Visible = False
        'btnQuery_Click(sender, e)
        'Call sSearch1()
    End Sub

    Protected Sub dgPlan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dgPlan.SelectedIndexChanged

    End Sub

#End Region

End Class