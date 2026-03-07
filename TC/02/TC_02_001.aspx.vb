Partial Class TC_02_001
    Inherits AuthBasePage

    '申請階段管理-受理期間設定 PLAN_APPSTAGE-APPLISTAGE 查無:false 有:true
    'Dim fg_can_applistage_G As Boolean = False
    Dim flag_EnterShelf As Boolean = False '在職-批次轉班上架
    'Dim CPdt As DataTable 'SearchData1
    Const Cst_index As Integer = 0
    Const Cst_checkbox As Integer = 1 '選取-'在職-批次轉班上架
    Const Cst_PlanYear As Integer = 2 '計畫年度
    Const Cst_AppliedDate As Integer = 3 '申請日期
    Const Cst_STDate As Integer = 4 '訓練起日
    Const Cst_FDDate As Integer = 5 '訓練迄日
    Const Cst_OrgName2 As Integer = 6 '管控單位機構名稱(補助地方政府)-機構名稱
    Const Cst_OrgName As Integer = 7 '機構名稱
    Const Cst_OCID As Integer = 8 '課程代碼
    Const Cst_ClassName As Integer = 9 '班級名稱
    Const Cst_AppliedResult As Integer = 10 '班級審核狀態
    Const Cst_VerReason As Integer = 11 '未通過原因
    Const Cst_TransFlag As Integer = 12 '是否轉班 已轉班
    Const Cst_Function1 As Integer = 13 '功能1-正式查詢 修改:update/刪除:Del/列印:Print/送出:Send/還原:Return/經費明細:Def/轉班上架:Shelf
    Const Cst_Function2 As Integer = 14 '功能2-草稿查詢 修改:btnEdit/刪除:btnDel
    Const cst_Sort As String = "Sort"
    ' e.CommandName

    Const cst_lbtUpdate_Txt_查詢 As String = "查詢"
    Const cst_lbtUpdate_Txt_修改 As String = "修改"
    '正式操作鈕
    Const cst_str_view As String = "view" '查詢('修改)
    Const cst_str_update As String = "update" '修改
    Const cst_str_Del As String = "Del" '刪除
    Const cst_str_Print As String = "Print" '列印 '列印訓練計畫
    Const cst_str_Send As String = "Send" '送出-產投
    Const cst_str_Return As String = "Return" '還原-產投
    Const cst_str_Def As String = "Def" '列印 '經費明細表
    Const cst_str_Shelf As String = "Shelf" '轉班上架'班級轉入 

    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim flag_amu_test As Boolean = False
    'Dim au As New cAUTH
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'center.Attributes.Add("onfocus", "this.blur();")
        'TB_career_id.Attributes.Add("onfocus", "this.blur();")
        'txtCJOB_NAME.Attributes.Add("onfocus", "this.blur();")
        'MyInput.Attributes.Add("style", "background-color: #F3F1F1; color: #584B4B; cursor: default; border: 2px solid #ccc;");
        TIMS.INPUT_ReadOnly2(center)
        TIMS.INPUT_ReadOnly2(TB_career_id)
        TIMS.INPUT_ReadOnly2(txtCJOB_NAME)

        '1.自辦 '2.委外
        Dim PlanKind As String = TIMS.Get_PlanKind(Me, objconn)
        '取得訓練計畫
        TPlanid.Value = sm.UserInfo.TPlanID
        '管控單位(補助地方政府)
        dtPlan.Columns(Cst_OrgName2).Visible = True
        dtPlan.Columns(Cst_OrgName2).Visible = If(PlanKind.Equals("1"), False, True)
        PageControler1.PageDataGrid = dtPlan '分頁設定

        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)

        'Dim flag_amu_test As Boolean 
        'flag_amu_test = TIMS.sUtl_ChkTest("amu_test") '測試2
        iPYNum = TIMS.sUtl_GetPYNum(Me)
        tr_audit1.Visible = False '產投-審核狀態

        flag_EnterShelf = False '在職-批次轉班上架
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_EnterShelf = True '在職-批次轉班上架
        If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_EnterShelf = True '07:接受企業委託訓練
        'btnEnter1.Visible = If(flag_EnterShelf, True, False)

        ' <asp:Button ID="btnEnter1" runat="server" Text="批次轉班上架" CssClass="asp_Export_M"></asp:Button>
        btnExport2.Visible = False '在職進修訓練
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then btnExport2.Visible = True '匯出開班預定表
        'If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then btnExport2.Visible = True '匯出開班預定表

        btnExport1.Visible = False
        '產投/非產投判斷
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投
            LabTMID.Text = "訓練業別"
            'audit1.Visible = False
            IsApprPaper.AutoPostBack = False
            btnExport1.Visible = True
        Else
            '非產投 , 如果是選擇正式則出現選擇審核中或己審核的 選單 penny 2007/10/17
            IsApprPaper.Attributes("onclick") = "checkaudit1();"
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        '帶入查詢參數
        If Not IsPostBack Then
            Call cCreate1()
        End If

        '因有傳入值 yearlist.SelectedValue.ToString 故放此位置，才可讀到值
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), TIMS.GetListValue(yearlist))

        '確認機構是否為黑名單
        Dim StrMsg2 As String = ""
        If Chk_OrgBlackList(StrMsg2) Then
            Dim strScript As String = String.Concat("<script>alert('", StrMsg2, "');</script>")
            Page.RegisterStartupScript("", strScript)
        End If
    End Sub

    Sub cCreate1()
        hid_PPINFOtable_guid1.Value = TIMS.GetGUID()
        Session(hid_PPINFOtable_guid1.Value) = Nothing

        '未查詢前一律不顯示
        btnEnter1.Visible = False

        Me.msg.Text = ""
        DataGridTable.Visible = False
        Select Case sm.UserInfo.LID
            Case "2" '委訓單位
                trSearchYear.Style("display") = "none" '不提供年度顯示
        End Select
        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years) '2005/4/1--Melody年度帶預設值
        '(加強操作便利性)
        RIDValue.Value = sm.UserInfo.RID
        center.Text = sm.UserInfo.OrgName

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        If tr_AppStage_TP28.Visible Then
            AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
        End If

        Call UseKeepSearchStr()
    End Sub

    '已列入處分名單 提醒功能
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = False
        Errmsg = ""
        Dim StrComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, StrComIDNO, objconn) Then
            Rst = True
            Dim StrMsg2 As String = String.Concat(sm.UserInfo.OrgName, "，已列入處分名單!!")
            Me.isBlack.Value = "Y"
            Me.orgname.Value = sm.UserInfo.OrgName
            Errmsg = StrMsg2
        End If
        Return Rst
    End Function

    ''' <summary>依目前排序 顯示正確的排序圖型</summary>
    ''' <param name="sortVal"></param>
    ''' <returns></returns>
    Function GET_ImageUrl_UD(ByRef sortVal As String, ByRef str_Sort As String) As String
        Return If(str_Sort.Equals(sortVal), "../../images/SortUp.gif", "../../images/SortDown.gif")
    End Function

    Sub ACT_ImageUrl_UD(ByRef mysort As System.Web.UI.WebControls.Image, ByRef i_Cell As Integer, ByRef str_Sort As String)
        Select Case str_Sort
            Case "ClassName", "ClassName DESC"
                i_Cell = Cst_ClassName '8
                mysort.ImageUrl = GET_ImageUrl_UD("ClassName", str_Sort)
            Case "AppliedDate", "AppliedDate DESC"
                i_Cell = Cst_AppliedDate '2
                mysort.ImageUrl = GET_ImageUrl_UD("AppliedDate", str_Sort)
            Case "STDate", "STDate DESC"
                i_Cell = Cst_STDate '4
                mysort.ImageUrl = GET_ImageUrl_UD("STDate", str_Sort)
            Case "FDDate", "FDDate DESC"
                i_Cell = Cst_FDDate '5
                mysort.ImageUrl = GET_ImageUrl_UD("FDDate", str_Sort)
            Case "OrgName", "OrgName DESC"
                i_Cell = Cst_OrgName '7
                mysort.ImageUrl = GET_ImageUrl_UD("OrgName", str_Sort)
        End Select
    End Sub

    ''' <summary>審核文字說明</summary>
    ''' <param name="s_TPLANID"></param>
    ''' <param name="s_AppliedResult"></param>
    ''' <param name="s_RESULTBUTTON"></param>
    ''' <returns></returns>
    Function Get_AppliedResultTxt(ByRef s_TPLANID As String, ByRef s_AppliedResult As String, ByRef s_RESULTBUTTON As String) As String
        Dim rst As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(s_TPLANID) > -1 Then
            '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
            'Dim strMsg As String = "班級審核中" '= ""

            rst = "班級審核中" '= ""
            Dim flag_AppliedResult_Red_1 As Boolean = False '(紅字加強)
            Select Case s_AppliedResult
                Case "Y"
                    rst = "班級審核通過" 'e.Item.Cells(Cst_AppliedResult).Text += strMsg '"班級審核通過"
                Case "N"
                    flag_AppliedResult_Red_1 = True
                    rst = "班級審核不通過" 'e.Item.Cells(Cst_AppliedResult).Text &= "<font color=red>" & strMsg & "</font>"
                Case "R"
                    rst = "班級退件修正" 'e.Item.Cells(Cst_AppliedResult).Text += strMsg ' "班級退件修正"
                Case "M"
                    rst = "請修正資料" 'e.Item.Cells(Cst_AppliedResult).Text += strMsg '"請修正資料"
                Case "O"
                    '產投為審核中狀態。
                    rst = "班級審核中(審核後修正)" 'e.Item.Cells(Cst_AppliedResult).Text += strMsg '"班級審核中(審核後修正)" '"審核後修正"
                Case Else
                    's_RESULTBUTTON Y/R 
                    Select Case s_RESULTBUTTON'Convert.ToString(drv("RESULTBUTTON"))
                        Case TIMS.cst_ResultButton_尚未送出_待送審 'Y
                            rst = "待送審"
                        Case TIMS.cst_ResultButton_尚未送出_未送出 'R
                            flag_AppliedResult_Red_1 = True
                            rst = "(未正式儲存)"
                            'Case Else 'strMsg = "班級審核中"
                    End Select
                    'e.Item.Cells(Cst_AppliedResult).Text += strMsg '"班級審核中"
            End Select
            'e.Item.Cells(Cst_AppliedResult).Text = Get_AppliedResultTxt(sm.UserInfo.TPlanID, Convert.ToString(drv("AppliedResult")))
            If flag_AppliedResult_Red_1 Then rst = String.Concat("<font color=red>", rst, "</font>")
            'e.Item.Cells(Cst_AppliedResult).Text = strMsg
        Else
            '非產投
            If s_AppliedResult = "" Then
                rst = "審核中"
                Return rst
            End If

            Select Case s_AppliedResult 'drv("AppliedResult")
                Case "Y"
                    rst = "審核通過"
                Case "N"
                    rst = "審核不通過"
                Case "R"
                    rst = "退件修正"
                Case "M"
                    rst = "請修正資料"
                Case "O"
                    rst = "審核後修正"
            End Select
        End If
        Return rst
    End Function

    '產投(企訓專用)
    Function GET_URL1(ByRef s_CmdArg As String, ByRef rqMID As String) As String
        'Dim rqMID As String = TIMS.Get_MRqID(Me)
        Dim YEARS As String = TIMS.GetMyValue(s_CmdArg, "YEARS")
        If YEARS = "" Then Throw New Exception("資料有誤，請重新檢查資料!!")

        '產投(企訓專用)
        Dim url1 As String = "../03/TC_03_006.aspx?ID=" & rqMID
        If YEARS < "2018" Then url1 = "../03/TC_03_003.aspx?ID=" & rqMID
        Return url1
    End Function

    Public Sub dtPlan_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dtPlan.ItemCommand
        'Call TIMS.OpenDbConn(objconn)
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Select Case e.CommandName
            Case cst_str_Shelf '轉班上架'班級轉入 
                Dim sCmdArg As String = e.CommandArgument
                TIMS.LOG.Debug(String.Format("#sCmdArg:(dtPlan_ItemCommand).{0}", sCmdArg))
                If sCmdArg = "" Then Return

                Dim vPCS As String = TIMS.GetMyValue(sCmdArg, "PCS") 'concat(p1.planid,'x',p1.ComIDNO,'x',p1.SeqNO
                'Dim vCLSID As String = TIMS.GetMyValue2(htSS, "CLSID")
                Dim vCJOB_UNKEY As String = TIMS.GetMyValue(sCmdArg, "CJOB_UNKEY")
                Dim vPlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
                Dim vCOMIDNO As String = TIMS.GetMyValue(sCmdArg, "COMIDNO")
                Dim vSEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")
                Dim vTPlanID As String = TIMS.GetMyValue(sCmdArg, "TPlanID")
                'Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
                'Dim vAPPSTAGE As String = TIMS.GetMyValue(sCmdArg, "APPSTAGE")
                Dim vRID As String = TIMS.GetMyValue(sCmdArg, "RID")
                Dim vRID1 As String = vRID.Substring(0, 1)

                Dim parms As New Hashtable From {{"PCS", vPCS}}
                'parms.Add("PCS", vPCS) 'concat(p1.planid,'x',p1.ComIDNO,'x',p1.SeqNO
                parms.Add("CJOB_UNKEY", vCJOB_UNKEY)
                parms.Add("PlanID", vPlanID)
                parms.Add("COMIDNO", vCOMIDNO)
                parms.Add("SEQNO", vSEQNO)
                parms.Add("TPlanID", vTPlanID)
                parms.Add("RID1", vRID1)
                '按鈕單筆-轉班上架/批次轉班上架
                Dim ErrMessage As String = Utl_EntreShelf2(parms)
                If ErrMessage <> "" Then
                    Common.MessageBox(Me, ErrMessage)
                    Return ' Exit Sub
                End If
                Common.MessageBox(Me, "班級轉入成功")
                Call SearchData1()

            Case cst_str_view, cst_str_update '"update", "view" '檢視/修改(正式)
                Call KeepSearchStr()
                Dim sCmdArg As String = e.CommandArgument
                If sCmdArg = "" Then Return

                Dim vPLANID As String = TIMS.GetMyValue(sCmdArg, "PLANID")
                Dim vCOMIDNO As String = TIMS.GetMyValue(sCmdArg, "COMIDNO")
                Dim vSEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")
                Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
                Dim vAPPSTAGE As String = TIMS.GetMyValue(sCmdArg, "APPSTAGE")

                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If e.CommandName = cst_str_update AndAlso sm.UserInfo.LID <> 0 Then
                        '修改檢核 '申請階段管理-受理期間設定 APPLISTAGE
                        Dim aParms As New Hashtable From {{"YEARS", vYEARS}, {"APPSTAGE", vAPPSTAGE}}
                        '開放受理之申請階段／PLAN_APPSTAGE
                        Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE01(objconn, aParms)
                        '檢核查詢 '開放受理之申請階段／PLAN_APPSTAGE
                        If Not fg_can_applistage Then
                            Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                            Return
                        End If
                    End If
                End If

                Dim url1 As String = ""
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '產投(企訓專用)
                    url1 = String.Concat(GET_URL1(sCmdArg, rqMID), "&PlanID=", vPLANID, "&ComIDNO=", vCOMIDNO, "&SeqNO=", vSEQNO, "&YEARS=", vYEARS)
                    If e.CommandName = cst_str_view Then url1 &= "&State=View" '"view" 
                    Call TIMS.Utl_Redirect(Me, objconn, url1)
                Else
                    'TIMS 
                    url1 = String.Concat("../03/TC_03_001.aspx?ID=", rqMID, "&PlanID=", vPLANID, "&ComIDNO=", vCOMIDNO, "&SeqNO=", vSEQNO, "&YEARS=", vYEARS)
                    Call TIMS.Utl_Redirect(Me, objconn, url1)
                End If

            Case cst_str_Del '"Del" '刪除
                Try
                    Dim strSql As String = ""
                    strSql &= " SELECT a.ClassName ,a.CyclType ,a.RID ,a.PLANID ,a.COMIDNO ,a.SEQNO ,b.OrgID ,b.OrgName ,d.PlanName" & vbCrLf
                    strSql &= " FROM (SELECT CLASSNAME ,CYCLTYPE ,RID ,PLANID ,COMIDNO ,SEQNO FROM PLAN_PLANINFO WHERE " & e.CommandArgument & ") a" & vbCrLf
                    strSql &= " LEFT JOIN VIEW_COSTITEM vc1 ON a.PlanID=vc1.PlanID and a.ComIDNO=vc1.ComIDNO and a.SeqNo=vc1.SeqNo" & vbCrLf
                    strSql &= " JOIN Org_OrgInfo b ON a.ComIDNO=b.ComIDNO" & vbCrLf
                    strSql &= " JOIN ID_Plan c ON a.PlanID=c.PlanID" & vbCrLf
                    strSql &= " JOIN Key_Plan d ON c.TPlanID=d.TPlanID" & vbCrLf
                    Dim dt As DataTable = DbAccess.GetDataTable(strSql, objconn)
                    If dt.Rows.Count <> 1 Then
                        '若不等於1不提供刪除(異常!!)
                        Common.MessageBox(Me, "刪除失敗，請重新檢查刪除資料!!")
                        Exit Sub
                    End If
                    Dim drP1 As DataRow = dt.Rows(0)
                    Dim dt3 As DataTable = TIMS.GET_ORG_BIDCASEPI_dt(objconn, drP1("PLANID"), drP1("COMIDNO"), drP1("SEQNO"))
                    If dt3 IsNot Nothing AndAlso dt3.Rows.Count > 0 Then
                        Common.MessageBox(Me, "刪除失敗(不可刪除)，已送線上申請!")
                        Return ' Exit Sub
                    End If

                    Dim DelNote As String = String.Concat("刪除[", drP1("PlanName"), "]-[", drP1("OrgName"), "]-[", drP1("ClassName"), "]-[", drP1("CyclType"), "]")
                    TIMS.InsertDelLog(sm.UserInfo.UserID, rqMID, sm.UserInfo.DistID, DelNote, drP1("OrgID"), drP1("RID"), drP1("PlanID"), drP1("ComIDNO"), drP1("SeqNO"))
                    strSql = " DELETE PLAN_PLANINFO WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_CostItem WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_Revise WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_VerRecord WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_OnClass WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_TrainDesc WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_PlanInfo2 WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_Teacher WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_Teacher2 WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_PayHourData WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    strSql = " DELETE Plan_VerReport WHERE " & e.CommandArgument : DbAccess.ExecuteNonQuery(strSql, objconn)
                    Common.MessageBox(Me, "刪除成功")
                Catch ex As Exception
                    Common.MessageBox(Me, "刪除失敗，請重新檢查刪除資料!!")
                    Common.MessageBox(Me, ex.ToString)
                    Exit Sub
                End Try
                Call SearchData1()
            Case cst_str_Print'"Print" '列印 '列印訓練計畫

            Case cst_str_Def'"Def" '列印 經費明細表

            Case cst_str_Send '"Send" '送出 'RESULTBUTTON = NULL /正式送出
                Dim sCmdArg As String = e.CommandArgument
                If sCmdArg = "" Then Return
                Dim vPLANID As String = TIMS.GetMyValue(sCmdArg, "PLANID")
                Dim vCOMIDNO As String = TIMS.GetMyValue(sCmdArg, "COMIDNO")
                Dim vSEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")
                Dim vDISTID As String = TIMS.GetMyValue(sCmdArg, "DISTID")
                Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
                Dim vAPPSTAGE As String = TIMS.GetMyValue(sCmdArg, "APPSTAGE")

                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If sm.UserInfo.LID <> 0 Then
                        '申請階段管理-受理期間設定 APPLISTAGE
                        Dim aParms As New Hashtable From {{"YEARS", vYEARS}, {"APPSTAGE", vAPPSTAGE}}
                        '開放受理之申請階段／PLAN_APPSTAGE
                        Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE01(objconn, aParms)
                        '檢核查詢 '開放受理之申請階段／PLAN_APPSTAGE
                        If Not fg_can_applistage Then
                            Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                            Return
                        End If
                    End If
                End If

                Dim hpmsU As New Hashtable From {{"PLANID", vPLANID}, {"COMIDNO", vCOMIDNO}, {"SEQNO", vSEQNO}}
                Dim strSql As String = "SELECT ISAPPRPAPER FROM PLAN_VERREPORT WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                Dim drPV As DataRow = DbAccess.GetOneRow(strSql, objconn, hpmsU)
                If drPV Is Nothing OrElse Convert.ToString(drPV("ISAPPRPAPER")) <> "Y" Then
                    Common.MessageBox(Me, "送出資料有誤，請確認資料(正式儲存)!!")
                    Return
                End If
                strSql = "SELECT ISAPPRPAPER FROM PLAN_PLANINFO WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                Dim drPP As DataRow = DbAccess.GetOneRow(strSql, objconn, hpmsU)
                If drPP Is Nothing OrElse Convert.ToString(drPP("ISAPPRPAPER")) <> "Y" Then
                    Common.MessageBox(Me, "送出資料有誤，請確認資料(正式儲存)!")
                    Return
                End If

                '"Send" '送出 'RESULTBUTTON = NULL /正式送出
                Dim hpmsUPP As New Hashtable From {{"PLANID", vPLANID}, {"COMIDNO", vCOMIDNO}, {"SEQNO", vSEQNO}, {"MODIFYACCT", sm.UserInfo.UserID}}
                Dim strSqlUPP As String = "UPDATE PLAN_PLANINFO SET RESULTBUTTON=NULL,MODIFYDATE=GETDATE(),MODIFYACCT=@MODIFYACCT WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                DbAccess.ExecuteNonQuery(strSqlUPP, objconn, hpmsUPP)
                Dim hpmsUPV As New Hashtable From {{"PLANID", vPLANID}, {"COMIDNO", vCOMIDNO}, {"SEQNO", vSEQNO}, {"MODIFYACCT", sm.UserInfo.UserID}}
                Dim strSqlUPV As String = "UPDATE PLAN_VERREPORT SET MODIFYDATE=GETDATE(),MODIFYACCT=@MODIFYACCT WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                DbAccess.ExecuteNonQuery(strSqlUPV, objconn, hpmsUPV)
                Dim sMemo As String = String.Concat("&動作=計畫送出", "&PLANID=", vPLANID, "&COMIDNO=", vCOMIDNO, "&SEQNO=", vSEQNO)
                Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm修改, TIMS.cst_wmdip0, "", sMemo, objconn)

                Dim hPMSO2 As New Hashtable From {{"COMIDNO", vCOMIDNO}, {"TPLANID", sm.UserInfo.TPlanID}, {"DISTID", vDISTID}, {"YEARS", vYEARS}, {"APPSTAGE", vAPPSTAGE}}
                Dim iOSID2 As Integer = TIMS.GET_ORG_SCORING2_OSID2(hPMSO2, objconn)
                If iOSID2 > 0 Then '(大於0表示有資料，才修改)
                    Dim hpmsU2 As New Hashtable From {{"OSID2", iOSID2}, {"PLANID", vPLANID}, {"COMIDNO", vCOMIDNO}, {"SEQNO", vSEQNO}}
                    Dim strSqlU2 As String = "UPDATE PLAN_PLANINFO SET OSID2=@OSID2 WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                    TIMS.ExecuteNonQuery(strSqlU2, objconn, hpmsU2)
                    RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                    If RIDValue.Value <> "" Then
                        Dim hpmsU3 As New Hashtable From {{"OSID2", iOSID2}, {"PLANID", vPLANID}, {"COMIDNO", vCOMIDNO}, {"RID", RIDValue.Value}}
                        Dim strSqlU3 As String = "UPDATE PLAN_PLANINFO SET OSID2=@OSID2 WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND RID=@RID AND OSID2 IS NULL"
                        TIMS.ExecuteNonQuery(strSqlU3, objconn, hpmsU3)
                    End If
                Else
                    '(查無資料清空異常)
                    Dim hpmsU3 As New Hashtable From {{"PLANID", vPLANID}, {"COMIDNO", vCOMIDNO}, {"SEQNO", vSEQNO}}
                    Dim strSqlU3 As String = "UPDATE PLAN_PLANINFO SET OSID2=NULL WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND OSID2 IS NOT NULL"
                    TIMS.ExecuteNonQuery(strSqlU3, objconn, hpmsU3)
                End If

                Call SearchData1()

            Case cst_str_Return '"Return" '還原 Y/R
                Dim sCmdArg As String = e.CommandArgument
                If sCmdArg = "" Then Return
                Dim vPLANID As String = TIMS.GetMyValue(sCmdArg, "PLANID")
                Dim vCOMIDNO As String = TIMS.GetMyValue(sCmdArg, "COMIDNO")
                Dim vSEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")
                Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
                Dim vAPPSTAGE As String = TIMS.GetMyValue(sCmdArg, "APPSTAGE")

                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If sm.UserInfo.LID <> 0 Then
                        '申請階段管理-受理期間設定 APPLISTAGE
                        Dim aParms As New Hashtable
                        aParms.Add("YEARS", vYEARS)
                        aParms.Add("APPSTAGE", vAPPSTAGE)
                        '開放受理之申請階段／PLAN_APPSTAGE
                        Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE01(objconn, aParms)
                        '檢核查詢 '開放受理之申請階段／PLAN_APPSTAGE
                        If Not fg_can_applistage Then
                            Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                            Return
                        End If
                    End If
                End If

                Dim hpmsU As New Hashtable From {{"PLANID", vPLANID}, {"COMIDNO", vCOMIDNO}, {"SEQNO", vSEQNO}}
                Dim strSql As String = "SELECT ISAPPRPAPER FROM PLAN_VERREPORT WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                Dim drPV As DataRow = DbAccess.GetOneRow(strSql, objconn, hpmsU)
                '可修改(未送出)
                Dim strSqlU As String = ""
                If drPV IsNot Nothing Then
                    strSqlU = "UPDATE PLAN_VERREPORT SET ISAPPRPAPER='N' WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                    DbAccess.ExecuteNonQuery(strSqlU, objconn, hpmsU)
                End If
                strSqlU = "UPDATE PLAN_PLANINFO SET RESULTBUTTON='Y' WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                DbAccess.ExecuteNonQuery(strSqlU, objconn, hpmsU)
                Dim sMemo As String = String.Concat("&動作=計畫還原", "&PLANID=", vPLANID, "&COMIDNO=", vCOMIDNO, "&SEQNO=", vSEQNO)
                Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm修改, TIMS.cst_wmdip0, "", sMemo, objconn)
                Call SearchData1()

            Case "btnEdit" '修改(草稿)(非產投)
                Call KeepSearchStr()
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '產投(企訓專用)
                    Dim url1 As String = String.Concat(GET_URL1(e.CommandArgument, rqMID), "&", e.CommandArgument)
                    Call TIMS.Utl_Redirect(Me, objconn, url1)
                Else
                    'TIMS
                    Dim url1 As String = String.Concat("../03/TC_03_001.aspx?ID=", rqMID, "&", e.CommandArgument)
                    Call TIMS.Utl_Redirect(Me, objconn, url1)
                End If

            Case "btnDel" '刪除(草稿)(非產投)
                Dim strSql As String = ""
                strSql &= " SELECT a.TPLANID,a.PLANID,a.COMIDNO,a.SEQNO,a.CLASSNAME,a.CYCLTYPE,a.RID,b.OrgID,b.OrgName,d.PlanName" & vbCrLf
                strSql &= " FROM (SELECT TPLANID,PLANID,COMIDNO,SEQNO,CLASSNAME,CYCLTYPE,RID FROM PLAN_PLANINFO WHERE " & e.CommandArgument & ") a" & vbCrLf
                strSql &= " JOIN Org_OrgInfo b ON a.ComIDNO = b.ComIDNO" & vbCrLf
                strSql &= " JOIN ID_Plan c ON a.PlanID = c.PlanID" & vbCrLf
                strSql &= " JOIN Key_Plan d ON a.TPlanID = d.TPlanID" & vbCrLf
                Dim dt As DataTable = DbAccess.GetDataTable(strSql, objconn)
                If dt.Rows.Count <> 1 Then
                    '若不等於1不提供刪除(異常!!)
                    Common.MessageBox(Me, "刪除失敗，請重新檢查刪除資料!!")
                    Exit Sub
                End If
                Dim drP1 As DataRow = dt.Rows(0)
                Dim dt3 As DataTable = TIMS.GET_ORG_BIDCASEPI_dt(objconn, drP1("PLANID"), drP1("COMIDNO"), drP1("SEQNO"))
                If dt3 IsNot Nothing AndAlso dt3.Rows.Count > 0 Then
                    Common.MessageBox(Me, "刪除失敗(不可刪除)，已送線上申請!!")
                    Return ' Exit Sub
                End If

                Dim DelNote As String = String.Concat("刪除[", drP1("PlanName"), "]-[", drP1("OrgName"), "]-[", drP1("ClassName"), "]-[", drP1("CyclType"), "]")
                TIMS.InsertDelLog(sm.UserInfo.UserID, rqMID, sm.UserInfo.DistID, DelNote, drP1("OrgID"), drP1("RID"), drP1("PlanID"), drP1("ComIDNO"), drP1("SeqNO"))

                strSql = " DELETE PLAN_PLANINFO WHERE " & e.CommandArgument
                DbAccess.ExecuteNonQuery(strSql, objconn)
                strSql = " DELETE PLAN_COSTITEM WHERE " & e.CommandArgument
                DbAccess.ExecuteNonQuery(strSql, objconn)
                strSql = " DELETE PLAN_TRAINDESC WHERE " & e.CommandArgument
                DbAccess.ExecuteNonQuery(strSql, objconn)
                Common.MessageBox(Me, "刪除成功")
                Call SearchData1()

        End Select
    End Sub

    Private Sub dtPlan_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dtPlan.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Hid_PCS As HiddenField = e.Item.FindControl("Hid_PCS") '選取id'concat(p1.planid,'x',p1.ComIDNO,'x',p1.SeqNO
                Dim chkItem As HtmlInputCheckBox = e.Item.FindControl("chkItem") '選取
                '正式  '(產業人才投資方案/TIMS方案)
                Dim lbtUpdate As LinkButton = e.Item.FindControl("lbtUpdate") '修改'update/VIEW
                Dim lbtDel As LinkButton = e.Item.FindControl("lbtDel") '刪除'Del
                Dim lbtPrint As LinkButton = e.Item.FindControl("lbtPrint") '列印 'print
                Dim lbtDef As LinkButton = e.Item.FindControl("lbtDef") '經費明細 'Def
                'Dim lbtCancel As LinkButton = e.Item.FindControl("lbtCancel") '取消審核'Cancel
                Dim lbtSend As LinkButton = e.Item.FindControl("lbtSend") '送出'Send
                Dim lbtReturn As LinkButton = e.Item.FindControl("lbtReturn") '還原'Return
                Dim lbtShelf As LinkButton = e.Item.FindControl("lbtShelf") '轉班上架'Shelf

                '草稿  '(產業人才投資方案/TIMS方案)
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit") '修改'btnEdit
                Dim lbtDel1 As LinkButton = e.Item.FindControl("lbtDel1") '刪除'btnDel
                '序號
                e.Item.Cells(Cst_index).Text = TIMS.Get_DGSeqNo(sender, e) 'e.Item.ItemIndex + 1 + dtPlan.PageSize * dtPlan.CurrentPageIndex

                Dim fg_can_applistage_PTYPE01 As Boolean = (Convert.ToString(drv("CAN_APPLISTAGE_PTYPE01")) = "Y")
                Dim str_cnValue As String = Convert.ToString(drv("CLASSNAME2"))
                e.Item.Cells(Cst_ClassName).ToolTip = str_cnValue
                e.Item.Cells(Cst_ClassName).Text = str_cnValue

                Dim s_CMDARG As String = String.Concat("PLANID=", drv("PLANID"), "&COMIDNO=", drv("COMIDNO"), "&SEQNO=", drv("SEQNO"), "&YEARS=", drv("YEARS"), "&DISTID=", drv("DISTID"), "&APPSTAGE=", drv("APPSTAGE"))
                '修改／檢視-產投
                lbtUpdate.CommandArgument = s_CMDARG 'String.Concat("PLANID=", drv("PLANID"), "&COMIDNO=", drv("COMIDNO"), "&SEQNO=", drv("SEQNO"), "&YEARS=", drv("YEARS"), "&DISTID=", drv("DISTID"), "&APPSTAGE=", drv("APPSTAGE"))
                '送出-產投
                lbtSend.CommandArgument = s_CMDARG 'String.Concat("PLANID=", drv("PLANID"), "&COMIDNO=", drv("COMIDNO"), "&SEQNO=", drv("SEQNO"), "&YEARS=", drv("YEARS"), "&DISTID=", drv("DISTID"), "&APPSTAGE=", drv("APPSTAGE"))
                '還原-產投
                lbtReturn.CommandArgument = s_CMDARG 'String.Concat("PLANID=", drv("PLANID"), "&COMIDNO=", drv("COMIDNO"), "&SEQNO=", drv("SEQNO"), "&YEARS=", drv("YEARS"), "&DISTID=", drv("DISTID"), "&APPSTAGE=", drv("APPSTAGE"))

                '檢查行政管理費
                Dim AdmName1 As String = "" '預設為空白文字。
                Dim AdmName2 As String = "0" '預設為數字0
                Dim TotalCost As Integer = 0

                '行政管理費文字。
                If AdmName1 <> "" Then
                    AdmName1 = "行政管理費：" & AdmName1 & "＊" & drv("AdmPercent").ToString & "％ ＝ " & Math.Round(TotalCost * drv("AdmPercent") / 100)
                    AdmName1 = HttpUtility.UrlEncode(AdmName1)
                    AdmName2 = Math.Round(TotalCost * drv("AdmPercent") / 100)
                End If

                '依選擇的計費方式
                '列印訓練計畫but2
                lbtPrint.CommandArgument = "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNO=" & drv("SeqNO") & Convert.ToString(drv("CostMode"))
                If sm.UserInfo.Years = "2006" Then
                    Select Case Convert.ToString(drv("CostMode"))
                        Case "1" '自辦
                            lbtPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, "TC_02_001_1", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&CostMode=" & Convert.ToString(drv("CostMode")))
                        Case "2" '每人每時
                            lbtPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, "TC_02_001_2", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&CostMode=" & Convert.ToString(drv("CostMode")))
                        Case "3" '每人輔助
                            lbtPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, "TC_02_001_3", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&CostMode=" & Convert.ToString(drv("CostMode")))
                        Case "4" '個人單價
                            lbtPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, "TC_02_001_4", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&CostMode=" & Convert.ToString(drv("CostMode")))
                        Case Else
                            lbtPrint.Attributes("onclick") = "alert('目前此筆計畫尚無填寫訓練費用!!\n不能列印經費明細表!');return false;"
                    End Select
                Else
                    Select Case Convert.ToString(drv("CostMode"))
                        Case "1" '自辦
                            'OJT-21061501：<系統> 自辦在職、區域 - 班級申請：隱藏【企業負擔金額】欄位 欄位為接受企業委託計畫才會使用
                            Dim s_prt_PlanInfo As String = If(TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1, "PlanInfo", "PlanInfo70")
                            lbtPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, s_prt_PlanInfo, "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&AdmName1=" & AdmName1 & "&AdmName2=" & AdmName2 & "&CostMode=" & Convert.ToString(drv("CostMode")))
                        Case "2" '每人每時
                            lbtPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, "PlanInfo_2", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&CostMode=" & Convert.ToString(drv("CostMode")))
                        Case "3" '每人輔助
                            lbtPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, "PlanInfo_3", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&CostMode=" & Convert.ToString(drv("CostMode")))
                        Case "4" '個人單價
                            lbtPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, "PlanInfo_4", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&AdmName1=" & AdmName1 & "&AdmName2=" & AdmName2 & "&CostMode=" & Convert.ToString(drv("CostMode")))
                        Case Else
                            lbtPrint.Attributes("onclick") = "alert('目前此筆計畫尚無填寫訓練費用!!\n不能列印經費明細表!');return false;"
                    End Select
                End If

                '列印經費明細表lbtDef
                lbtDef.CommandArgument = "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNO=" & drv("SeqNO")
                Select Case Convert.ToString(drv("CostMode"))
                    Case "1" '自辦
                        lbtDef.Attributes("onclick") = ReportQuery.ReportScript(Me, "Def_Rpt_1", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&AdmName1=" & AdmName1 & "&AdmName2=" & AdmName2)
                    Case "2" '每人每時
                        lbtDef.Attributes("onclick") = ReportQuery.ReportScript(Me, "Def_Rpt_2", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "")
                    Case "3" '每人輔助
                        lbtDef.Attributes("onclick") = ReportQuery.ReportScript(Me, "Def_Rpt_3", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "")
                    Case "4" '個人單價
                        lbtDef.Attributes("onclick") = ReportQuery.ReportScript(Me, "Def_Rpt_4", "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNO") & "&AdmName1=" & AdmName1 & "&AdmName2=" & AdmName2)
                    Case Else
                        lbtDef.Attributes("onclick") = "alert('目前此筆計畫尚無填寫訓練費用!!\n不能列印經費明細表!');return false;"
                End Select

                '產投
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    lbtPrint.Visible = False
                    lbtDef.Visible = False
                End If

                lbtDel.Attributes("onclick") = TIMS.cst_confirm_delmsg2
                Dim flag_classInfo_exists As Boolean = True 'CLASSINFO_EXISTS '已轉班
                If Convert.ToString(drv("OCID")) = "" Then flag_classInfo_exists = False '未轉班
                '已轉班-刪除-不顯示
                If flag_classInfo_exists Then lbtDel.Visible = False '已轉班class_classinfo不可刪除

                '只有正式資料(正式儲存過)且未轉班，才會顯示此按鈕。已轉班反灰
                Hid_PCS.Value = If(drv("IsApprPaper").ToString() = "Y", drv("PCS").ToString(), "")
                lbtShelf.Visible = If(flag_EnterShelf, True, False) '選取-'在職-批次轉班上架
                lbtShelf.Enabled = True '未轉班
                If flag_classInfo_exists Then lbtShelf.Enabled = False  '已轉班
                chkItem.Disabled = False '未轉班
                If flag_classInfo_exists Then chkItem.Disabled = True '已轉班
                'If Convert.ToString(drv("OCID")) <> "" Then                '    If (Val(drv("OCID")) Mod 2 = 1) Then chkItem.Disabled = False '未轉班
                If Not lbtShelf.Enabled Then TIMS.Tooltip(lbtShelf, "已轉班", True)
                If chkItem.Disabled Then TIMS.Tooltip(chkItem, "已轉班", True)
                If lbtShelf.Enabled Then
                    Dim sCmdArg As String = ""
                    TIMS.SetMyValue(sCmdArg, "PCS", drv("PCS").ToString())
                    TIMS.SetMyValue(sCmdArg, "CJOB_UNKEY", drv("CJOB_UNKEY").ToString())
                    TIMS.SetMyValue(sCmdArg, "PlanID", drv("PlanID").ToString())
                    TIMS.SetMyValue(sCmdArg, "COMIDNO", drv("COMIDNO").ToString())
                    TIMS.SetMyValue(sCmdArg, "SEQNO", drv("SEQNO").ToString())
                    TIMS.SetMyValue(sCmdArg, "TPlanID", drv("TPlanID").ToString())
                    TIMS.SetMyValue(sCmdArg, "RID", drv("RID").ToString())
                    'TIMS.LOG.Debug(String.Format("#sCmdArg:{0}", sCmdArg))
                    lbtShelf.CommandArgument = sCmdArg
                End If

                '(NOT)已轉班
                If Not flag_classInfo_exists Then
                    '未轉班-控制修改／刪除--顯示
                    '尚未轉班(TIMS : 有權限刪除者可刪除/產學訓:依狀況 不可刪除)
                    lbtDel.CommandArgument = String.Concat("PlanID='", drv("PlanID"), "' and ComIDNO='", drv("ComIDNO"), "' and SeqNO=", drv("SeqNO"))

                    If Convert.ToString(drv("PlanKind")) = "2" Then   '1.自辦 '2.委外
                        '開班送審(TIMS/產業人才投資方案, 都有此流程) '產投
                        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                            lbtUpdate.Visible = True '可修改(顯示)
                        Else
                            '非產投
                            If drv("AppliedResult").ToString = "N" Then
                                lbtUpdate.Enabled = False '不可修改
                                lbtUpdate.Visible = False '___by:20180824
                                lbtDel.Enabled = True '可刪除
                                lbtDel.ToolTip = "審核不通過，可刪除"
                            Else
                                lbtUpdate.Enabled = True '可修改
                                lbtUpdate.Visible = True '___by:20180824
                                lbtDel.Enabled = False '不可刪除
                                lbtDel.ToolTip = "審核通過／審核中，不可刪除"
                            End If
                        End If
                    End If
                End If

                lbtEdit.CommandArgument = String.Concat("PlanID=", drv("PlanID"), "&ComIDNO=", drv("ComIDNO"), "&SeqNO=", drv("SeqNO"), "&YEARS=", drv("YEARS"))
                lbtDel1.CommandArgument = String.Concat("PlanID='", drv("PlanID"), "' and ComIDNO='", drv("ComIDNO"), "' and SeqNO=", drv("SeqNO"))
                lbtDel1.Attributes("onclick") = "return confirm('確定要刪除這份草稿?');"

                '產投
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Dim s_FirResult As String = Convert.ToString(drv("FirResult")) '.ToString
                    lbtUpdate.Visible = True '可修改(顯示)
                    lbtDel.Visible = False '不可刪除(不顯示)
                    lbtDel.ToolTip = "產業人才投資方案初審通過，不通過及退件修正，不可刪除"

                    '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
                    Select Case s_FirResult
                        Case "Y", "N", "R"
                        Case Else
                            'e.Item.Cells(Cst_AppliedResult).Text = "尚未審核<br>"
                            Select Case sm.UserInfo.LID
                                Case 0, 1
                                    lbtDel.Visible = True '可刪除
                                    lbtDel.ToolTip = "產業人才投資方案 尚未審核，分署可刪除"
                                Case Else
                                    lbtDel.Visible = False '不可刪除
                                    lbtDel.ToolTip = "產業人才投資方案 尚未審核，委訓單位不可刪除"
                            End Select
                    End Select
                    If lbtDel.Visible AndAlso yearlist.SelectedValue <> "" AndAlso yearlist.SelectedValue <> sm.UserInfo.Years Then
                        lbtDel.Enabled = False
                        TIMS.Tooltip(lbtDel, "查詢年度與登入年度不同!")
                    End If
                End If
                Dim strMsg As String = Get_AppliedResultTxt(sm.UserInfo.TPlanID, Convert.ToString(drv("AppliedResult")), Convert.ToString(drv("RESULTBUTTON")))
                'TIMS.Tooltip(e.Item.Cells(Cst_AppliedResult), strMsg)
                e.Item.Cells(Cst_AppliedResult).Text = strMsg

                '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
                'PlanKind '1:自辦(內訓)'2:委外
                If Convert.ToString(drv("PlanKind")) = "2" Then
                    '通過時，0:署(局),1:分署(中心) 可修改 2:委訓可查詢 
                    If Convert.ToString(drv("AppliedResult")) = "Y" Then
                        'OJT-21041501：<系統> 產投 - 班級查詢：功能權限調整 BY AMU 20210611
                        '目前課程審核通過後， 分署要查詢課程內容須點選「修改」，
                        '為避免分署修改訓練單位已審核通過之課程資料(日前發生多次已審核通過的資料被分署修改， 與送審文件不同)，
                        '請將「修改」功能權限調整為
                        '審核通過的班級： 分署、訓練單位僅能使用「查詢」按鈕，僅能署可使用「修改」。
                        lbtUpdate.Text = If(sm.UserInfo.LID > 0, cst_lbtUpdate_Txt_查詢, cst_lbtUpdate_Txt_修改)
                        lbtUpdate.CommandName = If(sm.UserInfo.LID > 0, cst_str_view, cst_str_update)
                    ElseIf drv("FirResult").ToString = "Y" Then
                        '若不為R退件修正或是O審核後修正(即 N不通過)、0:署(局),1:分署(中心) 可修改 2:委訓可查詢 
                        If drv("AppliedResult").ToString <> "O" And drv("AppliedResult").ToString <> "R" Then
                            lbtUpdate.Text = If(sm.UserInfo.LID > 1, cst_lbtUpdate_Txt_查詢, cst_lbtUpdate_Txt_修改)
                            lbtUpdate.CommandName = If(sm.UserInfo.LID > 1, cst_str_view, cst_str_update)
                        End If
                    End If
                Else
                    '1:.自辦(內訓)任何情況下都可修改
                    lbtUpdate.CommandName = cst_str_update '"update"
                End If

                If isBlack.Value = "Y" Then
                    If lbtUpdate.CommandName = cst_str_update Then '"update" Then
                        lbtUpdate.Text = cst_lbtUpdate_Txt_查詢
                        lbtUpdate.CommandName = cst_str_view '"view"
                    End If
                End If
                If yearlist.SelectedValue <> "" AndAlso yearlist.SelectedValue <> sm.UserInfo.Years Then
                    If lbtUpdate.CommandName = cst_str_update Then '"update" Then
                        lbtUpdate.Text = cst_lbtUpdate_Txt_查詢
                        lbtUpdate.CommandName = cst_str_view '"view"
                    End If
                End If

                '再判斷是否為分署(中心)特許依此計畫查詢有關權限
                If lbtUpdate.CommandName = cst_str_update Then '"update" Then
                    Select Case sm.UserInfo.LID
                        Case 0 '署(局)的帳號功能
                        Case 1 '分署(中心)的帳號功能 / '委訓
                            '分署(中心)
                            lbtUpdate.Text = If(drv("CanEdit").ToString = "0", cst_lbtUpdate_Txt_查詢, cst_lbtUpdate_Txt_修改)
                            lbtUpdate.CommandName = If(drv("CanEdit").ToString = "0", cst_str_view, cst_str_update)
                        Case Else '委訓

                    End Select
                End If

                Dim strPlanSeq As String = String.Concat("Plan Seq:", drv("Seq")) '.ToString
                'strPlanSeq = "Plan Seq:" & drv("Seq").ToString
                Select Case Convert.ToString(drv("TransFlag"))'.ToString
                    Case "Y"
                        e.Item.Cells(Cst_TransFlag).Text = "是"
                        Dim s_tip1 As String = strPlanSeq
                        If drv("OCID").ToString() = "" Then
                            e.Item.Cells(Cst_TransFlag).Text &= "<br><font color='red'>(異常)</font>"
                            s_tip1 = "無轉班後課程代碼流水號"
                        End If
                        TIMS.Tooltip(e.Item.Cells(Cst_TransFlag), s_tip1)
                    Case "N"
                        e.Item.Cells(Cst_TransFlag).Text = "否"
                        TIMS.Tooltip(e.Item.Cells(Cst_TransFlag), strPlanSeq)
                    Case Else
                        e.Item.Cells(Cst_TransFlag).Text = "否"
                        TIMS.Tooltip(e.Item.Cells(Cst_TransFlag), strPlanSeq)
                End Select

                'Dim ppWhereStr1 As String = String.Concat("PlanID='", drv("PlanID"), "' and ComIDNO='", drv("ComIDNO"), "' and SeqNO=", drv("SeqNO"))
                lbtSend.Visible = False  '送出-產投
                lbtReturn.Visible = False '還原-產投

                '產投/非產投判斷
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso lbtUpdate.Text = cst_lbtUpdate_Txt_修改 Then
                    'If Convert.ToString(drv("RESULTBUTTON")) = "Y" Then '可修改送出
                    '    '還未送出
                    '    Dim strMsg As String = " "
                    '    e.Item.Cells(Cst_AppliedResult).Text = strMsg '"班級審核中"
                    '    TIMS.Tooltip(e.Item.Cells(Cst_AppliedResult), strMsg)
                    'End If
                    Dim strMsg2 As String = ""
                    Select Case sm.UserInfo.LID
                        Case 0
                            Select Case Convert.ToString(drv("AppliedResult"))
                                Case "Y", "N", "R"
                                    Select Case Convert.ToString(drv("RESULTBUTTON"))
                                        Case TIMS.cst_ResultButton_尚未送出_待送審 'Y
                                            strMsg2 = "分署、委訓已審核，待送出"
                                            TIMS.Tooltip(lbtUpdate, strMsg2)
                                    End Select
                            End Select

                        Case 1, 2 '"1", "2" '分署(中心)、委訓
                            Select Case drv("AppliedResult").ToString '審核狀況
                                Case "Y", "N", "R"
                                    Select Case Convert.ToString(drv("RESULTBUTTON"))
                                        Case TIMS.cst_ResultButton_尚未送出_待送審 'Y
                                            lbtSend.Visible = True
                                            lbtSend.Enabled = True '可送出
                                            strMsg2 = "分署、委訓已審核，待送出"
                                            TIMS.Tooltip(lbtSend, strMsg2)
                                    End Select

                                Case Else '審核狀況-未審
                                    lbtSend.Visible = True '送出可顯示
                                    lbtSend.Enabled = False '不可送出
                                    Select Case Convert.ToString(drv("RESULTBUTTON"))
                                        Case TIMS.cst_ResultButton_尚未送出_待送審 'Y
                                            lbtSend.Enabled = True '可送出
                                            'TIMS.Tooltip(btnSend, "中心、委訓尚未審核，待送出")
                                            strMsg2 = "分署、委訓尚未審核，待送出"
                                            TIMS.Tooltip(lbtSend, strMsg2)
                                            lbtUpdate.Enabled = True '可修改
                                            lbtUpdate.Visible = True '___by:20180824

                                        Case TIMS.cst_ResultButton_尚未送出_未送出 'R
                                            lbtSend.Enabled = False '不可送出
                                            'TIMS.Tooltip(btnSend, "中心、委訓尚未審核，待送出")
                                            strMsg2 = "分署、委訓尚未審核，不送出"
                                            TIMS.Tooltip(lbtSend, strMsg2)
                                            lbtUpdate.Enabled = True '可修改
                                            lbtUpdate.Visible = True '___by:20180824

                                        Case Else
                                            lbtSend.Enabled = False '不可送出
                                            'TIMS.Tooltip(btnSend, "中心、委訓尚未審核，已送出")
                                            strMsg2 = "分署、委訓尚未審核，已送出"
                                            TIMS.Tooltip(lbtSend, strMsg2)
                                            lbtUpdate.Enabled = False '不可修改
                                            If sm.UserInfo.LID = 1 Then '分署(中心)
                                                lbtUpdate.Enabled = True '可修改
                                                lbtUpdate.Visible = True '___by:20180824
                                            End If
                                    End Select

                                    '未提供申請日，還可以再修改
                                    If Convert.ToString(drv("AppliedDate")) = "" Then
                                        lbtUpdate.Enabled = True '可修改
                                        lbtUpdate.Visible = True '___by:20180824
                                    End If

                            End Select
                    End Select
                    'strMsg/strMsg2
                    TIMS.Tooltip(e.Item.Cells(Cst_AppliedResult), If(strMsg2 <> "", strMsg2, strMsg))

                    'If flag_amu_test Then
                    '    If Not lbtUpdate.Enabled Then
                    '        lbtUpdate.Enabled = True '可修改
                    '        TIMS.Tooltip(lbtUpdate, "測試環境,修改(Enabled)!")
                    '    End If
                    '    If Not lbtUpdate.Visible Then
                    '        lbtUpdate.Visible = True '可修改
                    '        TIMS.Tooltip(lbtUpdate, "測試環境,修改(Visible)!")
                    '    End If
                    'End If

                    If sm.UserInfo.LID = 1 Then '分署(中心)可用還原功能
                        Select Case drv("AppliedResult").ToString
                            Case "Y", "N", "R"
                            Case Else
                                lbtReturn.Visible = True '(還原)顯示
                                lbtReturn.Enabled = False '(還原)停用
                                Select Case Convert.ToString(drv("RESULTBUTTON"))
                                    Case TIMS.cst_ResultButton_尚未送出_待送審 'Y
                                        lbtReturn.Enabled = False '暫無法還原
                                        TIMS.Tooltip(lbtReturn, "分署尚未審核，未送出不需還原!")
                                    Case TIMS.cst_ResultButton_尚未送出_未送出 'R
                                        lbtReturn.Enabled = False '暫無法還原
                                        TIMS.Tooltip(lbtReturn, "分署尚未審核，未送出不可審核不需還原")
                                    Case Else
                                        lbtReturn.Enabled = True '可還原 '不可修改已送出
                                        TIMS.Tooltip(lbtReturn, "分署尚未審核，已送出可還原")
                                End Select
                        End Select
                    End If
                End If

                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If Not fg_can_applistage_PTYPE01 AndAlso sm.UserInfo.LID = 1 Then
                        If lbtUpdate.Enabled AndAlso lbtUpdate.CommandName = cst_str_update Then
                            'lbtUpdate.Enabled = False 'TIMS.Tooltip(lbtUpdate, "已超過該年度、階段之班級受理期間，分署不可修改")
                            lbtUpdate.Text = cst_lbtUpdate_Txt_查詢
                            lbtUpdate.CommandName = cst_str_view
                            TIMS.Tooltip(lbtUpdate, "已超過該年度、階段之班級受理期間，分署不可修改,僅可查詢")
                        End If
                        If lbtSend.Enabled Then
                            lbtSend.Enabled = False
                            TIMS.Tooltip(lbtSend, "已超過該年度、階段之班級受理期間，分署不可送出")
                        End If
                        If lbtReturn.Enabled Then
                            lbtReturn.Enabled = False
                            TIMS.Tooltip(lbtReturn, "已超過該年度、階段之班級受理期間，分署不可還原")
                        End If
                    End If
                End If
                '-- edit，by:20181001
                'e.Item.Cells(Cst_PlanYear).Text = Convert.ToString(drv("PlanYear"))        '計畫年度
                'e.Item.Cells(Cst_AppliedDate).Text = Convert.ToString(drv("AppliedDate"))  '申請日期
                'e.Item.Cells(Cst_STDate).Text = Convert.ToString(drv("STDate"))            '訓練起日
                'e.Item.Cells(Cst_FDDate).Text = Convert.ToString(drv("FDDate"))            '訓練迄日
                e.Item.Cells(Cst_AppliedDate).Text = TIMS.Cdate3(drv("AppliedDate"))  '申請日期
                e.Item.Cells(Cst_STDate).Text = TIMS.Cdate3(drv("STDate"))            '訓練起日
                e.Item.Cells(Cst_FDDate).Text = TIMS.Cdate3(drv("FDDate"))            '訓練迄日
                If flag_ROC Then
                    'If Convert.ToString(drv("PlanYear")) <> "" Then
                    '    e.Item.Cells(Cst_PlanYear).Text = String.Format("{0:000}", CInt(drv("PlanYear")) - 1911)  '計畫年度
                    'End If
                    e.Item.Cells(Cst_AppliedDate).Text = TIMS.Cdate17(drv("AppliedDate"))  '申請日期
                    e.Item.Cells(Cst_STDate).Text = TIMS.Cdate17(drv("STDate"))            '訓練起日
                    e.Item.Cells(Cst_FDDate).Text = TIMS.Cdate17(drv("FDDate"))            '訓練迄日
                End If

            Case ListItemType.Header
                If Me.ViewState(cst_Sort) <> "" Then
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i_Cell As Integer = -1
                    Dim str_Sort As String = ViewState(cst_Sort).ToString()
                    Call ACT_ImageUrl_UD(mysort, i_Cell, str_Sort)
                    ViewState(cst_Sort) = str_Sort
                    If i_Cell <> -1 Then e.Item.Cells(i_Cell).Controls.Add(mysort)
                End If

        End Select
    End Sub

    ''' <summary> SQL 取得資料</summary>
    ''' <returns></returns>
    Function GetSchDt() As DataTable
        Dim dt As DataTable = Nothing
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)

        '取出訓練計畫名稱
        Dim relship As String = "" 'RIDValue
        Dim PMS_9 As New Hashtable From {{"RID", RIDValue.Value}}
        Dim sql99 As String = "SELECT RELSHIP,ORGID FROM AUTH_RELSHIP WHERE RID=@RID"
        Dim dr99 As DataRow = DbAccess.GetOneRow(sql99, objconn, PMS_9)
        If dr99 Is Nothing Then
            RIDValue.Value = ""
            Return dt
        End If
        relship = Convert.ToString(dr99("relship")) 'RIDValue
        Orgidvalue.Value = Convert.ToString(dr99("OrgID"))
        Dim v_IsApprPaper As String = TIMS.GetListValue(IsApprPaper) '判斷是否是選正式-狀態
        Dim v_yearlist As String = TIMS.GetListValue(yearlist) '年度 'yearlist.SelectedValue
        Dim v_AppStage As String = If(tr_AppStage_TP28.Visible, TIMS.GetListValue(AppStage), "")
        Dim v_audit As String = TIMS.GetListValue(audit)  'audit.SelectedValue '審核狀態有選值

        '建立訓練的DATATABLE PLAN_PLANINFO
        Dim StrSql As String = ""
        StrSql &= " SELECT vc1.COSTMODE,I1.PLANKIND,P1.PLANYEAR" & vbCrLf
        StrSql &= " ,CONCAT(dbo.FN_GET_ROC_YEAR(P1.PLANYEAR),dbo.FN_GET_APPSTAGE2(p1.AppStage)) PlanYearROCAG" & vbCrLf
        StrSql &= " ,K1.PLANNAME,i2.NAME ORGNAME2,I1.SEQ,O1.OrgName" & vbCrLf 'OrgName2管控單位
        StrSql &= " ,P1.STDATE,P1.FDDATE" & vbCrLf
        StrSql &= " ,P1.PLANID,P1.COMIDNO,P1.SEQNO" & vbCrLf
        StrSql &= " ,P1.TMID,p1.CJOB_UNKEY,P1.AppliedResult,P1.RESULTBUTTON" & vbCrLf
        StrSql &= " ,p1.APPSTAGE,dbo.FN_GET_APPSTAGE(p1.APPSTAGE) APPSTAGE" & vbCrLf
        StrSql &= " ,CASE WHEN K2.JobID IS NULL THEN K2.TrainName ELSE K2.JobName END TRAINNAME" & vbCrLf
        StrSql &= " ,P1.CLASSNAME,P1.ADMPERCENT,P1.CYCLTYPE" & vbCrLf
        StrSql &= " ,P1.APPLIEDDATE" & vbCrLf
        StrSql &= " ,dbo.FN_GET_CLASSCNAME(P1.CLASSNAME ,P1.CYCLTYPE) CLASSNAME2" & vbCrLf
        StrSql &= " ,p1.THOURS,p1.TNUM,p1.ADVANCE" & vbCrLf
        StrSql &= " ,P1.TRANSFLAG,P1.RID" & vbCrLf
        StrSql &= " ,P1.PSNO28 ,C1.OCID" & vbCrLf
        StrSql &= " ,I1.YEARS,I1.TPLANID,I1.DISTID" & vbCrLf
        StrSql &= " ,concat(p1.planid,'x',p1.ComIDNO,'x',p1.SeqNO) PCS" & vbCrLf
        'StrSql &= " ,r3.OrgID2 ,ISNULL(r3.OrgName2,i2.name) OrgName2" & vbCrLf 'OrgName2管控單位
        StrSql &= " ,CASE WHEN CONVERT(varchar(30), P1.PlanID)='" & sm.UserInfo.PlanID & "' THEN 1 ELSE 0 END AS CanEdit" & vbCrLf
        StrSql &= " ,P1.IsApprPaper ,P2.VerReason" & vbCrLf '未通過原因
        StrSql &= " ,dbo.FN_PLAN_APPSTAGE_PTYPE01(I1.YEARS,p1.APPSTAGE) CAN_APPLISTAGE_PTYPE01" & vbCrLf '未通過原因
        'dbo.FN_PLAN_APPSTAGE_PTYPE1('2024',2)
        '產投/非產投判斷
        StrSql &= If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, " ,P3.FirResult", " ,'X' FirResult") & vbCrLf
        '產投/非產投判斷
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    StrSql &= " ,P3.FirResult, P3.SecResult, P3.IsApprPaper" & vbCrLf '產投
        'Else
        '    StrSql &= " ,'X' FirResult, 'X' SecResult, NULL IsApprPaper" & vbCrLf '給個假資料
        'End If
        StrSql &= " FROM PLAN_PLANINFO P1 WITH(NOLOCK)" & vbCrLf
        StrSql &= " JOIN ID_PLAN I1 WITH(NOLOCK) ON P1.PlanID=I1.PlanID" & vbCrLf
        StrSql &= " JOIN ID_DISTRICT I2 WITH(NOLOCK) ON I2.DISTID=I1.DISTID" & vbCrLf
        StrSql &= " JOIN KEY_PLAN K1 WITH(NOLOCK) ON P1.TPlanID=K1.TPlanID" & vbCrLf
        StrSql &= " JOIN ORG_ORGINFO O1 WITH(NOLOCK) ON P1.ComIDNO=O1.ComIDNO" & vbCrLf
        StrSql &= " JOIN dbo.AUTH_RELSHIP rr WITH(NOLOCK) ON rr.RID=p1.RID" & vbCrLf
        StrSql &= " LEFT JOIN VIEW_COSTITEM vc1 WITH(NOLOCK) ON P1.PlanID=vc1.PlanID AND P1.ComIDNO=vc1.ComIDNO AND P1.SeqNo=vc1.SeqNo" & vbCrLf
        StrSql &= " LEFT JOIN KEY_TRAINTYPE K2 WITH(NOLOCK) ON P1.TMID=K2.TMID" & vbCrLf
        StrSql &= " LEFT JOIN SHARE_CJOB s WITH(NOLOCK) ON s.CJOB_UNKEY=P1.CJOB_UNKEY" & vbCrLf
        StrSql &= " LEFT JOIN SHARE_CJOB_REL sr WITH(NOLOCK) ON sr.UNKEY1= P1.CJOB_UNKEY" & vbCrLf
        StrSql &= " LEFT JOIN PLAN_VERRECORD P2 WITH(NOLOCK) ON P1.PlanID=P2.PlanID AND P1.ComIDNO=P2.ComIDNO AND P1.SeqNo=P2.SeqNo" & vbCrLf
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then '產投/非產投判斷
        Dim s_LEFTJOINP3 As String = " LEFT JOIN PLAN_VERREPORT P3 WITH(NOLOCK) ON P1.PlanID=P3.PlanID AND P1.ComIDNO=P3.ComIDNO AND P1.SeqNo=P3.SeqNo"
        StrSql &= If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, s_LEFTJOINP3, "") & vbCrLf
        'C1.IsSuccess 是否轉入成功
        StrSql &= " LEFT JOIN CLASS_CLASSINFO C1 WITH(NOLOCK) ON P1.PlanID=C1.PlanID AND P1.ComIDNO=C1.ComIDNO AND P1.SeqNo=C1.SeqNo AND C1.IsSuccess='Y'" & vbCrLf
        StrSql &= " WHERE 1=1" & vbCrLf

        '課程代碼
        Dim flag_can_use_OCID As Boolean = False '轉換數字後相等:true/false:異常
        s_OCID.Text = TIMS.ClearSQM(s_OCID.Text)
        If s_OCID.Text <> "" AndAlso TIMS.IsNumeric2(s_OCID.Text) Then
            Dim vs_OCID As String = Convert.ToString(CInt(Val(s_OCID.Text)))  '轉換數字後相等:true/false:異常
            If vs_OCID.Equals(s_OCID.Text) Then flag_can_use_OCID = True '轉換數字後相等:true/false:異常
        End If
        If flag_can_use_OCID Then StrSql &= " AND C1.OCID='" & s_OCID.Text & "'" & vbCrLf
        '依申請階段
        'Dim v_AppStage As String = If(tr_AppStage_TP28.Visible, TIMS.GetListValue(AppStage), "")
        If v_AppStage <> "" Then StrSql &= " AND P1.AppStage='" & v_AppStage & "'" & vbCrLf '依申請階段

        'rbl_TransFlagS '增加【轉班上架】欄位，選項：不區分、未轉班、已轉班
        Dim v_rbl_TransFlagS As String = TIMS.GetListValue(rbl_TransFlagS)
        v_rbl_TransFlagS = If(v_rbl_TransFlagS = "A", "", v_rbl_TransFlagS)
        If v_rbl_TransFlagS <> "" Then StrSql &= " AND P1.TransFlag='" & v_rbl_TransFlagS & "'" & vbCrLf '轉班上架

        If Len(sm.UserInfo.RID) = 1 Then
            Select Case sm.UserInfo.RID
                Case "A" '署(局)權限
                    If v_yearlist <> "" Then '有選年度依年度
                        StrSql &= " AND I1.Years='" & v_yearlist & "'" & vbCrLf
                    Else '沒選年度依登入年度計畫@TPlanID
                        StrSql &= " AND I1.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                    End If
                Case Else
                    '分署(中心)特許依此計畫查詢有關權限 2008-09-30 AMU
                    If v_yearlist <> "" Then '有選年度依年度
                        StrSql &= " AND I1.Years='" & v_yearlist & "'" & vbCrLf
                    Else '沒選年度依登入年度計畫@TPlanID
                        StrSql &= " AND I1.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                    End If
            End Select
        Else
            StrSql &= " AND I1.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf  '依登入年度計畫 @PlanID
        End If

        StrSql &= " AND P1.TPlanID='" & TPlanid.Value & "'" & vbCrLf
        StrSql &= " AND P1.IsApprPaper='" & v_IsApprPaper & "'" & vbCrLf

        '產投/非產投判斷
        '如果不是產學訓計畫 penny 2007/10/17
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If v_IsApprPaper = "Y" Then  '判斷是否是選正式
                If v_audit <> "" Then       '如果審核狀態有選值
                    Select Case v_audit
                        Case "Y" '如果是選已審核
                            StrSql &= " AND P1.AppliedResult IN ('Y','N')" & vbCrLf
                        Case "N" '如果是選審核中
                            StrSql &= " AND (P1.AppliedResult NOT IN ('Y','N') OR P1.AppliedResult IS NULL)" & vbCrLf
                    End Select
                End If
            End If
        End If

        '有關機構的篩選
        Select Case Convert.ToString(sm.UserInfo.LID)  '登入者
            Case "0" '署(局)
            Case "1" '分署(中心)
                StrSql &= " AND I1.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
            Case Else '補助地方政府或委訓，只能查自己的單位
                '下列計畫,可能有補助地方政府3層式架構 傳入PlanID 判斷 TPlanID
                'If TIMS.Check_TPlanID17(sm.UserInfo.PlanID, objconn) Then
                '    StrSql &= " AND (O1.OrgID='" & Orgidvalue.Value & "' OR r3.OrgID2='" & Orgidvalue.Value & "')" & vbCrLf  '補助地方政府'依業務層面呈現下層資料
                'Else
                '    StrSql &= " AND O1.OrgID='" & Orgidvalue.Value & "'" & vbCrLf  '非補助 '非3層架構
                'End If
                StrSql &= " AND O1.OrgID='" & Orgidvalue.Value & "'" & vbCrLf  '非補助 '非3層架構
                StrSql &= " AND I1.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End Select
        '依業務RID查詢資料 RIDValue
        StrSql &= " AND rr.relship like '" & relship & "%'" & vbCrLf
        '年度有選擇
        If v_yearlist <> "" Then StrSql &= " AND P1.PlanYear='" & v_yearlist & "'" & vbCrLf

        '產投/非產投判斷
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投 'Me.LabTMID.Text="訓練業別"
            If Me.jobValue.Value <> "" Then
                StrSql &= " AND (P1.TMID=" & jobValue.Value & " OR P1.TMID IN (" & vbCrLf
                StrSql &= "  SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf   '職類別
                StrSql &= "  SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf   '業別
                StrSql &= "  SELECT TMID FROM Key_TrainType WHERE busid='G')" & vbCrLf  '產業人才投資方案類
                StrSql &= "  AND tmid=" & Me.jobValue.Value & " ))" & vbCrLf
                StrSql &= "  OR P1.TMID='" & trainValue.Value & "')" & vbCrLf  'edit，by:20181029
            End If
        Else
            If trainValue.Value <> "" Then StrSql &= " AND P1.TMID='" & trainValue.Value & "'" & vbCrLf  '非產投
        End If

        txtCJOB_NAME.Text = TIMS.ClearSQM(txtCJOB_NAME.Text)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then
            StrSql &= " AND P1.CJOB_UNKEY='" & cjobValue.Value & "'" & vbCrLf
        End If

        '班名查詢
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        If ClassName.Text <> "" Then
            'ClassName.Text=Trim(ClassName.Text)
            'Dim vOCID As Integer=0
            If IsNumeric(ClassName.Text) Then
                Dim vOCID As Integer = TIMS.GetValue2(ClassName.Text)
                StrSql &= " AND (P1.ClassName LIKE '%" & ClassName.Text & "%' OR C1.ClassCName LIKE '%" & ClassName.Text & "%'" & vbCrLf
                StrSql &= " OR CONVERT(VARCHAR(111),P1.ClassName) LIKE '%" & ClassName.Text & "%' OR CONVERT(VARCHAR(111),C1.ClassCName) LIKE '%" & ClassName.Text & "%'" & vbCrLf
                If vOCID > 0 Then StrSql &= " OR C1.OCID='" & vOCID & "')" & vbCrLf  '有效數字
            Else
                '非數字
                StrSql &= " AND (P1.ClassName LIKE '%" & ClassName.Text & "%' OR C1.ClassCName LIKE '%" & ClassName.Text & "%'" & vbCrLf
                StrSql &= " OR CONVERT(VARCHAR(111),P1.ClassName) LIKE '%" & ClassName.Text & "%' OR CONVERT(VARCHAR(111),C1.ClassCName) LIKE '%" & ClassName.Text & "%')" & vbCrLf
            End If
        End If

        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then StrSql &= " AND P1.CyclType='" & CyclType.Text & "'" & vbCrLf
        'Dim dt As DataTable'DataGridTable.Visible = False'msg.Visible = True'Me.msg.Text = "查無資料"'dt.Load(.ExecuteReader())

        'Dim flag_chktest As Boolean = TIMS.sUtl_ChkTest()
        'https://dotblogs.com.tw/harry/2016/10/14/181017
        'Dim slogMsg1 As String = ""'slogMsg1 &= "##TC_02_001, StrSql: " & StrSql & vbCrLf'slogMsg1 &= "##TC_02_001, myParam: " & TIMS.GetMyValue3(myParam) & vbCrLf'If flag_chktest Then TIMS.writeLog(Me, slogMsg1)

        Try
            dt = DbAccess.GetDataTable(StrSql, objconn)
        Catch ex As Exception
            dt = Nothing
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("/* StrSql: */ {0}", StrSql) & vbCrLf
            strErrmsg &= String.Format("/* ex.ToString: */ {0}", ex.ToString) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Throw ex
        End Try

        Return dt
    End Function

    ''' <summary>
    ''' 自辦在職匯出使用
    ''' </summary>
    ''' <returns></returns>
    Function GetSchDtExp2() As DataTable
        Dim dt As DataTable = Nothing
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)

        '取出訓練計畫名稱
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn) 'RIDValue
        If drRR Is Nothing Then Return dt

        Dim relship As String = Convert.ToString(drRR("RELSHIP"))
        Orgidvalue.Value = Convert.ToString(drRR("OrgID"))

        Dim v_IsApprPaper As String = TIMS.GetListValue(IsApprPaper) '判斷是否是選正式-狀態
        Dim v_yearlist As String = TIMS.GetListValue(yearlist) '年度 'yearlist.SelectedValue
        Dim v_audit As String = TIMS.GetListValue(audit)  'audit.SelectedValue '審核狀態有選值

        '建立訓練的DATATABLE
        Dim StrSql As String = ""
        StrSql &= " SELECT k2.TRAINNAME" & vbCrLf '/*訓練職類*/" & vbCrLf
        StrSql &= " ,dd.D20KNAME1,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5,dd.D20KNAME6" & vbCrLf
        StrSql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8" & vbCrLf
        'StrSql &= " ,convert(nvarchar(MAX),NULL) D20KNAME" & vbCrLf '/*政策性課程類型 D20KNAME*/" & vbCrLf
        StrSql &= " ,convert(nvarchar(MAX),NULL) D2025KNAME" & vbCrLf '/*政策性課程類型*/" & vbCrLf

        StrSql &= " ,p1.CLASSNAME" & vbCrLf '/*班別名稱*/" & vbCrLf
        StrSql &= " ,replace(dbo.FN_GET_TRAINDESC(p1.PLANID,p1.COMIDNO,p1.SEQNO,'PNAME'),'^',',') DESCPNAME" & vbCrLf '/*課程內容*/" & vbCrLf
        StrSql &= " ,p1.THOURS" & vbCrLf '/*訓練時數*/" & vbCrLf
        StrSql &= " ,p1.TNUM" & vbCrLf '/*預訓人數*/" & vbCrLf
        StrSql &= " ,p1.ADVANCE" & vbCrLf '訓練課程類型 ADVANCE
        StrSql &= " ,case p1.ADVANCE when '01' then '基礎' when '02' then '進階' end ADVANCE_N" & vbCrLf '訓練課程類型 ADVANCE
        StrSql &= " ,p1.TOTALCOST" & vbCrLf '/*訓練費用(元)*/" & vbCrLf
        StrSql &= " ,p1.DEFSTDCOST" & vbCrLf '/*學員負擔費用(元)*/" & vbCrLf
        StrSql &= " ,p1.DEFGOVCOST" & vbCrLf '/*政府負擔費用(元)*/" & vbCrLf
        StrSql &= " ,p1.CAPOTHER1,p1.CAPOTHER2,p1.CAPOTHER3" & vbCrLf '/*參訓資格-其他條件1.2.3.*/" & vbCrLf
        StrSql &= " ,convert(nvarchar(MAX),NULL)  CAPOTHER" & vbCrLf '/*受訓資格*/" & vbCrLf
        StrSql &= " ,case p1.COACHING when 'Y' then '是' when 'N' then '否' end COACHING_N" & vbCrLf '/*是否輔導考照*/" & vbCrLf
        StrSql &= " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(p1.EXAMLVID1" & vbCrLf
        StrSql &= " ,'1','甲級'),'2','乙級'),'3','丙級'),'4','單一級') ,'5','不分級') EXAMLVID1_N" & vbCrLf '/*完訓後可參加證照考試級別*/" & vbCrLf
        StrSql &= " ,S1.EXNAME S1EXNAME_N" & vbCrLf '/*完訓後可參加之全國技術士技能檢定職類*/" & vbCrLf '
        StrSql &= " ,FORMAT(p1.SENTERDATE,'yyyy/MM/dd') SENTERDATE" & vbCrLf '/*報名開始日期*/" & vbCrLf
        StrSql &= " ,FORMAT(p1.FENTERDATE,'yyyy/MM/dd') FENTERDATE" & vbCrLf '/*報名結束日期*/" & vbCrLf
        StrSql &= " ,FORMAT(p1.EXAMDATE,'yyyy/MM/dd')  EXAMDATE" & vbCrLf '/*甄試日期*/" & vbCrLf
        'StrSql &= " /*甄試方式*/" & vbCrLf
        StrSql &= " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(p1.GETTRAIN3, ',', '，')" & vbCrLf
        StrSql &= " ,'1', '職業適性測驗'), '2', '筆試'), '3', '口試'), '4', '實作'), '5', '體能測驗'), '6', concat('其他: ',ISNULL(p1.GETTRAIN3OTHER,''))) GETTRAIN3_N" & vbCrLf
        StrSql &= " ,FORMAT(p1.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '/*開訓日期*/" & vbCrLf
        StrSql &= " ,FORMAT(p1.FDDATE,'yyyy/MM/dd') FDDATE" & vbCrLf '/*結訓日期*/" & vbCrLf
        'StrSql &= " ,p1.TPERIOD" & vbCrLf '/*訓練週期*/" & vbCrLf
        StrSql &= " ,f.HOURRANNAME" & vbCrLf '/*訓練週期*/" & vbCrLf
        StrSql &= " ,p1.NOTE3" & vbCrLf '/*訓練時段*/" & vbCrLf
        StrSql &= " ,dbo.FN_ADDR2(p1.TADDRESSZIP,p1.TADDRESSZIP6W,iz.ZIPNAME,p1.TADDRESS) TADDRESS_N" & vbCrLf
        StrSql &= " ,p1.TRANSFLAG ,p1.ISAPPRPAPER,p1.APPLIEDRESULT" & vbCrLf
        StrSql &= " ,O1.ORGNAME" & vbCrLf
        StrSql &= " ,concat(p1.planid,'x',p1.ComIDNO,'x',p1.SeqNO) PCS" & vbCrLf

        StrSql &= " FROM dbo.PLAN_PLANINFO p1" & vbCrLf
        StrSql &= " JOIN dbo.ID_PLAN ip on ip.planid=p1.planid" & vbCrLf
        StrSql &= " JOIN dbo.ORG_ORGINFO O1 WITH(NOLOCK) ON p1.ComIDNO=O1.ComIDNO" & vbCrLf
        StrSql &= " JOIN dbo.AUTH_RELSHIP rr WITH(NOLOCK) ON rr.RID=p1.RID" & vbCrLf
        StrSql &= " LEFT JOIN dbo.KEY_TRAINTYPE k2 on k2.tmid=p1.tmid" & vbCrLf
        StrSql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz on iz.ZIPCODE=p1.TADDRESSZIP" & vbCrLf
        StrSql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=p1.PLANID and dd.COMIDNO=p1.COMIDNO and dd.SEQNO=p1.SEQNO" & vbCrLf
        StrSql &= " LEFT JOIN dbo.KEY_HOURRAN F ON F.HRID=p1.TPERIOD" & vbCrLf
        StrSql &= " LEFT JOIN dbo.KEY_EXAM3 S1 ON S1.EXAMID=p1.EXAMIDS1" & vbCrLf
        StrSql &= " LEFT JOIN dbo.CLASS_CLASSINFO C1 WITH(NOLOCK) ON C1.PlanID=p1.PlanID AND C1.COMIDNO=p1.COMIDNO AND C1.SEQNO=p1.SEQNO AND C1.IsSuccess='Y'" & vbCrLf
        StrSql &= " WHERE 1=1" & vbCrLf
        'StrSql &= " and ip.PLANID>=5027 and ip.PLANID<=5044" & vbCrLf
        'StrSql &= " AND PP.ISAPPRPAPER='Y'" & vbCrLf
        'StrSql &= " AND PP.AppliedResult IN ('Y','N')" & vbCrLf
        'StrSql &= " --and p1.PLANID =5034" & vbCrLf
        'StrSql &= " and ip.TPLANID='06'" & vbCrLf
        'StrSql &= " and ip.YEARS='2021'" & vbCrLf

        '課程代碼
        Dim flag_can_use_OCID As Boolean = False '轉換數字後相等:true/false:異常
        s_OCID.Text = TIMS.ClearSQM(s_OCID.Text)
        If s_OCID.Text <> "" AndAlso TIMS.IsNumeric2(s_OCID.Text) Then
            Dim vs_OCID As String = Convert.ToString(CInt(Val(s_OCID.Text)))  '轉換數字後相等:true/false:異常
            If vs_OCID.Equals(s_OCID.Text) Then flag_can_use_OCID = True '轉換數字後相等:true/false:異常
        End If
        If flag_can_use_OCID Then StrSql &= " AND C1.OCID='" & s_OCID.Text & "'" & vbCrLf
        '依申請階段
        Dim v_AppStage As String = If(tr_AppStage_TP28.Visible, TIMS.GetListValue(AppStage), "")
        If v_AppStage <> "" Then StrSql &= " AND P1.AppStage='" & v_AppStage & "'" & vbCrLf '依申請階段

        'rbl_TransFlagS '增加【轉班上架】欄位，選項：不區分、未轉班、已轉班
        Dim v_rbl_TransFlagS As String = TIMS.GetListValue(rbl_TransFlagS)
        If v_rbl_TransFlagS = "A" Then v_rbl_TransFlagS = ""
        If v_rbl_TransFlagS <> "" Then StrSql &= " AND P1.TransFlag='" & v_rbl_TransFlagS & "'" & vbCrLf '轉班上架

        If Len(sm.UserInfo.RID) = 1 Then
            Select Case sm.UserInfo.RID
                Case "A" '署(局)權限
                    If v_yearlist <> "" Then '有選年度依年度
                        StrSql &= " AND ip.Years='" & v_yearlist & "'" & vbCrLf
                    Else '沒選年度依登入年度計畫@TPlanID
                        StrSql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                    End If
                Case Else
                    '分署(中心)特許依此計畫查詢有關權限 2008-09-30 AMU
                    If v_yearlist <> "" Then '有選年度依年度
                        StrSql &= " AND ip.Years='" & v_yearlist & "'" & vbCrLf
                    Else '沒選年度依登入年度計畫@TPlanID
                        StrSql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                    End If
            End Select
        Else
            StrSql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf  '依登入年度計畫 @PlanID
        End If

        StrSql &= " AND P1.TPlanID='" & TPlanid.Value & "'" & vbCrLf
        StrSql &= " AND P1.IsApprPaper='" & v_IsApprPaper & "'" & vbCrLf

        '產投/非產投判斷
        '如果不是產學訓計畫 penny 2007/10/17
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If v_IsApprPaper = "Y" Then  '判斷是否是選正式
                If v_audit <> "" Then       '如果審核狀態有選值
                    Select Case v_audit
                        Case "Y" '如果是選已審核
                            StrSql &= " AND P1.AppliedResult IN ('Y','N')" & vbCrLf
                        Case "N" '如果是選審核中
                            StrSql &= " AND (P1.AppliedResult NOT IN ('Y','N') OR P1.AppliedResult IS NULL)" & vbCrLf
                    End Select
                End If
            End If
        End If

        '有關機構的篩選
        Select Case Convert.ToString(sm.UserInfo.LID)  '登入者
            Case "0" '署(局)
            Case "1" '分署(中心)
                StrSql &= " AND ip.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
            Case Else '補助地方政府或委訓，只能查自己的單位
                StrSql &= " AND O1.OrgID='" & Orgidvalue.Value & "'" & vbCrLf  '非補助 '非3層架構
                StrSql &= " AND ip.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End Select
        '依業務RID查詢資料 RIDValue
        StrSql &= " AND rr.relship like '" & relship & "%'" & vbCrLf
        '年度有選擇
        If v_yearlist <> "" Then StrSql &= " AND P1.PlanYear='" & v_yearlist & "'" & vbCrLf
        '非產投- 自辦在職
        If trainValue.Value <> "" Then StrSql &= " AND P1.TMID='" & trainValue.Value & "'" & vbCrLf

        txtCJOB_NAME.Text = TIMS.ClearSQM(txtCJOB_NAME.Text)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then StrSql &= " AND P1.CJOB_UNKEY='" & cjobValue.Value & "'" & vbCrLf

        '班名查詢
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        If ClassName.Text <> "" Then
            'ClassName.Text=Trim(ClassName.Text) 'Dim vOCID As Integer=0
            If IsNumeric(ClassName.Text) Then
                Dim vOCID As Integer = TIMS.GetValue2(ClassName.Text)
                StrSql &= " AND (P1.ClassName LIKE '%" & ClassName.Text & "%' OR C1.ClassCName LIKE '%" & ClassName.Text & "%'" & vbCrLf
                StrSql &= " OR CONVERT(VARCHAR(111),P1.ClassName) LIKE '%" & ClassName.Text & "%' OR CONVERT(VARCHAR(111),C1.ClassCName) LIKE '%" & ClassName.Text & "%'" & vbCrLf
                If vOCID > 0 Then StrSql &= " OR C1.OCID='" & vOCID & "')" & vbCrLf  '有效數字
            Else
                '非數字
                StrSql &= " AND (P1.ClassName LIKE '%" & ClassName.Text & "%' OR C1.ClassCName LIKE '%" & ClassName.Text & "%'" & vbCrLf
                StrSql &= " OR CONVERT(VARCHAR(111),P1.ClassName) LIKE '%" & ClassName.Text & "%' OR CONVERT(VARCHAR(111),C1.ClassCName) LIKE '%" & ClassName.Text & "%')" & vbCrLf
            End If
        End If

        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then StrSql &= " AND P1.CyclType='" & CyclType.Text & "'" & vbCrLf

        Try
            dt = DbAccess.GetDataTable(StrSql, objconn)
        Catch ex As Exception
            dt = Nothing
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("/* StrSql: */ {0}", StrSql) & vbCrLf
            strErrmsg &= String.Format("/* ex.ToString: */ {0}", ex.ToString) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Throw ex
        End Try
        Return dt
    End Function

    '此功能有跨年度查詢功能，但並非本系統(TIMS)慣例，若是為慣例，則登入功能失效，將造成整個系統(TIMS)邏輯上的失敗。 AMU 2008-10-15
    Sub SearchData1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dtPlan)
        TPlanName.Text = TIMS.GetTPlanName(Convert.ToString(sm.UserInfo.TPlanID), objconn)

        '申請階段管理-受理期間設定 APPLISTAGE
        'Dim v_yearlist As String = TIMS.GetListValue(yearlist) '年度 'yearlist.SelectedValue
        'Dim v_AppStage As String = If(tr_AppStage_TP28.Visible, TIMS.GetListValue(AppStage), "")
        'If v_yearlist <> "" AndAlso v_AppStage <> "" Then
        '    Dim aParms As New Hashtable
        '    aParms.Add("YEARS", v_yearlist)
        '    aParms.Add("APPSTAGE", v_AppStage)
        '    開放受理之申請階段/ PLAN_APPSTAGE
        '    fg_can_applistage_G = TIMS.CAN_APPLISTAGE_PTYPE01(objconn, aParms)
        'End If

        Dim dt As DataTable = GetSchDt()
        If dt Is Nothing Then
            Common.MessageBox(Me, "資料庫查詢失敗，請重新查詢") ', rtnPath
            Exit Sub
        End If

        DataGridTable.Visible = False
        msg.Visible = True
        Me.msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            Session(hid_PPINFOtable_guid1.Value) = dt
            DataGridTable.Visible = True
            Me.msg.Text = ""
            If Me.ViewState(cst_Sort) = "" Then Me.ViewState(cst_Sort) = "STDate,ClassName"
            PageControler1.PageDataTable = dt
            PageControler1.Sort = Me.ViewState(cst_Sort)
            PageControler1.ControlerLoad()
        End If

        'v_IsApprPaper "Y" '正式 "N" '草稿 Else 'N:'草稿
        Dim v_IsApprPaper As String = TIMS.GetListValue(IsApprPaper)

        btnEnter1.Visible = If(flag_EnterShelf, If(v_IsApprPaper = "Y", True, False), False) '在職-批次轉班上架
        dtPlan.Columns(Cst_checkbox).Visible = If(flag_EnterShelf, If(v_IsApprPaper = "Y", True, False), False) '在職-批次轉班上架
        dtPlan.Columns(Cst_AppliedResult).Visible = If(v_IsApprPaper = "Y", True, False)
        dtPlan.Columns(Cst_VerReason).Visible = If(v_IsApprPaper = "Y", True, False)
        dtPlan.Columns(Cst_TransFlag).Visible = If(v_IsApprPaper = "Y", True, False)
        dtPlan.Columns(Cst_Function1).Visible = If(v_IsApprPaper = "Y", True, False)
        dtPlan.Columns(Cst_Function2).Visible = If(v_IsApprPaper = "Y", False, True)

        Select Case v_IsApprPaper
            Case "Y" '正式
            Case "N" '草稿 Else 'N:'草稿
            Case Else '異常狀況
                'tr_audit1.Visible = False
                'If tr_audit1.Visible Then tr_audit1.Style("display") = "none"
                dtPlan.Columns(Cst_AppliedResult).Visible = False
                dtPlan.Columns(Cst_VerReason).Visible = False
                dtPlan.Columns(Cst_TransFlag).Visible = False
                dtPlan.Columns(Cst_Function1).Visible = False
                dtPlan.Columns(Cst_Function2).Visible = False
        End Select
    End Sub

    Private Sub dtPlan_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles dtPlan.SortCommand
        'TIMS.LOG.Debug(String.Format("#e.SortExpression: {0}", e.SortExpression))
        'TIMS.LOG.Debug(String.Format("#PageControler1.Sort: {0}", PageControler1.Sort))
        Me.ViewState(cst_Sort) = If(Me.ViewState(cst_Sort) <> e.SortExpression, e.SortExpression, e.SortExpression & " DESC")
        PageControler1.Sort = Me.ViewState(cst_Sort)
        'TIMS.LOG.Debug(String.Format("#e.SortExpression: {0}", e.SortExpression))
        'TIMS.LOG.Debug(String.Format("#PageControler1.Sort: {0}", PageControler1.Sort))
        'PageControler1.DataTableCreate(CPdt, PageControler1.Sort)
        'Call SearchData1()
        If Session(hid_PPINFOtable_guid1.Value) Is Nothing Then Call SearchData1()
        If Session(hid_PPINFOtable_guid1.Value) IsNot Nothing Then
            PageControler1.DataTableCreate(Session(hid_PPINFOtable_guid1.Value), PageControler1.Sort)
        End If
    End Sub

    Sub KeepSearchStr()
        Dim str_search As String = ""
        str_search &= "prg=TC_02_001"
        str_search &= "&yearlist=" & TIMS.ClearSQM(yearlist.SelectedValue)
        str_search &= "&TB_career_id=" & TIMS.ClearSQM(TB_career_id.Text)
        str_search &= "&jobValue=" & TIMS.ClearSQM(jobValue.Value)  'edit，by:20181030
        str_search &= "&trainValue=" & TIMS.ClearSQM(trainValue.Value)
        str_search &= "&center=" & TIMS.ClearSQM(center.Text)
        str_search &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        str_search &= "&ClassName=" & TIMS.ClearSQM(ClassName.Text)
        str_search &= "&IsApprPaper=" & TIMS.ClearSQM(IsApprPaper.SelectedValue)
        str_search &= "&audit=" & TIMS.ClearSQM(audit.SelectedValue)
        str_search &= "&TransFlagS=" & TIMS.GetListValue(rbl_TransFlagS)
        str_search &= "&PageIndex=" & dtPlan.CurrentPageIndex + 1

        Session("search") = str_search
    End Sub

    Sub UseKeepSearchStr()
        If Session("search") Is Nothing Then Return

        Dim str_search As String = Convert.ToString(Session("search"))
        Session("search") = Nothing

        Dim MyValue As String = TIMS.GetMyValue(str_search, "prg")
        If Not MyValue = "TC_02_001" Then Return

        Common.SetListItem(yearlist, TIMS.GetMyValue(str_search, "yearlist"))
        TB_career_id.Text = TIMS.GetMyValue(str_search, "TB_career_id")
        jobValue.Value = TIMS.GetMyValue(str_search, "jobValue")  'edit，by:20181030
        trainValue.Value = TIMS.GetMyValue(str_search, "trainValue")
        center.Text = TIMS.GetMyValue(str_search, "center")
        RIDValue.Value = TIMS.GetMyValue(str_search, "RIDValue")
        ClassName.Text = TIMS.GetMyValue(str_search, "ClassName")
        Common.SetListItem(IsApprPaper, TIMS.GetMyValue(str_search, "IsApprPaper"))
        Common.SetListItem(audit, TIMS.GetMyValue(str_search, "audit"))
        Common.SetListItem(rbl_TransFlagS, TIMS.GetMyValue(str_search, "TransFlagS"))

        'btnQuery_Click(sender, e)
        Call SearchData1()

        MyValue = TIMS.GetMyValue(str_search, "PageIndex")
        If IsNumeric(MyValue) AndAlso Session(hid_PPINFOtable_guid1.Value) IsNot Nothing Then
            PageControler1.PageIndex = MyValue
            PageControler1.DataTableCreate(Session(hid_PPINFOtable_guid1.Value), PageControler1.Sort, PageControler1.PageIndex)
        End If

    End Sub

    Private Sub yearlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yearlist.SelectedIndexChanged
        DataGridTable.Visible = False
        msg.Visible = False
    End Sub

    ''' <summary>
    ''' 檢核正常為true 異常 false
    ''' </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""

        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then
            If Not IsNumeric(CyclType.Text) Then
                Errmsg &= "期別需輸入數字型態!!" & vbCrLf
                'Common.MessageBox(Me, "期別需輸入數字型態!!") Exit Sub
            End If
        End If

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary> 匯出-增加-檢核</summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData2(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""
        Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 

        Dim v_IsApprPaper As String = TIMS.GetListValue(IsApprPaper) '判斷是否是選正式-狀態
        If v_IsApprPaper <> "Y" Then Errmsg &= "匯出 資料類型 只能選正式!!" & vbCrLf
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then Errmsg &= "匯出 訓練機構 不能為空!!" & vbCrLf
        If RIDValue.Value = "A" Then Errmsg &= "匯出 訓練機構 不能為全部!!" & vbCrLf

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>
    ''' 匯出功能
    ''' </summary>
    Sub ExportData1()
        Dim dt As DataTable = GetSchDt()
        'Dim rtnPath As String = Request.FilePath
        If dt Is Nothing Then
            Common.MessageBox(Me, "資料庫查詢失敗，請重新查詢!!") ', rtnPath
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料，請重新查詢!")
            Exit Sub
        End If
        'msg.Text = ""

        '匯出 Response 
        ExpReport1(dt)
    End Sub

    ''' <summary>
    ''' 匯出 Response - 產投
    ''' </summary>
    ''' <param name="dt"></param>
    Sub ExpReport1(ByRef dt As DataTable)
        '匯出表頭名稱
        Dim sFileName1 As String = "班級查詢" & TIMS.GetDateNo2()

        Const cst_tit1 As String = "年度,計畫,分署,申請日期,訓練起日,訓練迄日,申請階段,訓練單位,課程申請流水號,班別名稱,訓練時數,訓練人數,核定人數,審核狀態,未通過原因"
        Const cst_tit2 As String = "PlanYear,PlanName,OrgName2,AppliedDate,STDate,FDDate,AppStage,OrgName,PSNO28,CLASSNAME2,THOURS,TNUM,TNUM,AppliedResult,VerReason"
        Dim sta_tit1 As String() = Split(cst_tit1, ",")
        Dim sta_tit2 As String() = Split(cst_tit2, ",")

        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim s_TDVALUE As String = ""
        Dim flag_isDate As Boolean = False '日期判斷
        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        ExportStr = "<tr>" & vbCrLf
        For Each str1 As String In sta_tit1
            ExportStr &= "<td>" & str1 & "</td>" & vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        sbHTML.Append(TIMS.sUtl_AntiXss(ExportStr))

        For Each dr As DataRow In dt.Rows
            'For Each dr As DataRow In dt.Rows
            '建立資料面
            ExportStr = "<tr>" & vbCrLf
            For Each str_cl2 As String In sta_tit2
                flag_isDate = False
                If str_cl2.Equals("AppliedDate") Then flag_isDate = True
                If str_cl2.Equals("STDate") Then flag_isDate = True
                If str_cl2.Equals("FDDate") Then flag_isDate = True

                s_TDVALUE = If(flag_isDate, TIMS.Cdate3(dr(str_cl2)), Convert.ToString(dr(str_cl2)))
                If str_cl2.Equals("AppliedResult") Then
                    s_TDVALUE = Get_AppliedResultTxt(sm.UserInfo.TPlanID, Convert.ToString(dr("AppliedResult")), Convert.ToString(dr("RESULTBUTTON")))
                End If
                If str_cl2.Equals("VerReason") AndAlso s_TDVALUE <> "" Then
                    s_TDVALUE = Replace(s_TDVALUE, vbCrLf, TIMS.cst_Html2ExcelLineBreak)
                End If
                ExportStr &= "<td>" & s_TDVALUE & "</td>" & vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(TIMS.sUtl_AntiXss(ExportStr))
        Next
        sbHTML.Append("</table>")
        sbHTML.Append("</div>")


        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    ''' <summary>匯出功能 - 自辦在職</summary>
    Sub ExportData2()
        Dim dt2 As DataTable = GetSchDtExp2()
        'Dim rtnPath As String = Request.FilePath
        If dt2 Is Nothing Then
            Common.MessageBox(Me, "資料庫查詢失敗，請重新查詢!!") ', rtnPath
            Exit Sub
        End If
        If dt2.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料，請重新查詢!")
            Exit Sub
        End If
        'msg.Text = ""

        '匯出 Response 
        ExpReport2(dt2)
    End Sub

    Sub UPDATE_DataRow2(ByRef dr As DataRow)
        Dim s_D20KNAME As String = ""
        TIMS.ADD_PAR2(s_D20KNAME, $"{dr("D20KNAME1")}")
        TIMS.ADD_PAR2(s_D20KNAME, $"{dr("D20KNAME2")}")
        TIMS.ADD_PAR2(s_D20KNAME, $"{dr("D20KNAME3")}")
        TIMS.ADD_PAR2(s_D20KNAME, $"{dr("D20KNAME4")}")
        TIMS.ADD_PAR2(s_D20KNAME, $"{dr("D20KNAME5")}")
        TIMS.ADD_PAR2(s_D20KNAME, $"{dr("D20KNAME6")}")
        Dim s_D25KNAME As String = ""
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME1")}")
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME2")}")
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME3")}")
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME4")}")
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME5")}")
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME6")}")
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME7")}")
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME8")}")
        TIMS.ADD_PAR2(s_D25KNAME, $"{dr("D25KNAME8")}")
        Dim s_D2025KNAME As String = ""
        If (s_D20KNAME <> "") Then TIMS.ADD_PAR2(s_D2025KNAME, s_D20KNAME)
        If (s_D25KNAME <> "") Then TIMS.ADD_PAR2(s_D2025KNAME, s_D25KNAME)
        dr("D2025KNAME") = s_D2025KNAME
        Dim s_CAPOTHER As String = ""
        If (dr("CAPOTHER1").ToString() <> "") Then TIMS.ADD_PAR2(s_CAPOTHER, String.Format("1.{0}", dr("CAPOTHER1").ToString()))
        If (dr("CAPOTHER2").ToString() <> "") Then TIMS.ADD_PAR2(s_CAPOTHER, String.Format("2.{0}", dr("CAPOTHER2").ToString()))
        If (dr("CAPOTHER3").ToString() <> "") Then TIMS.ADD_PAR2(s_CAPOTHER, String.Format("3.{0}", dr("CAPOTHER3").ToString()))
        dr("CAPOTHER") = s_CAPOTHER
    End Sub

    ''' <summary>
    ''' 匯出 Response - 自辦在職
    ''' </summary>
    ''' <param name="dt"></param>
    Sub ExpReport2(ByRef dt As DataTable)
        If dt.Rows.Count = 0 Then Return
        Dim drT1 As DataRow = dt.Rows(0)

        '匯出表頭名稱
        Dim sFileName1 As String = String.Format("export_{0}", TIMS.GetDateNo2())
        Dim s_TitleName As String = String.Format("勞動部{0} 開班預定表", drT1("ORGNAME").ToString())

        Dim str_tit1 As String = ""
        str_tit1 = "訓練職類,政策性課程類型,班別名稱,課程內容,訓練課程類型,訓練時數,預訓人數,訓練費用(元),學員負擔費用(元),政府負擔費用(元),受訓資格"
        str_tit1 &= ",是否輔導考照,完訓後可參加證照考試級別,完訓後可參加之全國技術士技能檢定職類,報名開始日期,報名結束日期,甄試日期,甄試方式"
        '訓練週期→訓練時段、訓練時段→訓練週期及時間 
        str_tit1 &= ",開訓日期,結訓日期,訓練時段,訓練週期及時間,上課地點"
        Dim str_tit2 As String = ""
        str_tit2 = "TRAINNAME,D2025KNAME,CLASSNAME,DESCPNAME,ADVANCE_N,THOURS,TNUM,TOTALCOST,DEFSTDCOST,DEFGOVCOST,CAPOTHER"
        str_tit2 &= ",COACHING_N,EXAMLVID1_N,S1EXNAME_N,SENTERDATE,FENTERDATE,EXAMDATE,GETTRAIN3_N"
        str_tit2 &= ",STDATE,FDDATE,HOURRANNAME,NOTE3,TADDRESS_N"
        '訓練時數	預訓人數	訓練費用(元)	學員負擔費用(元)	政府負擔費用(元)
        Dim str_tit3 As String = "THOURS,TNUM,TOTALCOST,DEFSTDCOST,DEFGOVCOST" '(純數字)
        Dim sta_tit1 As String() = Split(str_tit1, ",")
        Dim sta_tit2 As String() = Split(str_tit2, ",")
        Dim sta_tit3 As String() = Split(str_tit3, ",") '(純數字)
        Dim iColSpanCount As Integer = sta_tit1.Length + 1

        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Const cst_ColFormat1 As String = "<td>{0}</td>"
        Const cst_ColFormat2 As String = "<td class=""noDecFormat"">{0}</td>" '(純數字)
        Dim ExportStr As String = ""
        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>" & vbCrLf)
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">" & vbCrLf)
        '表頭及查詢條件列
        ExportStr = String.Format("<tr><td align='center' colspan='{0}'>{1}</td></tr>", iColSpanCount, s_TitleName) & vbCrLf
        sbHTML.Append(ExportStr)
        '建立抬頭 '第1行
        ExportStr = "<tr>" & vbCrLf
        ExportStr &= String.Format(cst_ColFormat1, "序號") & vbTab
        For Each str1 As String In sta_tit1
            ExportStr &= String.Format(cst_ColFormat1, str1) & vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        sbHTML.Append(TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim iStudCnt As Integer = 0
        Dim iRow As Integer = 0
        For Each dr As DataRow In dt.Rows
            iStudCnt += Val(If(Convert.ToString(dr("TNum")) <> "", dr("TNum"), 0))
            iRow += 1
            UPDATE_DataRow2(dr)
            ExportStr = "<tr>" & vbCrLf
            ExportStr &= String.Format(cst_ColFormat1, iRow) & vbTab '"序號") & vbTab
            For Each str_cl2 As String In sta_tit2
                Dim flag_find1 As Boolean = TIMS.FindValue1(sta_tit3, str_cl2) '(純數字)
                Dim s_ColoumFMT2 As String = cst_ColFormat1
                If flag_find1 Then s_ColoumFMT2 = cst_ColFormat2 '(純數字)
                ExportStr &= String.Format(s_ColoumFMT2, dr(str_cl2)) & vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(TIMS.sUtl_AntiXss(ExportStr))
        Next

        Dim s_ClassNUM1 As String = String.Format("合計 {0} 班", dt.Rows.Count)
        ExportStr = String.Format("<tr><td align='left' colspan='{0}'>{1}</td></tr>", iColSpanCount, s_ClassNUM1) & vbCrLf
        sbHTML.Append(ExportStr)
        Dim s_StudTNUM1 As String = String.Format("訓練 {0} 人", iStudCnt)
        ExportStr = String.Format("<tr><td align='left' colspan='{0}'>{1}</td></tr>", iColSpanCount, s_StudTNUM1) & vbCrLf
        sbHTML.Append(ExportStr)
        sbHTML.Append("</table>" & vbCrLf)
        sbHTML.Append("</div>" & vbCrLf)

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    ''' <summary>查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Call SearchData1()
    End Sub

    ''' <summary>
    ''' 匯出1
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Call CheckData2(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        '匯出功能
        Call ExportData1()
    End Sub

    ''' <summary>匯出2 - 匯出開班預定表</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnExport2_Click(sender As Object, e As EventArgs) Handles btnExport2.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Call CheckData2(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        '匯出功能
        Call ExportData2()
    End Sub


    Public Shared Function Get_PLANINFOdata(ByRef oConn As SqlConnection, ByVal vPCS As String) As DataRow
        Dim drPP As DataRow = Nothing
        If vPCS = "" Then Return drPP

        Dim parms As New Hashtable
        parms.Add("PCS", vPCS)

        Dim sql As String = ""
        sql &= " SELECT p1.PlanID,p1.COMIDNO,p1.SEQNO" & vbCrLf
        sql &= " ,p1.TMID,p1.CJOB_UNKEY,p1.RID" & vbCrLf
        sql &= " ,p1.CLASSNAME,p1.CLASSENGNAME,p1.CYCLTYPE" & vbCrLf
        sql &= " ,p1.IsApprPaper" & vbCrLf
        sql &= " ,ip.YEARS,ip.DISTID,ip.TPLANID "
        sql &= " FROM dbo.PLAN_PLANINFO p1" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN ip on ip.PlanID=p1.PlanID" & vbCrLf
        sql &= " WHERE concat(p1.planid,'x',p1.ComIDNO,'x',p1.SeqNO)=@PCS"
        sql &= " AND p1.APPLIEDRESULT='Y'" & vbCrLf '班級審核狀態
        sql &= " AND p1.ISAPPRPAPER='Y'" & vbCrLf '正式資料-只有正式資料(正式儲存過)且未轉班，才會顯示
        sql &= " AND p1.TRANSFLAG='N'" & vbCrLf '未轉班

        drPP = DbAccess.GetOneRow(sql, oConn, parms)
        Return drPP
    End Function

    ''' <summary> '1.產生/取得一組班別代碼 </summary>
    ''' <param name="vPCS"></param>
    ''' <returns></returns>
    Public Shared Function GET_IDCLASS_CLSID(ByRef oConn As SqlConnection, ByVal vPCS As String) As String
        Dim rst As String = "" 'CLSID

        Dim drPP As DataRow = Get_PLANINFOdata(oConn, vPCS)
        If drPP Is Nothing Then Return rst
        'Dim vCLASSNAME As String = drPP("CLASSNAME").ToString() '班別名稱
        'Dim vCLASSENAME As String = drPP("CLASSENGNAME").ToString()
        Dim vTPLANID As String = drPP("TPLANID").ToString()
        Dim vDISTID As String = drPP("DISTID").ToString()
        Dim vYEARS As String = drPP("YEARS").ToString()
        Dim vTMID As String = drPP("TMID").ToString()
        Dim vCJOB_UNKEY As String = drPP("CJOB_UNKEY").ToString()
        'Dim v_CYCLTYPE As String = Convert.ToString(drPP("CYCLTYPE"))
        'Dim v_PlanID As String = drPP("PlanID").ToString()
        'Dim v_RID As String = drPP("RID").ToString()

        '同(轄區／年度／計畫／職類) , 開班時，產生一個新的班級代碼 , 但不同的 (轄區／年度／計畫／職類), 則不會產生新的一組
        Dim parms As New Hashtable
        parms.Add("DISTID", vDISTID)
        parms.Add("YEARS", vYEARS)
        parms.Add("TPLANID", vTPLANID)
        parms.Add("TMID", vTMID)
        parms.Add("CJOB_UNKEY", vCJOB_UNKEY)
        Dim sql As String = ""
        sql &= " SELECT CLSID FROM ID_CLASS a"
        sql &= " WHERE a.DISTID=@DISTID"
        sql &= " AND a.YEARS=@YEARS"
        sql &= " AND a.TPLANID=@TPLANID"
        sql &= " and a.TMID=@TMID"
        sql &= " and a.CJOB_UNKEY=@CJOB_UNKEY"
        Dim drDC As DataRow = DbAccess.GetOneRow(sql, oConn, parms)

        If drDC Is Nothing Then
            Dim iCLSID As Integer = ADD_IDCLASS_CLSID(oConn, drPP)
            If iCLSID = -1 Then Return rst '(異常離開)

            drDC = DbAccess.GetOneRow(sql, oConn, parms)
            If drDC Is Nothing Then Return rst
            rst = drDC("CLSID").ToString()
            Return rst 'CLSID
        Else
            Dim vCLSID As String = drDC("CLSID").ToString() '(使用現有的CLSID)

            Dim blnChkIsDouble As Boolean = False '沒有重複:False /重複:True CLSID 

            blnChkIsDouble = CHECK_DOUBLE_CLSID(oConn, drPP, vCLSID)

            If blnChkIsDouble Then '有重複，再產生一個新的CLSID, 避免重複
                Dim iCLSID As Integer = ADD_IDCLASS_CLSID(oConn, drPP)
                If iCLSID = -1 Then Return rst '(異常離開)

                rst = iCLSID.ToString() '新的CLSID
                Return rst 'CLSID
                'Common.MessageBox(MyPage, "轉入班級資料 班別代碼與期別重複!!")
                'rErrMsg1 = "轉入班級資料 班別代碼與期別重複!!"
                'Return rErrMsg1 'Exit Sub
            End If

            '(查無重複可以使用現有的CLSID)
            rst = vCLSID
        End If
        Return rst 'CLSID
    End Function

    ''' <summary> 檢核CLSID是否為重複使用 重複為TRUE 其它為FALSE (同轄區計畫只能1筆)</summary>
    ''' <param name="drPP"></param>
    ''' <returns></returns>
    Public Shared Function CHECK_DOUBLE_CLSID(ByRef oConn As SqlConnection, ByRef drPP As DataRow, ByVal vCLSID As String) As Boolean
        Dim blnChkIsDouble As Boolean = False '沒有重複FALSE (重複為TRUE 其它為FALSE)
        If drPP Is Nothing Then Return blnChkIsDouble

        'Dim vTMID As String = drPP("TMID").ToString()
        'Dim vCJOB_UNKEY As String = drPP("CJOB_UNKEY").ToString()
        'Dim vCLASSNAME As String = drPP("CLASSNAME").ToString() '班別名稱
        'Dim vCLASSENAME As String = drPP("CLASSENGNAME").ToString()
        'Dim vTPLANID As String = drPP("TPLANID").ToString()
        'Dim vDISTID As String = drPP("DISTID").ToString()
        'Dim vYEARS As String = drPP("YEARS").ToString()
        Dim v_CYCLTYPE As String = $"{drPP("CYCLTYPE")}"
        Dim v_PlanID As String = $"{drPP("PlanID")}"
        Dim v_RID As String = $"{drPP("RID")}"

        Dim PMSck As New Hashtable() From {{"CLSID", vCLSID}, {"PlanID", v_PlanID}, {"RID", v_RID}}
        Dim check_sql As String = ""
        check_sql &= " SELECT concat('(',dbo.FN_CLASSID2(cc.CLSID),')',dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE)) CLASSCNAME" & vbCrLf
        check_sql &= " FROM dbo.CLASS_CLASSINFO cc" & vbCrLf
        check_sql &= " WHERE cc.CLSID=@CLSID" & vbCrLf '依班別代碼(重複)
        check_sql &= " AND cc.PlanID=@PlanID AND cc.RID=@RID" & vbCrLf 'PlanID,機構
        If v_CYCLTYPE <> "" Then
            check_sql &= " AND cc.CyclType=@CyclType" & vbCrLf '期別(重複)
            PMSck.Add("CyclType", v_CYCLTYPE)
        Else
            check_sql &= " AND cc.CyclType IS NULL" & vbCrLf '期別(重複)
        End If
        'check = False '沒有重複
        Dim dr9 As DataRow = DbAccess.GetOneRow(check_sql, oConn, PMSck)

        If dr9 IsNot Nothing Then blnChkIsDouble = True '重複

        Return blnChkIsDouble
    End Function

    ''' <summary> 檢核 CLASSID 是否為重複使用 重複為TRUE 沒有重複FALSE 其它異常為TRUE (同年同計畫同轄區只能1筆) </summary>
    Public Shared Function CHECK_DOUBLE_CLASSID(ByRef oConn As SqlConnection, ByRef drPP As DataRow, ByVal vCLASSID As String) As Boolean
        Dim blnChkIsDouble As Boolean = True '重複為TRUE (沒有重複FALSE) (其它異常為TRUE)
        If drPP Is Nothing Then Return blnChkIsDouble
        If vCLASSID = "" Then Return blnChkIsDouble

        Dim vDISTID As String = drPP("DISTID").ToString()
        Dim vTPLANID As String = drPP("TPLANID").ToString()
        Dim vYEARS As String = drPP("YEARS").ToString()

        Dim s_parms As New Hashtable
        s_parms.Add("YEARS", vYEARS)
        s_parms.Add("TPLANID", vTPLANID)
        s_parms.Add("DISTID", vDISTID)
        s_parms.Add("CLASSID", vCLASSID)
        Dim sql As String = ""
        sql &= " SELECT CLASSID FROM ID_CLASS WHERE YEARS=@YEARS" & vbCrLf
        sql &= " AND TPLANID=@TPLANID" & vbCrLf
        sql &= " AND DISTID=@DISTID" & vbCrLf
        sql &= " AND CLASSID=@CLASSID" & vbCrLf
        Dim dr9 As DataRow = DbAccess.GetOneRow(sql, oConn, s_parms)

        If dr9 Is Nothing Then blnChkIsDouble = False '沒有重複

        Return blnChkIsDouble
    End Function

    ''' <summary> 產生一組新的CLSID 回傳CLSID (-1 表示有誤) </summary>
    ''' <param name="drPP"></param>
    ''' <returns></returns>
    Public Shared Function ADD_IDCLASS_CLSID(ByRef oConn As SqlConnection, ByRef drPP As DataRow) As Integer
        'Dim drPP As DataRow = Get_PLANINFOdata(vPCS)
        If drPP Is Nothing Then Return -1 '資料異常失敗

        Dim vCLASSNAME As String = drPP("CLASSNAME").ToString() '班別名稱
        Dim vCLASSENAME As String = drPP("CLASSENGNAME").ToString()
        Dim vTPLANID As String = drPP("TPLANID").ToString()
        Dim vTMID As String = drPP("TMID").ToString()
        Dim vDISTID As String = drPP("DISTID").ToString()
        Dim vCJOB_UNKEY As String = drPP("CJOB_UNKEY").ToString()
        Dim vYEARS As String = drPP("YEARS").ToString()
        'Dim v_CYCLTYPE As String = Convert.ToString(drPP("CYCLTYPE"))
        'Dim v_PlanID As String = drPP("PlanID").ToString()
        'Dim v_RID As String = drPP("RID").ToString()

        '檢核 CLASSID 是否為重複使用 重複為TRUE 沒有重複FALSE 其它為TRUE (同年同計畫同轄區只能1筆)
        '試著做3次
        Dim vCLASSID1 As String = TIMS.GetRndEngN(4) '取出4碼(亂數)-1
        Dim flag_doubleCLASSID1 As Boolean = CHECK_DOUBLE_CLASSID(oConn, drPP, vCLASSID1)
        Dim vCLASSID2 As String = TIMS.GetRndEngN(4) '取出4碼(亂數)-2
        Dim flag_doubleCLASSID2 As Boolean = CHECK_DOUBLE_CLASSID(oConn, drPP, vCLASSID2)
        Dim vCLASSID3 As String = TIMS.GetRndEngN(4) '取出4碼(亂數)-3
        Dim flag_doubleCLASSID3 As Boolean = CHECK_DOUBLE_CLASSID(oConn, drPP, vCLASSID3)

        If flag_doubleCLASSID1 AndAlso flag_doubleCLASSID2 AndAlso flag_doubleCLASSID3 Then Return -1 '班別代碼(都)重複 失敗
        '取得未重複的1筆資料
        Dim vCLASSID As String = If(Not flag_doubleCLASSID1, vCLASSID1, If(Not flag_doubleCLASSID2, vCLASSID2, If(Not flag_doubleCLASSID3, vCLASSID3, "")))
        '最後沒資料 異常
        If vCLASSID = "" Then Return -1

        Dim iCLSID As Integer = DbAccess.GetNewId(oConn, "ID_CLASS_CLSID_SEQ,ID_CLASS,CLSID")
        Dim i_sql As String = ""
        i_sql &= " INSERT INTO ID_CLASS( CLSID ,CLASSID, CLASSNAME, CLASSENAME ,TPLANID, TMID, DISTID, CJOB_UNKEY, YEARS" & vbCrLf
        i_sql &= " ,MODIFYACCT, MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES ( @CLSID ,@CLASSID, @CLASSNAME, @CLASSENAME ,@TPLANID, @TMID, @DISTID, @CJOB_UNKEY, @YEARS" & vbCrLf
        i_sql &= " ,@MODIFYACCT, GETDATE())" & vbCrLf
        Dim i_parms As New Hashtable
        i_parms.Add("CLSID", iCLSID)
        i_parms.Add("CLASSID", vCLASSID)
        i_parms.Add("CLASSNAME", vCLASSNAME) '班別名稱
        i_parms.Add("CLASSENAME", vCLASSENAME) '班別名稱(ENG)
        i_parms.Add("TPLANID", vTPLANID)
        i_parms.Add("TMID", vTMID)
        i_parms.Add("DISTID", vDISTID)
        i_parms.Add("CJOB_UNKEY", If(vCJOB_UNKEY <> "", Val(vCJOB_UNKEY), Convert.DBNull))
        i_parms.Add("YEARS", vYEARS)
        i_parms.Add("MODIFYACCT", "system1")
        TIMS.LOG.Debug(String.Format("#ID_CLASS: {0}", TIMS.GetMyValue3(i_parms)))
        DbAccess.ExecuteNonQuery(i_sql, oConn, i_parms)
        Return iCLSID
        'drDC = DbAccess.GetOneRow(Sql, objconn, parms)
    End Function

    ''' <summary> 轉入資料(SAVE) PLAN_PLANINFO CLASS_CLASSINFO - 自辦在職 </summary>
    Public Shared Function SAVE_CHANGEDATA_06(ByRef sm As SessionModel, ByRef htSS As Hashtable, ByRef oConn As SqlConnection) As String
        Dim rErrMsg1 As String = ""

        Dim vCLSID As String = TIMS.GetMyValue2(htSS, "CLSID")
        Dim vCJOB_UNKEY As String = TIMS.GetMyValue2(htSS, "CJOB_UNKEY")
        Dim vPCS As String = TIMS.GetMyValue2(htSS, "PCS")
        Dim vPlanID As String = TIMS.GetMyValue2(htSS, "PlanID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(htSS, "COMIDNO")
        Dim vSEQNO As String = TIMS.GetMyValue2(htSS, "SEQNO")
        Dim vTPlanID As String = TIMS.GetMyValue2(htSS, "TPlanID")
        Dim vRID1 As String = TIMS.GetMyValue2(htSS, "RID1")
        Dim vModifyUserID As String = TIMS.GetMyValue2(htSS, "ModifyUserID")

        If vCLSID = "" Then Return "傳入參數有誤!"
        If vCJOB_UNKEY = "" Then Return "傳入參數有誤!"
        If vPCS = "" Then Return "傳入參數有誤!"
        If vPlanID = "" Then Return "傳入參數有誤!"
        If vCOMIDNO = "" Then Return "傳入參數有誤!"
        If vSEQNO = "" Then Return "傳入參數有誤!"
        If vTPlanID = "" Then Return "傳入參數有誤!"
        If vRID1 = "" Then Return "傳入參數有誤!"
        If vModifyUserID = "" Then Return "傳入參數有誤!"

        'Dim htPP As New Hashtable
        'sqldr("ModifyAcct") = sm.UserInfo.UserID
        Dim parms As Hashtable = New Hashtable From {{"PLANID", vPlanID}, {"COMIDNO", vCOMIDNO}, {"SEQNO", vSEQNO}}
        Dim sqlpp As String = ""
        sqlpp = "SELECT * FROM PLAN_PLANINFO WHERE PLANID=@PLANID and COMIDNO=@COMIDNO and SEQNO=@SEQNO"
        Dim drPPinfo As DataRow = DbAccess.GetOneRow(sqlpp, oConn, parms)
        If drPPinfo Is Nothing Then
            'Common.MessageBox(MyPage, "計畫資料有誤，請重新選擇!!")
            rErrMsg1 = "計畫資料有誤，請重新選擇!!"
            Return rErrMsg1 'Exit Sub
        End If

        If Convert.ToString(drPPinfo("TNum")) = "" OrElse Val(drPPinfo("TNum")) = 0 Then rErrMsg1 &= "轉入計畫 資料有誤，訓練人數不可為0!!" & vbCrLf
        If rErrMsg1 <> "" Then Return rErrMsg1

        '(TIMS專用非產投) 'TC_01_004_InsertPlan.aspx
        '(產投)TIMS.Utl_Redirect1(Me, "TC_01_004_BusAdd.aspx?ID=" & Request("ID") & "&STDate=" & vsSTDate)
        'strScript1 &= "location.href='TC_01_004_add.aspx?ProcessType=PlanUpdate&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        'Const cst_temp_classinfo As String = "temp_classinfo" 'Session(cst_temp_classinfo)

        Dim blnChkIsDouble As Boolean = False '沒有重複FALSE (重複為TRUE 其它為FALSE)

        blnChkIsDouble = CHECK_DOUBLE_CLSID(oConn, drPPinfo, vCLSID)

        If blnChkIsDouble Then 'true:重複
            rErrMsg1 = "轉入班級資料 班別代碼與期別重複!!"
            Return rErrMsg1 'Exit Sub
        End If

        Dim parms9 As New Hashtable From {{"CLSID", vCLSID}}
        Dim sql9 As String = "SELECT * FROM ID_CLASS WHERE CLSID=@CLSID"
        Dim dr9 As DataRow = DbAccess.GetOneRow(sql9, oConn, parms9)
        If dr9 Is Nothing Then
            'Common.MessageBox(MyPage, "轉入用班級代碼異常，不可轉入!!")
            rErrMsg1 = "轉入用班級代碼異常，不可轉入!!"
            Return rErrMsg1 'Exit Sub
        End If
        If Convert.ToString(dr9("CJOB_UNKEY")) = "" Then
            'Common.MessageBox(MyPage, "轉入失敗,請聯絡承辦人設定此班別代碼的通俗職類資料,才可進行開班轉入動作!!")
            rErrMsg1 = "轉入失敗,請聯絡承辦人設定此班別代碼的通俗職類資料,才可進行開班轉入動作!!"
            Return rErrMsg1 'Exit Sub
        End If

        'MyPage.Session(cst_temp_classinfo) = Nothing
        Dim vCLASSENAME As String = dr9("CLASSENAME").ToString()
        '一般計畫檢核。
        Dim pms1 As New Hashtable From {{"RID", drPPinfo("RID")}}
        Dim sql As String = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID=@RID"
        Dim dtX As DataTable = DbAccess.GetDataTable(sql, oConn, pms1)
        If dtX.Rows.Count <> 1 Then
            'Common.MessageBox(MyPage, "業務權限異常，不可轉入!!")
            rErrMsg1 = "業務權限異常，不可轉入!!"
            Return rErrMsg1 'Exit Sub
        End If
        dr9 = dtX.Rows(0)
        Dim relship As String = dr9("RELSHIP").ToString()

        Dim pms2 As New Hashtable From {{"PLANID", drPPinfo("PLANID")}, {"COMIDNO", drPPinfo("COMIDNO")}, {"SEQNO", drPPinfo("SEQNO")}}
        Dim sql2 As String = "SELECT 'x' FROM PLAN_TRAINDESC WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO ORDER BY PTDID"
        Dim dtX2 As DataTable = DbAccess.GetDataTable(sql2, oConn, pms2)
        If TIMS.dtNODATA(dtX2) Then
            'Common.MessageBox(MyPage, "計畫訓練內容簡介無資料，不可轉入!!")
            rErrMsg1 = "計畫訓練內容簡介無資料，不可轉入!!"
            Return rErrMsg1 'Exit Sub
        End If

        'TIMS專用
        Dim pms3 As New Hashtable From {{"PLANID", drPPinfo("PLANID")}, {"COMIDNO", drPPinfo("COMIDNO")}, {"SEQNO", drPPinfo("SEQNO")}}
        Dim sql3 As String = " SELECT PCONT,PNAME FROM PLAN_TRAINDESC WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO ORDER BY PTDID"
        Dim dtTRAINDESC As DataTable = DbAccess.GetDataTable(sql3, oConn, pms3)
        Dim class_PName As String = ""
        For Each drTra As DataRow In dtTRAINDESC.Rows
            Dim sPName As String = TIMS.ClearSQM(drTra("PName"))
            If sPName <> "" Then class_PName &= String.Concat(If(class_PName <> "", ",", ""), sPName)
        Next

        'Dim objTrans As SqlTransaction
        Dim oAdp As SqlDataAdapter = Nothing
        Dim sqldr As DataRow = Nothing
        Dim pp_years As String = CInt(drPPinfo("PlanYear"))
        Dim sqlTable As New DataTable

        Dim iOCID_New As Integer = DbAccess.GetNewId(oConn, "CLASS_CLASSINFO_OCID_SEQ,CLASS_CLASSINFO,OCID")

        sql = " SELECT * FROM CLASS_CLASSINFO WHERE 1<>1"
        sqlTable = DbAccess.GetDataTable(sql, oAdp, oConn)
        sqldr = sqlTable.NewRow 'CLASS_CLASSINFO
        sqlTable.Rows.Add(sqldr)
        sqldr("OCID") = iOCID_New
        sqldr("ONSHELLDATE") = TIMS.Cdate2(CDate(Now).ToString("yyyy/MM/dd"))

        sqldr("RID") = drPPinfo("RID")
        sqldr("Relship") = relship
        sqldr("Content") = class_PName   '2007/11/19 修改成將訓練內容簡介的課程單元帶入--Charles
        sqldr("Years") = pp_years.Substring(2) '012
        sqldr("PlanID") = drPPinfo("PlanID")
        sqldr("ComIDNO") = drPPinfo("ComIDNO")
        sqldr("SeqNO") = drPPinfo("SeqNO")

        Dim vTPropertyID As String = "1" '1:在職 TPropertyID 訓練性質 0職前 1在職(進修) 2委託訓練
        sqldr("TPropertyID") = Val(vTPropertyID)
        sqldr("TMID") = drPPinfo("TMID")
        sqldr("CJOB_UNKEY") = drPPinfo("CJOB_UNKEY") '通俗職類

        sqldr("CLSID") = vCLSID
        sqldr("ClassEngName") = If(Convert.ToString(drPPinfo("ClassEngName")) <> "", drPPinfo("ClassEngName").ToString(), vCLASSENAME)
        sqldr("CLASSCNAME") = drPPinfo("ClassName")

        Dim vCyclType As String = TIMS.ClearSQM(drPPinfo("CyclType"))
        If vCyclType = "" Then vCyclType = TIMS.cst_Default_CyclType
        vCyclType = TIMS.FmtCyclType(vCyclType)
        sqldr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)

        sqldr("ClassNum") = Convert.DBNull '1 'vCyclType '班數
        sqldr("LevelCount") = 0 '無'(課程階段)
        sqldr("ISFullDate") = Convert.DBNull '全日制
        sqldr("CLASS_UNIT") = Convert.DBNull '"" '班級單元-Melody,for 學習卷
        sqldr("BGTime") = 0 '勾稽次數
        sqldr("ClassNum") = Convert.DBNull '1 'vCyclType '班數
        '訓練課程類型 ADVANCE
        sqldr("ADVANCE") = drPPinfo("ADVANCE")
        sqldr("TNum") = drPPinfo("TNum")
        sqldr("THours") = drPPinfo("THours")
        sqldr("Companyname") = Convert.DBNull '企業名稱

        sqldr("STDate") = drPPinfo("STDate")
        sqldr("FTDate") = drPPinfo("FDDate")
        'SELECT SENTERDATE,FENTERDATE,EXAMDATE,ExamPeriod FROM PLAN_PLANINFO WHERE ROWNUM  <=10
        'SELECT SENTERDATE,FENTERDATE,EXAMDATE,FENTERDATE2,ExamPeriod FROM CLASS_CLASSINFO  WHERE ROWNUM  <=10
        sqldr("SENTERDATE") = drPPinfo("SENTERDATE")
        sqldr("FENTERDATE") = drPPinfo("FENTERDATE")
        '甄試日期
        sqldr("EXAMDATE") = drPPinfo("EXAMDATE")
        '(甄試時段)
        sqldr("ExamPeriod") = drPPinfo("ExamPeriod")
        Dim sFENTERDATE As String = TIMS.Cdate3(drPPinfo("FENTERDATE"))
        Dim sEXAMDATE As String = TIMS.Cdate3(drPPinfo("EXAMDATE"))
        Dim SS1 As String = ""
        TIMS.SetMyValue(SS1, "RID1", vRID1) : TIMS.SetMyValue(SS1, "TPlanID", vTPlanID)
        Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, oConn)
        '報名登錄最晚 可作業時間
        sqldr("FENTERDATE2") = If(sFENTERDATE2 <> "", CDate(sFENTERDATE2), Convert.DBNull)
        '報到日期
        sqldr("CheckInDate") = drPPinfo("CheckInDate")
        '2005/8/11新增轉入訓練地點--Melody
        sqldr("TaddressZip") = drPPinfo("TaddressZip")
        sqldr("TaddressZIP6W") = drPPinfo("TaddressZIP6W")
        sqldr("TAddress") = drPPinfo("TAddress")

        '2005/8/12新增轉入課程目標--Melody，2007/9/26 修改成將訓練目標帶入即可--Charles
        'sqldr("Purpose") = "一、學科：" & drPlaninfo("PurScience") & "二、術科：" & drPlaninfo("PurTech")
        sqldr("Purpose") = drPPinfo("PurScience")
        sqldr("NotOpen") = "N"  '不開班原因代碼
        sqldr("NORID") = Convert.DBNull '不開班原因代碼
        sqldr("OtherReason") = Convert.DBNull '不開班其他原因說明
        sqldr("LastState") = "A" 'M: 修改(最後異動狀態)

        '班級英文名稱
        sqldr("CLASSENGNAME") = drPPinfo("CLASSENGNAME")
        '訓練時段'取得鍵值-訓練時段
        sqldr("TPERIOD") = drPPinfo("TPERIOD")
        sqldr("NOTE3") = drPPinfo("NOTE3")
        '「訓練期限」
        sqldr("TDEADLINE") = drPPinfo("TDEADLINE")
        '導師名稱
        sqldr("CTName") = drPPinfo("CTName")

        sqldr("EADDRESSZIP") = drPPinfo("EADDRESSZIP")
        sqldr("EADDRESSZIP6W") = drPPinfo("EADDRESSZIP6W")
        sqldr("EADDRESS") = drPPinfo("EADDRESS")

        sqldr("IsCalculate") = "N" '是否試算
        sqldr("IsClosed") = "N" '是否結訓
        sqldr("IsSuccess") = "Y" '是否轉入成功
        sqldr("IsApplic") = "N" '納入志願

        sqldr("ModifyAcct") = sm.UserInfo.UserID 'vModifyUserID 'sm.UserInfo.UserID
        sqldr("ModifyDate") = Now()
        'sqlTable.Rows.Add(sqldr)
        'sqldr_new("TransFlag") = "Y"
        'sqldr_new("ModifyAcct") = sm.UserInfo.UserID
        'sqldr_new("ModifyDate") = Now()

        'CLASS_CLASSINFO
        'MyPage.Session(cst_temp_classinfo) = sqlTable
        Dim flag_Checkok As Boolean = CheckSaveData06(sm, rErrMsg1, sqldr, oConn)
        If rErrMsg1 <> "" Then Return rErrMsg1 'Exit Sub 
        If Not flag_Checkok Then Return "檢核有誤轉入失敗" 'Exit Sub 

        '2019-02-19 add 記操作歷程（SYS_TRANS_LOG）'新增
        Dim htPP As New Hashtable From {
            {"TransType", TIMS.cst_TRANS_LOG_Insert},
            {"TargetTable", "CLASS_CLASSINFO"},
            {"FuncPath", "/TC/02/TC_02_001"},
            {"s_WHERE", String.Format("OCID='{0}'", iOCID_New)}
        }
        TIMS.SaveTRANSLOG(sm, oConn, sqldr, htPP)

        DbAccess.UpdateDataTable(sqlTable, oAdp)

        Dim parms_pp As New Hashtable From {
            {"ModifyAcct", sm.UserInfo.UserID},
            {"PlanID", vPlanID},
            {"COMIDNO", vCOMIDNO},
            {"SEQNO", vSEQNO}
        }
        Dim sql_pp As String = ""
        sql_pp &= " UPDATE PLAN_PLANINFO"
        sql_pp &= " SET TransFlag='Y',ModifyAcct=@ModifyAcct,ModifyDate=GETDATE()"
        sql_pp &= " WHERE PlanID=@PlanID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        DbAccess.ExecuteNonQuery(sql_pp, oConn, parms_pp)

        Dim parms_cc As New Hashtable From {
            {"PlanID", vPlanID},
            {"COMIDNO", vCOMIDNO},
            {"SEQNO", vSEQNO}
        }
        Dim sql_cc As String = ""
        sql_cc &= " SELECT OCID FROM CLASS_CLASSINFO"
        sql_cc &= " WHERE PlanID=@PlanID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        Dim rqOCID As String = DbAccess.ExecuteScalar(sql_cc, oConn, parms_cc)
        TIMS.Insert_Auth_AccRWClass(sm, rqOCID, -1, oConn)

        Return rErrMsg1
    End Function

    '[檢核-儲存前-檢核] CLASS_CLASSINFO - sqldr
    Public Shared Function CheckSaveData06(ByRef sm As SessionModel, ByRef Errmsg As String, ByRef sqldr As DataRow, ByRef oConn As SqlConnection) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""

        If Convert.ToString(sqldr("SEnterDate")) = "" Then Errmsg &= "[報名開始日期] 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("FEnterDate")) = "" Then Errmsg &= "[報名結束日期] 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("STDate")) = "" Then Errmsg &= "[開訓日期] 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("FTDate")) = "" Then Errmsg &= "[結訓日期] 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("ExamDate")) = "" Then Errmsg &= "[甄試日期] 資料有誤!" & vbCrLf
        If Errmsg <> "" Then Return False

        If (CDate(sqldr("SEnterDate")) >= CDate(sqldr("FEnterDate"))) Then Errmsg &= "[報名結束日期]必須大於[報名開始日期]!" & vbCrLf
        If (CDate(sqldr("STDate")) <= CDate(sqldr("FEnterDate"))) Then Errmsg &= "[開訓日期]必須大於[報名結束日期]!" & vbCrLf
        If Convert.ToString(sqldr("ExamDate")) <> "" Then
            If Convert.ToString(sqldr("ExamPeriod")) = "" Then '20100329 add 甄試時段
                Errmsg &= "「甄試時段」全天、上午、下午 時段請擇一選擇!" & vbCrLf
            End If
            If (CDate(sqldr("ExamDate")) <= CDate(sqldr("FEnterDate"))) Then
                Errmsg &= "「甄試日期」必須大於「報名結束日期」!" & vbCrLf
            End If
            If (CDate(sqldr("ExamDate")) > CDate(sqldr("STDate"))) Then
                Errmsg &= "[甄試日期]必須小於或等於[開訓日期]!" & vbCrLf
            End If
        End If
        If Errmsg <> "" Then Return False

        Dim vRID1 As String = sqldr("RID").ToString().Substring(0, 1)

        If Errmsg = "" Then
            'Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(sFENTERDATE, sEXAMDATE)
            Dim sFENTERDATE As String = TIMS.Cdate3(sqldr("FEnterDate"))
            Dim sEXAMDATE As String = TIMS.Cdate3(sqldr("ExamDate"))
            Dim SS1 As String = ""
            TIMS.SetMyValue(SS1, "RID1", vRID1) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
            Dim strFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, oConn)
            sqldr("FEnterDate2") = If(strFENTERDATE2 <> "", CDate(strFENTERDATE2), Convert.DBNull)

            If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                If strFENTERDATE2 = "" Then Errmsg &= "報名登錄最晚可作業時間 不可為空白，請確認報名結束日期正確性" & vbCrLf
            End If

            If Errmsg = "" AndAlso Not TIMS.Chk_FENTERDATE2(sFENTERDATE, strFENTERDATE2) Then
                Errmsg &= String.Format("報名登錄最晚可作業時間 報名結束日期 到 (報名結束日期+3)日曆天(ERROR:{0})", TIMS.STR2NUL(strFENTERDATE2)) & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Return False

        'If FEnterDate2.Text = "" Then
        'End If
        'If FEnterDate2.Text = "" Then Errmsg &= "報名登錄最晚可作業時間 不可為空白，請點選計算" & vbCrLf
        'If TB_QaySDate.Text = "" Then Errmsg &= "請選擇問卷調查開始日期" & vbCrLf
        'If TB_QayFDate.Text = "" Then Errmsg &= "請選擇問卷調查結束日期" & vbCrLf
        'If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '甄試日期 自辦在職 為必填
        '    If ExamDate.Text = "" Then Errmsg &= "請選擇甄試日期" & vbCrLf
        'End If
        If Convert.ToString(sqldr("TDeadline")) = "" Then Errmsg &= "訓練期限 資料有誤!" & vbCrLf

        If Convert.ToString(sqldr("TADDRESSZIP")) = "" Then Errmsg &= "上課地址 郵遞區號前3碼 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("TADDRESSZIP6W")) = "" Then Errmsg &= "上課地址 郵遞區號後2碼 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("TADDRESS")) = "" Then Errmsg &= "上課地址 地址 資料有誤!" & vbCrLf

        If Convert.ToString(sqldr("EADDRESSZIP")) = "" Then Errmsg &= "報名地點 郵遞區號前3碼 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("EADDRESSZIP6W")) = "" Then Errmsg &= "報名地點 郵遞區號後2碼 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("EADDRESS")) = "" Then Errmsg &= "報名地點 地址 資料有誤!" & vbCrLf

        If Convert.ToString(sqldr("TPeriod")) = "" Then Errmsg &= "訓練時段 資料有誤!" & vbCrLf
        '自辦在職顯示此功能 且為必填
        If Convert.ToString(sqldr("NOTE3")) = "" Then Errmsg &= "訓練時段 上課時間 資料有誤!" & vbCrLf
        If Convert.ToString(sqldr("CheckInDate")) = "" Then Errmsg &= "報到日期 資料有誤!" & vbCrLf
        If Errmsg <> "" Then Return False

        Dim vCLSID As String = sqldr("CLSID").ToString()
        ',PlanID,COMIDNO,SEQNO
        Dim vPlanID As String = sqldr("PlanID").ToString()
        Dim vCOMIDNO As String = sqldr("COMIDNO").ToString()
        Dim vSEQNO As String = sqldr("SEQNO").ToString()
        'CyclType
        Dim vCyclType As String = sqldr("CyclType").ToString()
        Dim vCLASSCNAME As String = Convert.ToString(sqldr("CLASSCNAME"))
        Dim vRID As String = sqldr("RID").ToString()

        Dim parms_B As New Hashtable From {
            {"CLSID", vCLSID},
            {"PlanID", vPlanID},
            {"CyclType", vCyclType},
            {"CLASSCNAME", vCLASSCNAME},
            {"RID", vRID}
        }
        Dim sqlstr_B As String = ""
        sqlstr_B &= " SELECT 1 FROM CLASS_CLASSINFO"
        sqlstr_B &= " WHERE NOTOPEN='N' AND CLSID=@CLSID AND PlanID=@PlanID AND ISNULL(CyclType,'')=@CyclType AND CLASSCNAME=@CLASSCNAME AND RID=@RID"
        Dim dt1B As DataTable = DbAccess.GetDataTable(sqlstr_B, oConn, parms_B)
        If dt1B.Rows.Count > 0 Then '>0資料重複
            Errmsg &= "開班資料重複!!!!" & vbCrLf
            Return False
        End If

        Dim parms_C As New Hashtable From {
            {"PLANID", vPlanID},
            {"COMIDNO", vCOMIDNO},
            {"SEQNO", vSEQNO}
        }
        Dim sqlstr_C As String = ""
        sqlstr_C &= " SELECT 1 FROM CLASS_CLASSINFO WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
        Dim dt1C As DataTable = DbAccess.GetDataTable(sqlstr_C, oConn, parms_C)
        If dt1C.Rows.Count > 0 Then '>0資料重複
            Errmsg &= "開班資料重複!!!" & vbCrLf
            Return False
        End If

        Return rst
    End Function

    ''' <summary> 轉班上架 </summary>
    ''' <param name="htSS"></param>
    Function Utl_EntreShelf2(ByRef htSS As Hashtable) As String
        Dim vPCS As String = TIMS.GetMyValue2(htSS, "PCS")
        'Dim vCLSID As String = TIMS.GetMyValue2(htSS, "CLSID")
        Dim vCJOB_UNKEY As String = TIMS.GetMyValue2(htSS, "CJOB_UNKEY")
        Dim vPlanID As String = TIMS.GetMyValue2(htSS, "PlanID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(htSS, "COMIDNO")
        Dim vSEQNO As String = TIMS.GetMyValue2(htSS, "SEQNO")
        Dim vTPlanID As String = TIMS.GetMyValue2(htSS, "TPlanID")
        Dim vRID1 As String = TIMS.GetMyValue2(htSS, "RID1")

        If vPCS = "" Then Return "傳入參數有誤!"
        If vCJOB_UNKEY = "" Then Return "傳入參數有誤!"
        If vPlanID = "" Then Return "傳入參數有誤!"
        If vCOMIDNO = "" Then Return "傳入參數有誤!"
        If vSEQNO = "" Then Return "傳入參數有誤!"
        If vTPlanID = "" Then Return "傳入參數有誤!"
        If vRID1 = "" Then Return "傳入參數有誤!"

        'https://localhost:44383/TC/01/TC_01_004_classid.aspx?pp=cc&TMID=3113&PlanID=5018
        'View-source: https : //localhost:44383/TC/01/TC_01_004_InsertPlan.aspx?planid=5018&ComIDNO=00500027&SeqNO=1&ProcessType=Update&ID=246&RID=B
        '計畫訓練內容簡介無資料， 不可轉入!!

        '1.產生/取得一組班別代碼
        Dim vCLSID As String = GET_IDCLASS_CLSID(objconn, vPCS)
        If vCLSID = "" Then Return "傳入參數有誤!"

        '2.檢核資料是否可轉入 / '3.自動轉班、上架
        Dim parms As New Hashtable From {
            {"CLSID", vCLSID},
            {"CJOB_UNKEY", vCJOB_UNKEY},
            {"PCS", vPCS},
            {"PlanID", vPlanID},
            {"COMIDNO", vCOMIDNO},
            {"SEQNO", vSEQNO},
            {"TPlanID", vTPlanID},
            {"RID1", vRID1},
            {"ModifyUserID", sm.UserInfo.UserID}
        }

        ' 轉入資料(SAVE) PLAN_PLANINFO CLASS_CLASSINFO - 自辦在職
        Dim rErrMsg1 As String = SAVE_CHANGEDATA_06(sm, parms, objconn)
        Return rErrMsg1
    End Function

    '需求編號：OJT-21032502、OJT-21032504、OJT-21032505
    ''' <summary>
    ''' 轉班上架-批次轉班上架
    ''' </summary>
    Sub Utl_EntreShelf1()
        Dim i_checkitem As Integer = 0
        For Each eItem As DataGridItem In dtPlan.Items
            'Dim drv As DataRowView = eItem.DataItem
            Dim Hid_PCS As HiddenField = eItem.FindControl("Hid_PCS") '選取id
            Dim chkItem As HtmlInputCheckBox = eItem.FindControl("chkItem") '選取
            If Not chkItem.Disabled AndAlso chkItem.Checked Then i_checkitem += 1
        Next
        '查無點選資料離開
        If i_checkitem = 0 Then
            Common.MessageBox(Me, "請勾選要批次轉班上架的班級!")
            Return
        End If

        '按鈕動作：當按下去時，系統自動產生一組班別代碼，並自動轉班、上架 (上架日期即設為當下)。
        For Each eItem As DataGridItem In dtPlan.Items
            'Dim drv As DataRowView = eItem.DataItem
            Dim Hid_PCS As HiddenField = eItem.FindControl("Hid_PCS") '選取id
            Dim chkItem As HtmlInputCheckBox = eItem.FindControl("chkItem") '選取
            Dim flag_create_CCINFO As Boolean = False
            If Not chkItem.Disabled AndAlso chkItem.Checked Then flag_create_CCINFO = True
            If flag_create_CCINFO Then
                Dim vPCS As String = TIMS.ClearSQM(Hid_PCS.Value)
                Dim drPP As DataRow = Get_PLANINFOdata(objconn, vPCS)
                If drPP Is Nothing Then
                    Common.MessageBox(Me, "傳入參數有誤!")
                    Return ' Exit Sub
                End If

                'vTMID = drPP("TMID").ToString()
                Dim vCJOB_UNKEY As String = drPP("CJOB_UNKEY").ToString()
                Dim vPlanID As String = drPP("PlanID").ToString()
                Dim vCOMIDNO As String = drPP("COMIDNO").ToString()
                Dim vSEQNO As String = drPP("SEQNO").ToString()
                Dim vTPlanID As String = drPP("TPlanID").ToString()
                Dim vRID As String = drPP("RID").ToString()
                Dim vRID1 As String = vRID.Substring(0, 1)

                Dim parms As New Hashtable From {
                    {"PCS", vPCS},
                    {"CJOB_UNKEY", vCJOB_UNKEY},
                    {"PlanID", vPlanID},
                    {"COMIDNO", vCOMIDNO},
                    {"SEQNO", vSEQNO},
                    {"TPlanID", vTPlanID},
                    {"RID1", vRID1}
                }
                '多筆-批次轉班上架/按鈕單筆-轉班上架
                Dim ErrMessage As String = Utl_EntreShelf2(parms)
                If ErrMessage <> "" Then
                    Common.MessageBox(Me, ErrMessage)
                    Return ' Exit Sub
                End If
            End If
        Next

        Dim s_OkMessage As String = "班級轉入成功!"
        If i_checkitem > 1 Then s_OkMessage = String.Format("班級轉入成功!共有{0}筆資料!", i_checkitem)
        Common.MessageBox(Me, s_OkMessage)
        Call SearchData1()
    End Sub

    ''' <summary>
    ''' 批次轉班上架
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnEnter1_Click(sender As Object, e As EventArgs) Handles btnEnter1.Click
        Call Utl_EntreShelf1()
    End Sub

End Class
