Partial Class SD_14_011
    Inherits AuthBasePage

    'Const cst_printFN1 As String="SD_14_011_1_2013_b"
    Const cst_printFN1 As String = "SD_14_011_1_2020_b"

    '/**NEW 2014**/,'SD_14_011_1_2013_b,'預估參訓學員補助經費清冊
    '2010年改成產學訓與在職勞工同一格式,'SD_14_011_1_2009 @BussinessTrain,'SD_14_011_1_2013(停用),

    'Dim iPYNum14 As Integer=1 'iPYNum14=TIMS.sUtl_GetPYNum14(Me)
    Dim prtFilename As String = "" '列印表件名稱
    'Dim gflag_test As Boolean=False '測試
    Const cst_errMsg1 As String = "請完成[學員資料維護]的[學員資料確認]與[學員資料審核]，才可使用此列印功能!!!"

    Dim sMemo As String = "" '(查詢原因)
    Dim OCIDArray As Array
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

        PageControler1.PageDataGrid = DataGrid1
        'iPYNum14=TIMS.sUtl_GetPYNum14(Me)
        'Hidorder1.Value="d.IDNO"
        'Hidorder2.Value="c.StudentID"

        'gflag_test=TIMS.sUtl_ChkTest()
        'Dim gflag_test As Boolean=False '測試
        'If TIMS.Utl_GetConfigSet("TestTPID54x2012")="Y" Then gTestflag=True '測試
        OCIDArray = Nothing
        If SelectValue.Value <> "" Then OCIDArray = Split(SelectValue.Value, ",")

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "1")

            '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage=TIMS.Get_AppStage(AppStage)
            If tr_AppStage_TP28.Visible Then
                AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
                TIMS.SET_MY_APPSTAGE_LIST_VAL(Me, AppStage)
            End If
        End If
        Years.Value = sm.UserInfo.Years - 1911
        PlanID.Value = sm.UserInfo.PlanID

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        'print_type.Attributes("onclick")="printkind();" '列印時排序方式
        Button1.Attributes("onclick") = "return CheckSearch();"
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = ""
                End If
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HtmlInputCheckBox = e.Item.FindControl("OCID")
                Dim AppliedResultR As HtmlInputHidden = e.Item.FindControl("AppliedResultR")
                Dim j As Integer = 0
                OCID.Value = drv("OCID")
                OCID.Attributes("onclick") = "SelectItem(this.checked,this.value);"
                'Dim OCIDArray As Array=Split(SelectValue.Value, ",")
                If Not OCIDArray Is Nothing Then
                    For i As Integer = 0 To OCIDArray.Length - 1
                        If drv("OCID").ToString = OCIDArray(i) Then OCID.Checked = True
                    Next
                End If
        End Select
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""
        If Trim(STDate1.Text) <> "" Then
            STDate1.Text = Trim(STDate1.Text)
            If Not TIMS.IsDate1(STDate1.Text) Then Errmsg += "開訓期間 起始日期格式有誤" & vbCrLf
            If Errmsg = "" Then STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
        Else
            STDate1.Text = ""
        End If
        If Trim(STDate2.Text) <> "" Then
            STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then Errmsg += "開訓期間 迄止日期格式有誤" & vbCrLf
            If Errmsg = "" Then STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
        Else
            STDate2.Text = ""
        End If
        If Trim(FTDate1.Text) <> "" Then
            FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then Errmsg += "結訓期間 起始日期格式有誤" & vbCrLf
            If Errmsg = "" Then FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
        Else
            FTDate1.Text = ""
        End If
        If Trim(FTDate2.Text) <> "" Then
            FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then Errmsg += "結訓期間 迄止日期格式有誤" & vbCrLf
            If Errmsg = "" Then FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
        Else
            FTDate2.Text = ""
        End If
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function
    '查詢原因
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        'center,RIDValue,Button2,HistoryList2,HistoryRID,STDate1,STDate2,FTDate1,FTDate2,tr_AppStage_TP28,AppStage,TRPlanPoint28,PlanPoint,
        '訓練機構,開訓期間,結訓期間,申請階段,
        center.Text = TIMS.ClearSQM(center.Text)
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text)
        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text)
        Dim v_AppStage As String = TIMS.GetListValue(AppStage) 'AppStage.SelectedValue

        If center.Text <> "" Then RstMemo &= String.Concat("&訓練機構=", center.Text)
        If STDate1.Text <> "" Then RstMemo &= String.Concat("&開訓期間1=", STDate1.Text)
        If STDate2.Text <> "" Then RstMemo &= String.Concat("&開訓期間2=", STDate2.Text)
        If FTDate1.Text <> "" Then RstMemo &= String.Concat("&結訓期間1=", FTDate1.Text)
        If FTDate2.Text <> "" Then RstMemo &= String.Concat("&結訓期間2=", FTDate2.Text)
        If v_AppStage <> "" Then RstMemo &= String.Concat("&申請階段=", v_AppStage)
        Return RstMemo
    End Function
    '查詢
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        SelectValue.Value = ""

        Dim sql As String = ""
        sql &= " SELECT cc.OCID ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) ClassCName" & vbCrLf
        sql &= " ,cc.STDate ,cc.FTDate" & vbCrLf
        sql &= " ,rr.OrgName ,cc.appliedresultR" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.planid=cc.planid AND pp.comidno=cc.comidno AND pp.seqno=cc.seqno" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME rr ON rr.RID=cc.RID" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
        sql &= " WHERE rr.RelShip LIKE '" & RelShip & "%'" & vbCrLf

        If sm.UserInfo.RID = "A" Then
            'change by nick ~ take ")" away from string in >> sm.UserInfo.TPlanID & "'* and Years<<< 
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        If STDate1.Text <> "" Then sql += " AND cc.STDate >= CONVERT(DATETIME, '" & STDate1.Text & "', 111)" & vbCrLf
        If STDate2.Text <> "" Then sql += " AND cc.STDate <= CONVERT(DATETIME, '" & STDate2.Text & "', 111)" & vbCrLf
        If FTDate1.Text <> "" Then sql += " AND cc.FTDate >= CONVERT(DATETIME, '" & FTDate1.Text & "', 111)" & vbCrLf
        If FTDate2.Text <> "" Then sql += " AND cc.FTDate <= CONVERT(DATETIME, '" & FTDate2.Text & "', 111)" & vbCrLf

        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'PlanPoint.SelectedValue
            Dim v_PlanPoint As String = TIMS.GetListValue(PlanPoint)
            If v_PlanPoint = "1" Then
                sql &= " AND rr.OrgKind <> '10'" & vbCrLf '產業人才投資計畫
                orgid.Value = "G"
            Else
                sql &= " AND rr.OrgKind='10'" & vbCrLf '提升勞工自主學習計畫
                orgid.Value = "W"
            End If
        Else
            orgid.Value = ""
        End If
        If tr_AppStage_TP28.Visible Then
            Dim v_AppStage As String = TIMS.GetListValue(AppStage) 'AppStage.SelectedValue
            If (v_AppStage <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage
            If v_AppStage <> "" Then sql += " AND pp.AppStage='" & v_AppStage & "'" & vbCrLf
        End If

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn)

        '查詢原因
        'Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        'Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode)
        sMemo = GET_SEARCH_MEMO()
        sMemo &= $"&sql={sql}"
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt2, "OCID,ORGNAME,CLASSCNAME,STDATE,FTDATE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, "", sMemo, objconn, v_INQUIRY, dt2.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt2) Then Return

        'sMemo &= $"&sql={sql}"
        'TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm查詢, 2, "", sMemo)

        DataGridTable.Visible = True
        msg.Text = ""
        'PageControler1.SqlString=sql
        PageControler1.PageDataTable = dt2
        PageControler1.ControlerLoad()
        'PageControler1.SqlPrimaryKeyDataCreate(sql, "OCID")
    End Sub

    ''' <summary>列印機構所有班級</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        'Dim dr As DataRow
        'Dim sql As String
        Dim strMsg As String = ""
        'print_type.SelectedValue 'SD_14_011_1_2013_b
        Dim v_print_type As String = TIMS.GetListValue(print_type)
        print_orderyby.Value = If(v_print_type = "2", "2", "1")  '2:by StudentID,'1:by IDNO

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then
            strMsg = "請選擇有效訓練機構/班級!"
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If

        Dim pms1 As New Hashtable From {{"RID", RIDValue.Value}}
        Dim sql1 As String = " SELECT RID,PLANID FROM CLASS_CLASSINFO WHERE RID=@RID"
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql1, objconn, pms1) '未選擇班級
        If dr1 Is Nothing Then
            strMsg = "請選擇有效訓練機構/班級!!"
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If
        RIDValue.Value = Convert.ToString(dr1("RID"))
        PlanID.Value = Convert.ToString(dr1("PLANID"))

        Dim flagCanPrint As Boolean = True '已完成學員資料審核 與 學員資料確認
        Dim pms2 As New Hashtable From {{"RID", RIDValue.Value}}
        Dim sql2 As String = " SELECT TOP 11 RID,PLANID FROM VIEW2 WHERE RID=@RID AND ISNULL(AppliedResultR,'N')<>'Y' "
        Dim dr2 As DataRow = DbAccess.GetOneRow(sql2, objconn, pms2)
        If dr2 IsNot Nothing Then flagCanPrint = False '未完成學員資料審核 與 學員資料確認
        If Not flagCanPrint Then
            '未完成學員資料審核 與 學員資料確認
            strMsg = cst_errMsg1 '"請完成[學員資料維護]的[學員資料確認]與[學員資料審核]，才可使用此列印功能!!!"
            Common.MessageBox(Me, strMsg)
            'If Not gflag_test Then Exit Sub
            Exit Sub
        End If

        '已完成學員資料審核 與 學員資料確認 'prtFilename="SD_14_011_1_2013_b" 'Printtype1 'Printtype2
        Dim MyValue1 As String = $"RID={RIDValue.Value}&PlanID={PlanID.Value}&Years={Years.Value}&Printtype={print_orderyby.Value}"
        Select Case print_orderyby.Value
            Case "1"
                MyValue1 &= "&Printtype1=Y"
            Case Else
                MyValue1 &= "&Printtype2=Y"
        End Select

        '已完成學員資料審核 與 學員資料確認
        'prtFilename="SD_14_011_1_2013_b"
        sMemo = $"{MyValue1}&prt={cst_printFN1}"
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm列印, 2, "", sMemo)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)

    End Sub

    '列印
    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'SelectValue.Value = TIMS.ClearSQM(SelectValue.Value)
        SelectValue.Value = TIMS.CombiSQLINM3(SelectValue.Value)

        Dim sql As String = ""
        Dim strMsg As String = ""

        'print_type.SelectedValue 'SD_14_011_1_2013_b
        Dim v_print_type As String = TIMS.GetListValue(print_type)
        print_orderyby.Value = If(v_print_type = "2", "2", "1")  '2:by StudentID,'1:by IDNO
        If SelectValue.Value = "" Then '未選擇班級
            strMsg = "請選擇訓練機構/班級!"
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If

        Dim V_SELECTVALUE As String = SelectValue.Value
        Select Case sm.UserInfo.LID
            Case 0
                sql = $"SELECT RID,PLANID FROM CLASS_CLASSINFO WHERE OCID IN ({V_SELECTVALUE}) "
            Case Else
                sql = $"SELECT RID,PLANID FROM CLASS_CLASSINFO WHERE OCID IN ({V_SELECTVALUE}) AND PLANID={sm.UserInfo.PlanID}"
        End Select
        Dim dtCC As DataTable = DbAccess.GetDataTable(sql, objconn) '未選擇班級
        If dtCC.Rows.Count = 0 Then
            strMsg = "請選擇有效訓練機構/班級!!"
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn) '未選擇班級
        If dr Is Nothing Then
            strMsg = "請選擇有效訓練機構/班級!!!"
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If
        RIDValue.Value = Convert.ToString(dr("RID"))
        PlanID.Value = Convert.ToString(dr("PLANID"))

        Dim flagCanPrint As Boolean = True '已完成學員資料審核 與 學員資料確認
        Dim pms2 As New Hashtable From {{"RID", RIDValue.Value}}
        sql = $" SELECT RID,PLANID FROM VIEW2 WHERE RID=@RID AND OCID IN ({SelectValue.Value}) AND ISNULL(AppliedResultR,'N')<>'Y'"
        dr = DbAccess.GetOneRow(sql, objconn, pms2)
        If dr IsNot Nothing Then flagCanPrint = False '未完成學員資料審核 與 學員資料確認
        If Not flagCanPrint Then
            '未完成學員資料審核 與 學員資料確認
            strMsg = cst_errMsg1 '"請完成[學員資料維護]的[學員資料確認]與[學員資料審核]，才可使用此列印功能!!!"
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If
        Dim pParms As New Hashtable From {{"RID", RIDValue.Value}, {"PlanID", PlanID.Value}}
        Dim SSQL As String = ""
        SSQL &= " SELECT a.OCID,c.SOCID,c.NAME,c.IDNO,c.BIRTHDAY" & vbCrLf
        SSQL &= " ,concat(CASE WHEN h.AcctMode=0 THEN h.PostNo WHEN h.AcctMode=1 THEN concat(h.AcctHeadNo, CASE WHEN h.AcctExNo IS NOT NULL THEN concat('-',h.AcctExNo) END) END ,' ',h.AcctNo) BANKCODE" & vbCrLf
        SSQL &= " ,h.EXBANKNAME ,CONVERT(NUMERIC,ISNULL(b.Defstdcost+b.DefGovCost,0)/ISNULL(b.TNum,1)) OPRICE" & vbCrLf
        SSQL &= " ,CONVERT(NUMERIC,CASE WHEN c.SUPPLYID='1' AND c.budgetid <> '97' THEN ISNULL(b.Defstdcost+b.DefGovCost,0)/ISNULL(b.TNum,1)*0.8 ELSE ISNULL(b.Defstdcost+b.DefGovCost,0)/ISNULL(b.TNum,1) END) PERPAY" & vbCrLf
        sSql &= " ,CASE WHEN c.SUPPLYID='1' THEN '' ELSE g.Name END + CASE WHEN c.BudgetID='97' THEN '(協助基金)' ELSE '' END IDENTITYNAME" & vbCrLf
        sSql &= " FROM CLASS_CLASSINFO a" & vbCrLf
        sSql &= " JOIN PLAN_PLANINFO b WITH(NOLOCK) ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNo=b.SeqNo" & vbCrLf
        sSql &= " JOIN V_STUDENTINFO c ON a.OCID=c.OCID" & vbCrLf
        sSql &= " LEFT JOIN KEY_IDENTITY g WITH(NOLOCK) ON g.IdentityID=c.MIdentityID" & vbCrLf
        sSql &= " LEFT JOIN STUD_SERVICEPLACE h WITH(NOLOCK) ON h.SOCID=c.SOCID" & vbCrLf
        sSql &= " WHERE a.AppliedResultR='Y' AND c.budgetid IS NOT NULL AND c.budgetid <> '99' AND ISNULL(c.AppliedResult,'Y')='Y'" & vbCrLf
        SSQL &= $" AND a.RID=@RID AND a.PlanID=@PlanID AND a.OCID IN ({V_SELECTVALUE})" & vbCrLf
        SSQL &= " ORDER BY c.IDNO" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(SSQL, objconn, pParms)
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        sMemo = GET_SEARCH_MEMO()
        'Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "OCID,SOCID,NAME,IDNO,BIRTHDAY,BANKCODE,EXBANKNAME,OPRICE,PERPAY,IDENTITYNAME")
        'Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm列印, TIMS.cst_wmdip2, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        '"Years=" & Years.Value & "&RID=" & RIDValue.Value & "&PlanID=" & PlanID.Value & "&Printtype=" & print_orderyby.Value
        '已完成學員資料審核 與 學員資料確認
        'prtFilename="SD_14_011_1_2013_b"
        Dim MyValue1 As String = $"RID={RIDValue.Value}&PlanID={PlanID.Value}&Years={Years.Value}&OCID={SelectValue.Value}&Printtype={print_orderyby.Value}"
        'Printtype1 'Printtype2
        Select Case print_orderyby.Value
            Case "1"
                MyValue1 &= "&Printtype1=Y"
            Case Else
                MyValue1 &= "&Printtype2=Y"
        End Select
        '已完成學員資料審核 與 學員資料確認
        sMemo &= $"{MyValue1}&prt={cst_printFN1}"
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "OCID,SOCID,NAME,IDNO,BIRTHDAY,BANKCODE,EXBANKNAME,OPRICE,PERPAY,IDENTITYNAME")
        Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm列印, TIMS.cst_wmdip2, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)

    End Sub
End Class