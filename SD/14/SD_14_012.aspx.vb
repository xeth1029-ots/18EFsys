Partial Class SD_14_012
    Inherits AuthBasePage

    'SD_14_012_1_2013
    'Const cst_printFN1 As String="SD_14_012_1_2013"
    'Const cst_printFN1 As String="SD_14_012_1_2020"
    Const cst_printFN1 As String = "SD_14_012_1_2021"

    '2010 年改成產學訓跟在職勞工印同一格式
    'SD_14_012_1_2013  參訓學員補助經費申請表
    'FROM Stud_SubSidyCost
    'Dim gTestflag As Boolean=False '測試
    Const cst_errMsg1 As String = "請先填寫完所有學員的補助申請,再列印支付參訓學員補助經費申請表!"
    'Const cst_errMsg2 As String="查無資料!! 請先填寫完學員的 支付參訓學員補助經費申請表,再列印此表單!!"
    Const cst_errMsg2 As String = "查無資料!! 請先完成補助申請,再列印支付參訓學員補助經費申請表!!"

    Dim sMemo As String = "" '(查詢原因)
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

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
        PageControler1.PageDataGrid = DataGrid1

        'Dim gTestflag As Boolean=False '測試
        'If TIMS.Utl_GetConfigSet("TestTPID54x2012")="Y" Then gTestflag=True '測試
        'hidgTestflag.Value=TIMS.Utl_GetConfigSet("TestTPID54x2012")

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

            'print_type.Attributes("onclick")="printkind();" '列印時排序方式
            Button1.Attributes("onclick") = "return CheckSearch();"
            Button5.Attributes("onclick") = "return CheckPrint();"
        End If

        HidYears.Value = sm.UserInfo.Years
        Years.Value = sm.UserInfo.Years - 1911
        PlanID.Value = sm.UserInfo.PlanID

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

    End Sub
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
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        SelectValue.Value = ""

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim SearchStr As String = ""
        If STDate1.Text <> "" Then
            SearchStr &= " AND cc.STDate >= " & TIMS.To_date(IIf(flag_ROC, TIMS.Cdate18(STDate1.Text), STDate1.Text)) & vbCrLf  'edit，by:20181023
        End If
        If STDate2.Text <> "" Then
            SearchStr &= " AND cc.STDate <= " & TIMS.To_date(IIf(flag_ROC, TIMS.Cdate18(STDate2.Text), STDate2.Text)) & vbCrLf  'edit，by:20181023
        End If
        If FTDate1.Text <> "" Then
            SearchStr &= " AND cc.FTDate >= " & TIMS.To_date(IIf(flag_ROC, TIMS.Cdate18(FTDate1.Text), FTDate1.Text)) & vbCrLf  'edit，by:20181023
        End If
        If FTDate2.Text <> "" Then
            SearchStr &= " AND cc.FTDate <= " & TIMS.To_date(IIf(flag_ROC, TIMS.Cdate18(FTDate2.Text), FTDate2.Text)) & vbCrLf  'edit，by:20181023
        End If
        If tr_AppStage_TP28.Visible Then
            Dim v_AppStage As String = TIMS.GetListValue(AppStage)
            If (v_AppStage <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage
            If v_AppStage <> "" AndAlso v_AppStage > "0" Then SearchStr &= " AND pp.AppStage='" & v_AppStage & "'" & vbCrLf
        End If

        Dim sql As String = ""
        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.STDATE ,cc.FTDATE ,ar.ORGNAME" & vbCrLf
        sql &= " ,ISNULL(c.ctotal,0) ctotal" & vbCrLf
        sql &= " FROM class_classinfo cc" & vbCrLf
        sql &= " JOIN plan_planinfo pp ON pp.planid=cc.planid AND pp.comidno=cc.comidno AND pp.seqno=cc.seqno" & vbCrLf
        sql &= " JOIN id_plan ip ON ip.PlanID=cc.PlanID" & vbCrLf
        sql &= " JOIN view_RIDName ar ON ar.RID=cc.RID" & vbCrLf
        sql &= " LEFT JOIN ( "
        sql &= "   SELECT cs.ocid ,COUNT(1) ctotal" & vbCrLf
        sql &= "   FROM Stud_SubSidyCost c" & vbCrLf
        sql &= "   JOIN class_studentsofclass cs ON cs.socid=c.socid" & vbCrLf
        sql &= "   JOIN class_classinfo cc ON cc.ocid=cs.ocid" & vbCrLf
        sql &= "   JOIN id_plan ip ON ip.planid=cc.planid" & vbCrLf
        sql &= "   WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        sql &= " GROUP BY cs.ocid" & vbCrLf
        sql &= " ) c ON c.ocid=cc.ocid "
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cc.notopen='N'" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If

        '28:產業人才投資方案
        orgid.Value = ""
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If PlanPoint.SelectedValue = "1" Then
                '產業人才投資計畫
                sql &= " AND ar.ORGKIND <> '10'" & vbCrLf
                orgid.Value = "G"
            Else
                '提升勞工自主學習計畫
                sql &= " AND ar.ORGKIND='10'" & vbCrLf
                orgid.Value = "W"
            End If
        End If
        sql += SearchStr
        sql &= " AND EXISTS (SELECT 'x' FROM AUTH_RELSHIP x WHERE x.RID=ar.RID AND x.RELSHIP LIKE '" & RelShip & "%')" & vbCrLf
        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        '查詢原因
        'Dim v_INQUIRY As String=TIMS.GetListValue(ddl_INQUIRY_Sch)
        'Dim v_rblWorkMode As String=TIMS.GetListValue(rblWorkMode)
        sMemo = GET_SEARCH_MEMO()
        sMemo &= $"&sql={sql}"
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "OCID,ORGNAME,CLASSCNAME,STDATE,FTDATE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Case ListItemType.Header
        '    e.Item.CssClass="SD_TD1"
        'If e.Item.ItemType=ListItemType.Item Then e.Item.CssClass="SD_TD2"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HtmlInputCheckBox = e.Item.FindControl("OCID")
                If drv("ctotal") = 0 Then
                    OCID.Attributes("onclick") = "this.checked=false;alert('" & cst_errMsg2 & "');"
                Else
                    OCID.Value = drv("OCID")
                    OCID.Attributes("onclick") = "SelectItem(this.checked,'" & drv("OCID") & "');"
                End If
                Dim OCIDArray As Array = Split(SelectValue.Value, ",")
                For i As Integer = 0 To OCIDArray.Length - 1
                    If drv("OCID").ToString = OCIDArray(i) Then OCID.Checked = True
                Next
                '20181023 切換[開訓日期]、[結訓日期](民國年/西元年)日期格式
                e.Item.Cells(3).Text = IIf(flag_ROC, TIMS.Cdate17(drv("STDATE")), drv("STDATE"))  'edit，by:20181023
                e.Item.Cells(4).Text = IIf(flag_ROC, TIMS.Cdate17(drv("FTDATE")), drv("FTDATE"))  'edit，by:20181023
        End Select
    End Sub

    Private Sub printall_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles printall.ServerClick
        Dim sql As String = ""
        Dim dt3Q As DataTable = CheckinputQ("")  '判斷申請補助的學員是否有填寫受訓意見調查表

        If dt3Q.Rows.Count <> 0 Then
            For i As Integer = 0 To dt3Q.Rows.Count - 1
                '補助申請人數不等於填寫人數
                If Convert.ToString(dt3Q.Rows(i)("ctotal")) <> Convert.ToString(dt3Q.Rows(i)("qtotal")) Then
                    'Common.MessageBox(Me, "請先填寫完所有申請補助學員的受訓意見調查表,再列印此表單!")
                    Common.MessageBox(Me, cst_errMsg1)
                    Exit Sub
                End If
            Next
        End If

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        'JAVASCRIPT 要注意
        'print_orderyby.Value="d.IDNO"
        'If print_type.SelectedValue=2 Then print_orderyby.Value="c.StudentID" ' "R.StudentID"
        print_orderyby.Value = "1" '"d.IDNO"
        If print_type.SelectedValue = 2 Then print_orderyby.Value = "2" '"c.StudentID"

        Dim MyValue1 As String = ""
        MyValue1 &= "&Years=" & Years.Value
        MyValue1 &= "&RelShip=" & RelShip
        MyValue1 &= "&Years2=" & Convert.ToString(CInt(Years.Value) + 1911)
        'MyValue &= "&OCID=" & SelectValue.Value
        MyValue1 &= "&RID=" & RIDValue.Value
        MyValue1 &= "&Printtype=" & print_orderyby.Value
        'Printtype1
        'Printtype2
        Select Case print_orderyby.Value
            Case "1"
                MyValue1 &= "&Printtype1=Y"
            Case Else
                MyValue1 &= "&Printtype2=Y"
        End Select

        sMemo = $"{MyValue1}&prt={cst_printFN1}"
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm列印, 2, "", sMemo)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)
    End Sub

    '判斷申請補助的學員是否有填寫受訓意見調查表
    Function sUtl_ChkData1(ByRef sMsg As String) As Boolean
        Dim rst As Boolean = True 'ok為True 異常為False
        sMsg = ""
        If SelectValue.Value = "" Then
            sMsg &= "請選擇有效班級資料!" & vbCrLf
            Return False
        End If

        'ISprint: 預設=1表示可以印報表,,'2表示不能印,如果有申請補助的學員沒有寫意見調查表就不能印報表
        ISprint.Value = 1
        '判斷申請補助的學員是否有填寫受訓意見調查表
        Dim dt3Q As DataTable = CheckinputQ(SelectValue.Value)
        If dt3Q.Rows.Count <> 0 Then
            For i As Integer = 0 To dt3Q.Rows.Count - 1
                '補助申請人數不等於填寫人數
                If Convert.ToString(dt3Q.Rows(i)("ctotal")) <> Convert.ToString(dt3Q.Rows(i)("qtotal")) Then
                    'Common.MessageBox(Me, "請先填寫完所有申請補助學員的受訓意見調查表,再列印此表單!")
                    'sMsg &= "請先填寫完所有申請補助學員的受訓意見調查表,再列印此表單!" & vbCrLf
                    sMsg &= cst_errMsg1 & vbCrLf 'Common.MessageBox(Me, cst_errMsg1)
                    ISprint.Value = 2 '2表示不能印,如果有申請補助的學員沒有寫意見調查表就不能印報表
                    Return False
                    'Exit Function
                End If
            Next
        End If

        Return rst
    End Function

    '判斷申請補助的學員是否有填寫受訓意見調查表
    Function CheckinputQ(ByVal OCIDs As String) As DataTable
        Dim sql2 As String = ""
        sql2 &= " SELECT cc.ocid" & vbCrLf
        sql2 &= " ,COUNT(1) stotal" & vbCrLf '所有學員數
        sql2 &= " ,COUNT(ss.socid) ctotal" & vbCrLf '補助申請人數
        'sql2 += ",COUNT(ISNULL(sq.socid,sq2.socid)) qtotal" & vbCrLf  '填寫問卷數
        sql2 &= " ,COUNT(CASE WHEN dbo.FN_GET_GOVCNT(ss.socid)=1 THEN 1 END) qtotal" & vbCrLf  '填寫問卷數
        sql2 &= " FROM Plan_PlanInfo pp" & vbCrLf
        sql2 &= " JOIN CLASS_CLASSINFO cc ON pp.planid=cc.planid AND pp.comidno=cc.comidno AND pp.seqno=cc.seqno" & vbCrLf
        sql2 &= " JOIN CLASS_STUDENTSOFCLASS cs ON cc.ocid=cs.ocid" & vbCrLf
        sql2 &= " LEFT JOIN STUD_SUBSIDYCOST ss ON cs.socid=ss.socid" & vbCrLf
        'sql2 += " LEFT JOIN STUD_QUESTIONFAC sq ON ss.socid=sq.socid" & vbCrLf
        'sql2 += " LEFT JOIN STUD_QUESTIONFAC2 sq2 ON ss.socid=sq2.socid" & vbCrLf
        sql2 &= " WHERE cc.rid='" & sm.UserInfo.RID & "' AND pp.planid='" & sm.UserInfo.PlanID & "'" & vbCrLf
        If OCIDs <> "" Then OCIDs = Trim(OCIDs)
        If OCIDs <> "" Then sql2 &= $" AND cc.OCID IN ({OCIDs})" & vbCrLf
        sql2 &= " GROUP BY cc.ocid "

        Dim dt3Q As DataTable = Nothing

        Try
            dt3Q = DbAccess.GetDataTable(sql2, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg += "/* sql2: */" & vbCrLf
            strErrmsg += sql2 & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

        Return dt3Q
    End Function

    '列印
    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        SelectValue.Value = TIMS.Get_SelectValue(DataGrid1, "OCID")
        SelectValue.Value = TIMS.CombiSQLINM3(SelectValue.Value)
        'Dim V_SELECTVALUE As String = SelectValue.Value
        Dim sMsg As String = ""
        If Not sUtl_ChkData1(sMsg) Then
            Common.MessageBox(Me, sMsg)
            Exit Sub
        End If

        '&Years=' + document.getElementById('Years').value + '&OCID=' + document.getElementById('SelectValue').value + '&Printtype=' + document.getElementById('print_orderyby').value + '&RID=' + document.getElementById('RIDValue').value);
        'JAVASCRIPT 要注意
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        print_orderyby.Value = If(print_type.SelectedValue = "2", "2", "1") '2:c.StudentID"'1:d.IDNO"

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If
        Dim pParms As New Hashtable From {{"RID", RIDValue.Value}}
        Dim SSQL As String = ""
        SSQL &= " SELECT a.OCID,c.SOCID,c.NAME,c.IDNO,c.BIRTHDAY" & vbCrLf
        SSQL &= " ,concat(CASE WHEN h.AcctMode=0 THEN h.PostNo WHEN h.AcctMode=1 THEN concat(h.AcctHeadNo, CASE WHEN h.AcctExNo IS NOT NULL THEN concat('-',h.AcctExNo) END) END ,' ',h.AcctNo) BANKCODE" & vbCrLf
        SSQL &= " ,h.EXBANKNAME ,CONVERT(NUMERIC,ISNULL(b.Defstdcost+b.DefGovCost,0)/ISNULL(b.TNum,1)) OPRICE" & vbCrLf
        SSQL &= " ,CONVERT(NUMERIC,CASE WHEN c.SUPPLYID='1' AND c.budgetid <> '97' THEN ISNULL(b.Defstdcost+b.DefGovCost,0)/ISNULL(b.TNum,1)*0.8 ELSE ISNULL(b.Defstdcost+b.DefGovCost,0)/ISNULL(b.TNum,1) END) PERPAY" & vbCrLf
        SSQL &= " ,CASE WHEN c.SUPPLYID='1' THEN '' ELSE g.Name END + CASE WHEN c.BudgetID='97' THEN '(協助基金)' ELSE '' END IDENTITYNAME" & vbCrLf
        SSQL &= " FROM CLASS_CLASSINFO a" & vbCrLf
        SSQL &= " JOIN PLAN_PLANINFO b WITH(NOLOCK) ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNo=b.SeqNo" & vbCrLf
        SSQL &= " JOIN V_STUDENTINFO c ON a.OCID=c.OCID" & vbCrLf
        SSQL &= " JOIN KEY_IDENTITY g WITH(NOLOCK) ON g.IdentityID=c.MIdentityID" & vbCrLf
        SSQL &= " JOIN STUD_SERVICEPLACE h WITH(NOLOCK) ON h.SOCID=c.SOCID" & vbCrLf
        SSQL &= " JOIN STUD_SUBSIDYCOST t WITH(NOLOCK) ON t.SOCID=c.SOCID" & vbCrLf
        SSQL &= " WHERE a.AppliedResultR='Y' AND c.StudStatus NOT IN (2,3)" & vbCrLf
        SSQL &= $" AND a.RID=@RID AND a.OCID IN ({SelectValue.Value})" & vbCrLf
        SSQL &= " ORDER BY c.IDNO" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(SSQL, objconn, pParms)
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        sMemo = GET_SEARCH_MEMO()
        'Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "OCID,SOCID,NAME,IDNO,BIRTHDAY,BANKCODE,EXBANKNAME,OPRICE,PERPAY,IDENTITYNAME")
        'Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm列印, TIMS.cst_wmdip2, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        Dim MyValue1 As String = $"&Years={Years.Value}&OCID={SelectValue.Value}&RID={RIDValue.Value}&Printtype={print_orderyby.Value}"
        'Printtype1 'Printtype2
        Select Case print_orderyby.Value
            Case "1"
                MyValue1 &= "&Printtype1=Y"
            Case Else
                MyValue1 &= "&Printtype2=Y"
        End Select

        'ISprint: 預設=1表示可以印報表,,'2表示不能印,如果有申請補助的學員沒有寫意見調查表就不能印報表
        If ISprint.Value <> "1" Then
            'sMsg &= "請先填寫完所有申請補助學員的受訓意見調查表,再列印此表單!" & vbCrLf 'Common.MessageBox(Me, sMsg)
            Common.MessageBox(Me, cst_errMsg1)
            Exit Sub
        End If

        sMemo &= $"{MyValue1}&prt={cst_printFN1}"
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "OCID,SOCID,NAME,IDNO,BIRTHDAY,BANKCODE,EXBANKNAME,OPRICE,PERPAY,IDENTITYNAME")
        Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm列印, TIMS.cst_wmdip2, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)

    End Sub

End Class