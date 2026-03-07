Partial Class SD_01_008
    Inherits AuthBasePage

    Dim CPdt As DataTable = Nothing
    'Dim ProcessType As String
    'Dim RelshipTable As DataTable
    Dim flagExportExcel As Boolean = False '匯出Excel

    Dim iClassCnt As Integer = 0     '班級數(總)
    Dim iStudCnt1 As Integer = 0     '訓練人數(總)
    Dim iStudCnt2 As Integer = 0     '報名人數(總)

    'Cells Columns
    'DG_ClassInfo 非產投 (TIMS)
    Const cst_dg1_序號 As Integer = 0
    'Const cst_dg1_訓練機構 As Integer = 1
    'Const cst_dg1_班別代碼 As Integer = 2
    Const cst_dg1_開結訓日 As Integer = 3
    'Const cst_dg1_班別名稱 As Integer = 4
    'Const cst_dg1_訓練職類 As Integer = 5
    'Const cst_dg1_訓練人數 As Integer = 6
    'Const cst_dg1_報名人數 As Integer = 7
    Const cst_dg1_甄試人數 As Integer = 8
    Const cst_dg1_開訓人數 As Integer = 9
    Const Cst_dg1_colspan As String = "8" '依dataGrid資料欄

    'DG_ClassInfo2 (產投)
    Const cst_dg2_序號 As Integer = 0
    Const cst_dg2_訓練機構 As Integer = 1
    'Const cst_dg2_課程代碼 As Integer = 2
    'Const cst_dg2_開結訓日 As Integer = 3
    'Const cst_dg2_班別名稱 As Integer = 4
    'Const cst_dg2_訓練業別 As Integer = 5
    'Const cst_dg2_訓練人數 As Integer = 6
    'Const cst_dg2_報名人數1 As Integer = 7
    'Const cst_dg2_報名人數2 As Integer = 8
    'Const cst_dg2_招生狀態 As Integer = 9
    'Const Cst_dg2_colspan28 As String = "10" '依dataGrid資料欄

    Const cst_sortOrgname As String = "orgname"

#Region "NO USE"
    ''傳進ocid值，傳回該班的報名人數 (扣除e網審核失敗)
    'Public Shared Function Get_EnterCount(ByVal OCID As String) As Integer
    '    Dim objstr As String
    '    '取得報名人數(網路)
    '    '扣除e網審核失敗
    '    objstr = ""
    '    objstr += " select ISNULL(count(se2.esetid),0) as total" & vbCrLf
    '    objstr += " from Stud_EnterTemp2 se2" & vbCrLf
    '    objstr += " join Stud_EnterType2 sy2 on se2.esetid = sy2.esetid" & vbCrLf
    '    objstr += " where sy2.ocid1 = '" & OCID & "' "
    '    objstr += " and ISNULL(sy2.signUpStatus,0) != 2 " '扣除e網審核失敗
    '    Return DbAccess.ExecuteScalar(objstr)
    'End Function
#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        'TIMS.TestDbConn(Me, objConn, True)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '分頁設定 Start
        msg2_28.Visible = False
        DG_Classinfo.Visible = False
        DG_Classinfo2.Visible = False

        '產投/非產投判斷
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投
            msg2_28.Visible = True
            Me.labtmid.Text = "訓練業別"
            DG_Classinfo2.Visible = True
            PageControler1.PageDataGrid = DG_Classinfo2
            btu_sel.Attributes("onclick") = "openTrain(document.getElementById('jobValue').value);"
        Else
            '非產投
            DG_Classinfo.Visible = True
            PageControler1.PageDataGrid = DG_Classinfo
            btu_sel.Attributes("onclick") = "openTrain(document.getElementById('trainValue').value);"
        End If
        '分頁設定 End

        'ProcessType = Request("ProcessType")
        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
            org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If
        TIMS.ShowHistoryRID(Me, historyrid, "HistoryList2", "RIDValue", "center")
        If historyrid.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not Page.IsPostBack Then
            btnexport1.Visible = False '預設要查詢一次再顯示匯出鍵
            msg.Text = ""
            table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Classinfo.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(cst_dg1_序號).Text = TIMS.Get_DGSeqNo(sender, e)
        End Select
    End Sub

    Private Sub DG_ClassInfo_SortCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DG_Classinfo.SortCommand
        ViewState("sort") = If(Not flagExportExcel, If(e.SortExpression = ViewState("sort"), String.Concat(e.SortExpression, " desc"), e.SortExpression), "")
        PageControler1.Sort = ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    '產投list
    Private Sub DG_ClassInfo2_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Classinfo2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If Not flagExportExcel Then
                    If ViewState("sort") IsNot Nothing Then
                        Dim i As Integer = -1
                        Select Case ViewState("sort")
                            Case cst_sortOrgname, String.Concat(cst_sortOrgname, " desc")
                                i = cst_dg2_訓練機構
                        End Select
                        If i > -1 Then
                            Dim img As New UI.WebControls.Image
                            img.ImageUrl = If(ViewState("sort").ToString.IndexOf("desc") = -1, "../../images/SortUp.gif", "../../images/SortDown.gif")
                            e.Item.Cells(i).Controls.Add(img)
                        End If
                    End If
                Else
                    e.Item.Cells(cst_dg2_訓練機構).Text = "訓練機構"
                    e.Item.Cells(cst_dg2_訓練機構).ForeColor = e.Item.Cells(cst_dg2_序號).ForeColor
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(cst_dg2_序號).Text = TIMS.Get_DGSeqNo(sender, e)

                If Not flagExportExcel Then
                    Dim Label99 As Label = e.Item.FindControl("Label99")
                    Dim ddlADMISSIONS As DropDownList = e.Item.FindControl("ddlADMISSIONS")
                    Dim hid_OCID As HiddenField = e.Item.FindControl("hid_OCID")
                    hid_OCID.Value = Convert.ToString(drv("OCID"))
                    '沒有下述狀況顯示 招生中
                    ddlADMISSIONS.Visible = False
                    Label99.Text = Convert.ToString(drv("ADMISSIONS_N"))
                    If DateDiff(DateInterval.Second, CDate(drv("today")), CDate(drv("SEnterDate"))) > 0 Then '過報名時間大於0
                        'Label99.Text = "尚未開始招生"
                    ElseIf DateDiff(DateInterval.Minute, CDate(drv("FEnterDate")), CDate(drv("today"))) > 0 Then
                        'Label99.Text = "已招生結束"
                    ElseIf Convert.ToString(drv("IsClosed")) = "Y" Then
                        'Label99.Text = "此班已結訓"
                    ElseIf Convert.ToString(drv("IsClosed")) = "N" _
                        AndAlso DateDiff(DateInterval.Minute, CDate(drv("today")), CDate(drv("FEnterDate"))) > 0 _
                        AndAlso Val(drv("TYPECNT1")) >= CInt(drv("TNum")) Then
                        Label99.Text = "" ' "報名額滿 (接受以備取身分報名)"
                        '招生狀態。'1:招生中'2:接受以備取報名'3:報名額滿。
                        ddlADMISSIONS.Visible = True
                        ddlADMISSIONS = TIMS.Get_ddlADMISSIONS(ddlADMISSIONS)
                        If Convert.ToString(drv("ADMISSIONS")) <> "" Then
                            Label99.Text = ""
                            Common.SetListItem(ddlADMISSIONS, Convert.ToString(drv("ADMISSIONS")))
                        Else
                            '預設為 2.接受以備取報名
                            Label99.Text = ""
                            Common.SetListItem(ddlADMISSIONS, "2")
                        End If
                    End If
                End If

                'If TestStr = "AmuTest" Then    '測試用
                '    If Convert.ToString(drv("IsClosed")) = "N" _
                '        AndAlso DateDiff(DateInterval.Minute, CDate(drv("today")), CDate(drv("FEnterDate"))) > 0 Then
                '        Label99.Text = "" ' "報名額滿 (接受以備取身分報名)"
                '        '招生狀態。'1:招生中'2:接受以備取報名'3:報名額滿。
                '        ddlADMISSIONS.Visible = True
                '        ddlADMISSIONS = TIMS.Get_ddlADMISSIONS(ddlADMISSIONS)
                '        If Convert.ToString(drv("ADMISSIONS")) <> "" Then
                '            Label99.Text = ""
                '            Common.SetListItem(ddlADMISSIONS, Convert.ToString(drv("ADMISSIONS")))
                '        Else
                '            '預設為 2.接受以備取報名
                '            Label99.Text = ""
                '            Common.SetListItem(ddlADMISSIONS, "2")
                '        End If

                '    End If
                'End If

        End Select

    End Sub

    Private Sub DG_ClassInfo2_SortCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DG_Classinfo2.SortCommand
        ViewState("sort") = If(Not flagExportExcel, If(e.SortExpression = ViewState("sort"), String.Concat(e.SortExpression, " desc"), e.SortExpression), "")
        PageControler1.Sort = ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '查詢 '非產投查詢 (現場?)
    Sub Search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DG_Classinfo)

        Dim vs_redate_start As String = If(redate_start.Text <> "", Common.FormatDate(CDate(redate_start.Text)), "")
        Dim vs_redate_end As String = If(redate_end.Text <> "", String.Concat(Common.FormatDate(CDate(redate_end.Text)), " 23:59:59"), "")

        Dim PlanKind As String = ""
        '依sm.UserInfo.PlanID取得PlanKind
        PlanKind = TIMS.Get_PlanKind(Me, objconn)
        Dim RWClassFlag As Boolean = False '自辦者只能列出賦予給此帳號的班級
        Select Case sm.UserInfo.RID
            Case "A"
            Case Else
                If sm.UserInfo.RoleID > 1 Then '非署 才有此限制
                    If PlanKind = "1" Then '自辦者只能列出賦予給此帳號的班級
                        RWClassFlag = True
                    End If
                End If
        End Select

        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select f.Years" & vbCrLf
        sql &= " ,f.CyclType" & vbCrLf
        sql &= " ,concat(f.YEARS,'0',ISNULL(g.CLASSID2,g.CLASSID),ISNULL(f.CYCLTYPE,'01')) OClassID" & vbCrLf
        sql &= " ,f.OCID" & vbCrLf
        sql &= " ,f.PlanID" & vbCrLf
        sql &= " ,f.ComIDNO" & vbCrLf
        sql &= " ,f.SeqNO" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(f.CLASSCNAME,f.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,f.TPropertyID" & vbCrLf
        sql &= " ,f.STDate" & vbCrLf
        sql &= " ,f.FTDate" & vbCrLf
        sql &= " ,f.RID" & vbCrLf
        sql &= " ,f.TNum" & vbCrLf
        sql &= " ,g.ClassID" & vbCrLf
        sql &= " ,e.OrgName" & vbCrLf
        sql &= " ,h.TrainName" & vbCrLf
        'sql &= " ,r3.OrgName2" & vbCrLf
        sql &= " FROM ID_Plan c" & vbCrLf
        sql &= " JOIN Class_ClassInfo f on f.PlanID =c.PlanID" & vbCrLf
        sql &= " JOIN ID_Class g on g.CLSID =f.CLSID" & vbCrLf
        sql &= " JOIN Org_OrgInfo e on e.ComIDNO=f.ComIDNO" & vbCrLf
        sql &= " JOIN Auth_Relship d on d.RID=f.RID" & vbCrLf
        sql &= " JOIN Key_TrainType h on h.TMID=f.TMID" & vbCrLf
        'sql &= " LEFT JOIN MVIEW_RELSHIP23 r3 on r3.RID3=f.RID" & vbCrLf
        'sql &= " LEFT JOIN Org_OrgInfo o2 on o2.OrgID=r3.OrgID2" & vbCrLf
        sql &= " WHERE f.NotOpen=@NotOpen" & vbCrLf

        parms.Add("NotOpen", NotOpen.SelectedValue)
        If sm.UserInfo.RID = "A" Then
            sql &= " AND c.TPlanID=@TPlanID" & vbCrLf
            sql &= " AND c.Years=@Years" & vbCrLf

            parms.Add("TPlanID", sm.UserInfo.TPlanID)
            parms.Add("Years", sm.UserInfo.Years)
        Else
            sql &= " AND c.PlanID=@PlanID" & vbCrLf
            parms.Add("PlanID", sm.UserInfo.PlanID)
        End If
        If Me.RIDValue.Value <> "" Then
            sql &= " and f.RID like @RID" & vbCrLf
            parms.Add("RID", Me.RIDValue.Value & "%")
        Else
            sql &= " and f.RID like @RID" & vbCrLf
            parms.Add("RID", sm.UserInfo.RID & "%")
        End If
        If Me.trainValue.Value <> "" Then
            sql &= " and f.TMID=@TMID" & vbCrLf
            parms.Add("TMID", Me.trainValue.Value)
        End If
        If Me.tb_classname.Text <> "" Then
            sql &= " and f.ClassCName like @ClassCName" & vbCrLf
            parms.Add("ClassCName", "%" & Me.tb_classname.Text & "%")
        End If
        If txtCJOB_NAME.Text <> "" Then   '通俗職類
            sql &= " and f.CJOB_UNKEY = @CJOB_UNKEY" & vbCrLf
            parms.Add("CJOB_UNKEY", cjobValue.Value)
        End If
        If Me.start_date.Text <> "" Then
            sql &= " and f.STDate >= @STDate1" & vbCrLf
            parms.Add("STDate1", Me.start_date.Text)
        End If
        If Me.end_date.Text <> "" Then
            sql &= " and f.STDate <= @STDate2" & vbCrLf
            parms.Add("STDate2", Me.end_date.Text)
        End If
        If RWClassFlag Then
            sql &= " and EXISTS (select 'x' from Auth_AccRWClass x where x.OCID =f.OCID AND x.Account=@Account)" & vbCrLf
            parms.Add("Account", sm.UserInfo.UserID)
        End If
        sql &= " )" & vbCrLf

        sql &= " ,WS1 AS (" & vbCrLf
        sql &= " SELECT b.OCID1" & vbCrLf
        '報名人數查詢
        sql &= " ,count(1) STUDETNUM" & vbCrLf
        '就服站協助報名
        'sql &= " ,count(case when b.EnterPath='W' THEN 1 end) Stud_NumW" & vbCrLf
        '甄試人數查詢
        '有任1分數(筆試、口試、總分)，大於等於0，即為甄試名單
        'sql &= " ,count(case when b.writeresult>=0 or b.oralresult>=0 or b.totalresult>=0 then 1 end) StudETNum2" & vbCrLf
        'sql &= " ,count(case when b.writeresult>=0 or b.oralresult>=0 then 1 end) StudETNum2" & vbCrLf
        '若總分，大於等於0，即為甄試名單'https://jira.turbotech.com.tw/browse/TIMSC-3
        sql &= " ,count(case when b.totalresult>=0 then 1 end) STUDETNUM2" & vbCrLf

        sql &= " FROM STUD_ENTERTYPE b" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP a on a.setid =b.setid" & vbCrLf
        sql &= " JOIN WC1 cc on cc.ocid =b.ocid1" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

        '1.若未填報名日期區間,搜尋方式同現在的處理方式.
        '2.若報名日期區間,起日為2009/3/25,但未填迄日,則搜尋日期的區間為2009/3/25至今天.
        '3.若報名日期區間,迄日為2009/4/2,但未填起日,則搜尋日期的區間為課程有報名資料至2009/3/25這天.
        '4.若報名日期區間,起日與迄日都為2009/4/2,則搜尋日期的區間為2009/4/2這天.
        '5.若報名日期區間,起日為2009/3/25,迄日為2009/4/2,則搜尋日期的區間為2009/3/25~2009/4/2.
        If vs_redate_start <> "" OrElse vs_redate_end <> "" Then
            If vs_redate_start <> "" AndAlso vs_redate_end = "" Then
                sql &= " and b.RelEnterDate >= @RelEnterDate" & vbCrLf
                sql &= " and b.RelEnterDate <= getdate()" & vbCrLf
                parms.Add("RelEnterDate", vs_redate_start)
            Else
                If vs_redate_start <> "" Then
                    sql &= " and b.RelEnterDate >= @RelEnterDate1" & vbCrLf
                    parms.Add("RelEnterDate1", vs_redate_start)
                End If
                If vs_redate_end <> "" Then
                    sql &= " and b.RelEnterDate <= @RelEnterDate2" & vbCrLf
                    parms.Add("RelEnterDate2", vs_redate_end)
                Else
                    sql &= " and b.RelEnterDate <= getdate()" & vbCrLf
                End If
            End If
        End If
        sql &= " GROUP BY b.OCID1" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WS2 AS (" & vbCrLf
        sql &= " SELECT b.OCID" & vbCrLf
        sql &= " ,COUNT(1) openTNum" & vbCrLf
        sql &= " FROM V_STUDENTINFO b" & vbCrLf
        sql &= " JOIN WC1 cc on cc.ocid =b.ocid" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " GROUP BY b.OCID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " select f.Years" & vbCrLf
        sql &= " ,f.CyclType" & vbCrLf
        sql &= " ,f.OClassID" & vbCrLf
        sql &= " ,f.OCID" & vbCrLf
        sql &= " ,f.PlanID" & vbCrLf
        sql &= " ,f.ComIDNO" & vbCrLf
        sql &= " ,f.SeqNO" & vbCrLf
        sql &= " ,f.ClassCName" & vbCrLf
        sql &= " ,f.TPropertyID" & vbCrLf
        sql &= " ,f.STDate" & vbCrLf
        sql &= " ,f.FTDate" & vbCrLf
        sql &= " ,CONVERT(varchar, f.STDate, 111)+'<br>'+CONVERT(varchar, f.FTDate, 111) S2FDATE" & vbCrLf
        sql &= " ,f.RID" & vbCrLf
        sql &= " ,f.TNum" & vbCrLf '訓練人數
        sql &= " ,f.ClassID" & vbCrLf
        sql &= " ,f.OrgName" & vbCrLf
        sql &= " ,f.TrainName" & vbCrLf
        'sql &= " ,f.OrgName2" & vbCrLf
        sql &= " ,dbo.NVL(s.StudETNum,0) StudETNum" & vbCrLf '報名人數
        sql &= " ,dbo.NVL(s.StudETNum2,0) StudETNum2" & vbCrLf '甄試人數
        'sql &= " ,dbo.NVL(s.Stud_NumW,0) Stud_NumW" & vbCrLf
        sql &= " ,dbo.NVL(s2.openTNum,0) openTNum" & vbCrLf '開訓人數
        sql &= " FROM WC1 f" & vbCrLf
        sql &= " JOIN WS1 s on s.ocid1=f.ocid" & vbCrLf
        sql &= " LEFT JOIN WS2 s2 on s2.ocid=f.ocid" & vbCrLf

        Dim dt As DataTable = Nothing
        '改用參數式查詢
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        ''一般輸出(匯出Excel)
        If Not flagExportExcel Then
            '一般輸出
            DG_Classinfo.Columns(cst_dg1_甄試人數).Visible = False '甄試人數
            DG_Classinfo.Columns(cst_dg1_開訓人數).Visible = False '開訓人數

            msg.Text = "查無資料!!"
            table4.Visible = False
            DG_Classinfo.Visible = False
            'DG_ClassInfo2.Visible = False '產投
            btnexport1.Visible = False
            PageControler1.Visible = False

            If dt.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無資料")
                Exit Sub
            End If

            msg.Text = ""
            table4.Visible = True

            DG_Classinfo.Visible = True
            btnexport1.Visible = True
            PageControler1.Visible = True

            'PageControler1.SqlString = sqlstr_class
            'PageControler1.ControlerLoad()
            ViewState("sort") = "OrgName"
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
            Return '  Exit Sub
        End If

        '匯出Excel
        If dt.Rows.Count = 0 Then
            msg.Text = "查無資料!!"
            Common.MessageBox(Me, "查無資料")
            Exit Sub
        End If

        DG_Classinfo.Columns(cst_dg1_甄試人數).Visible = True '甄試人數
        DG_Classinfo.Columns(cst_dg1_開訓人數).Visible = True '開訓人數

        iClassCnt = If(dt IsNot Nothing, dt.Rows.Count, 0) '班級數(總)
        iStudCnt1 = 0 '訓練人數(總)
        iStudCnt2 = 0 '報名人數(總)
        For Each dr As DataRow In dt.Rows
            iStudCnt1 += If(Convert.ToString(dr("TNum")) <> "", Val(dr("TNum")), 0) '訓練人數(總)
            iStudCnt2 += If(Convert.ToString(dr("StudETNum")) <> "", Val(dr("StudETNum")), 0) '報名人數(總)
        Next
        ViewState("sort") = "OrgName"
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢SQL '產投28
    Sub Search2()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DG_Classinfo2)

        Dim vs_redate_start As String = If(redate_start.Text <> "", Common.FormatDate(CDate(redate_start.Text)), "")
        Dim vs_redate_end As String = If(redate_end.Text <> "", String.Concat(Common.FormatDate(CDate(redate_end.Text)), " 23:59:59"), "")

        Dim parms As Hashtable = New Hashtable()
        Dim dt As DataTable
        '依sm.UserInfo.PlanID取得PlanKind
        Dim PlanKind As String = TIMS.Get_PlanKind(Me, objconn)

        Dim sSql As String = ""
        sSql &= " SELECT cc.OCID,cc.ISCLOSED,cc.ORGNAME" & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sSql &= " ,cc.OCLASSID,cc.CLASSID" & vbCrLf
        sSql &= " ,format(cc.STDate,'yyyy/MM/dd') STDATE" & vbCrLf
        sSql &= " ,format(cc.FTDate,'yyyy/MM/dd') FTDATE" & vbCrLf
        sSql &= " ,concat(format(cc.STDate,'yyyy/MM/dd'),'<br>',format(cc.FTDate,'yyyy/MM/dd')) S2FDATE" & vbCrLf
        sSql &= " ,cc.YEARS,cc.CYCLTYPE,cc.TNUM" & vbCrLf
        sSql &= " ,ISNULL(cc.TRAINNAME,cc.JOBNAME) TRAINNAME" & vbCrLf
        'sSql &= " --課程分類" & vbCrLf
        sSql &= " ,cc.GCID3,g3.PNAME" & vbCrLf
        'sSql &= " --網路報名人數-報名者自行取消人數) 扣除 網路審核失敗人數" & vbCrLf
        sSql &= " ,dbo.FN_GET_ENTERTYPE2(cc.OCID,1) TYPECNT1" & vbCrLf
        'sSql &= " --網路報名人數-報名者自行取消人數" & vbCrLf
        sSql &= " ,dbo.FN_GET_ENTERTYPE2(cc.OCID,2) TYPECNT2" & vbCrLf
        'sSql &= " --內網報名人數" & vbCrLf
        sSql &= " ,dbo.FN_GET_ENTERIN28(cc.OCID) TYPE2_IN28" & vbCrLf
        'sSql &= " --外網報名人數" & vbCrLf
        sSql &= " ,dbo.FN_GET_ENTEROUT28(cc.OCID) TYPE2_OUT28" & vbCrLf
        sSql &= " ,format(cc.SEnterDate,'yyyy/MM/dd') SENTERDATE" & vbCrLf
        sSql &= " ,format(cc.FEnterDate,'yyyy/MM/dd') FENTERDATE" & vbCrLf
        sSql &= " ,cc.ADMISSIONS ,dbo.FN_GET_ADMISSIONS(cc.OCID) ADMISSIONS_N" & vbCrLf
        sSql &= " ,convert(varchar, getdate(), 120) TODAY" & vbCrLf
        sSql &= " FROM VIEW2 cc" & vbCrLf
        sSql &= " LEFT JOIN V_GOVCLASSCAST3 g3 on g3.GCID3=cc.GCID3" & vbCrLf
        'sSql &= " WHERE cc.TPLANID='28' AND cc.YEARS='2023'" & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sSql &= " and cc.TPlanID =@TPlanID" & vbCrLf
            sSql &= " and cc.Years =@Years" & vbCrLf
            parms.Add("TPlanID", sm.UserInfo.TPlanID)
            parms.Add("Years", sm.UserInfo.Years)
        Else
            sSql &= " and cc.PlanID=@PlanID" & vbCrLf
            parms.Add("PlanID", sm.UserInfo.PlanID)
        End If
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value <> "" Then
            sSql &= " and cc.RID LIKE @RID" & vbCrLf
            parms.Add("RID", RIDValue.Value & "%")
        Else
            sSql &= " and cc.RID LIKE @RID" & vbCrLf
            parms.Add("RID", sm.UserInfo.RID & "%")
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            ' LabTMID.Text = "訓練業別"
            If jobValue.Value <> "" Then
                sSql &= " and ( cc.TMID = @TMID" & vbCrLf
                sSql &= " OR cc.TMID IN (" & vbCrLf
                sSql &= " select TMID from Key_TrainType where parent IN (" & vbCrLf '職類別
                sSql &= " select TMID from Key_TrainType where parent IN (" & vbCrLf '業別
                sSql &= " select TMID from Key_TrainType where busid ='G')" & vbCrLf '產業人才投資方案類
                sSql &= " AND TMID = @TMID" & vbCrLf
                sSql &= " )))" & vbCrLf
                parms.Add("TMID", jobValue.Value)
            End If
        Else
            If trainValue.Value <> "" Then
                sSql &= " and cc.TMID = @TMID" & vbCrLf
                parms.Add("TMID", trainValue.Value)
            End If
        End If

        If tb_classname.Text <> "" Then
            sSql &= " and cc.ClassCName like @ClassCName" & vbCrLf
            parms.Add("ClassCName", "%" & tb_classname.Text & "%")
        End If
        If txtCJOB_NAME.Text <> "" Then   '通俗職類
            sSql &= " and cc.CJOB_UNKEY = @CJOB_UNKEY" & vbCrLf
            parms.Add("CJOB_UNKEY", cjobValue.Value)
        End If
        If start_date.Text <> "" Then
            sSql &= " and cc.STDate >= @STDate1" & vbCrLf
            parms.Add("STDate1", start_date.Text)
        End If
        If end_date.Text <> "" Then
            sSql &= " and cc.STDate <= @STDate2" & vbCrLf
            parms.Add("STDate2", end_date.Text)
        End If
        If vs_redate_start <> "" Then
            sSql &= " and cc.SEnterDate >= @SEnterDate1" & vbCrLf
            parms.Add("SEnterDate1", vs_redate_start)
        End If
        If vs_redate_end <> "" Then
            sSql &= " and cc.SEnterDate <= @SEnterDate2" & vbCrLf
            parms.Add("SEnterDate2", vs_redate_end)
        End If
        If NotOpen.SelectedValue <> "" Then
            sSql &= " and cc.NotOpen=@NotOpen" & vbCrLf
            parms.Add("NotOpen", NotOpen.SelectedValue)
        End If

        Dim RWClassFlag As Boolean = False
        Select Case sm.UserInfo.RID
            Case "A"
            Case Else
                If sm.UserInfo.RoleID > 1 Then '非署(局)才有此限制
                    If PlanKind = "1" Then '自辦者只能列出賦予給此帳號的班級
                        RWClassFlag = True
                    End If
                End If
        End Select

        If RWClassFlag Then
            sSql &= " and EXISTS (select 'x' from Auth_AccRWClass x where x.OCID =cc.OCID AND x.Account=@Account)" & vbCrLf
            parms.Add("@Account", sm.UserInfo.UserID)
        End If

        Try
            '改用參數式查詢
            dt = DbAccess.GetDataTable(sSql, objconn, parms)
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Common.RespWrite(Me, sqlstr)
            'Throw ex
        End Try

        If Not flagExportExcel Then
            '一般輸出
            msg.Text = "查無資料!!"
            table4.Visible = False
            'DG_ClassInfo.Visible = False
            DG_Classinfo2.Visible = False '產投
            btnexport1.Visible = False
            PageControler1.Visible = False

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                msg.Text = "查無資料!!"
                Common.MessageBox(Me, "查無資料")
                Return
            End If

            msg.Text = ""
            table4.Visible = True

            DG_Classinfo2.Visible = True
            btnexport1.Visible = True
            PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        Else
            '匯出Excel
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                msg.Text = "查無資料!!"
                Common.MessageBox(Me, "查無資料")
                Return
            End If
            CPdt = dt.Copy()
        End If

    End Sub

    '匯出 TIMS (EXCEL)
    Sub Export1()
        flagExportExcel = True '匯出Excel

        DG_Classinfo.AllowPaging = False
        'DG_ClassInfo.Columns(8).Visible = False
        DG_Classinfo.EnableViewState = False  '把ViewState給關了

        Call Search1()

        Dim sFileName1 As String = "學員報名人數統計表"

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        '加title條件區
        Dim tdstyle As String = "" '依dataGrid資料欄
        Dim myTable As String = ""

        tdstyle = " colspan='" & Cst_dg1_colspan & "'" '依dataGrid資料欄

        Dim strHTML As String = ""
        myTable = ""
        myTable &= "<table border='0' cellspacing='0' cellpadding='0' align='center' style='width:100%;border-collapse@collapse;'>"
        myTable &= "<tr><td " & tdstyle & ">" & "查詢條件如下：" & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練機構：" & center.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練職類：" & TB_career_id.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "班級名稱：" & tb_classname.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "開訓日期：" & start_date.Text & "~" & end_date.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "報名日期：" & redate_start.Text & "~" & redate_end.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "開班狀態：" & NotOpen.SelectedItem.Text & "</td></tr>"
        myTable &= "<table>"
        strHTML &= (myTable)
        '加title條件區

        DG_Classinfo.AllowPaging = False
        DG_Classinfo.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        div1.RenderControl(objHtmlTextWriter)
        strHTML &= (Convert.ToString(objStringWriter))

        '加bottom 底部條件區
        myTable = ""
        myTable &= "<table border='0' cellspacing='0' cellpadding='0' align='center' style='width:100%;border-collapse@collapse;'>"
        myTable &= "<tr><td " & tdstyle & ">" & "統計數量：" & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "班級：" & iClassCnt & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練人數：" & iStudCnt1 & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "報名人數：" & iStudCnt2 & "筆</td></tr>"
        myTable &= "<table>"
        strHTML &= (myTable)

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        '加title條件區
        DG_Classinfo.AllowPaging = True
        'DG_ClassInfo.Columns(8).Visible = True
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub

    '匯出 產投28
    Sub Export2()
        flagExportExcel = True '匯出Excel

        Call Search2()

        Dim sFileName1 As String = "學員報名人數統計表"

        Dim str_COLUMN1 As String = "序號,訓練機構,課程代碼,開結訓日,結訓日期,班別名稱,課程分類,訓練人數,報名人數1,報名人數2,內網報名人數,外網報名人數,招生狀態"
        Dim str_COLUMN2 As String = "SEQNO,ORGNAME,OCID,STDATE,FTDATE,CLASSCNAME,PNAME,TNUM,TYPECNT1,TYPECNT2,TYPE2_IN28,TYPE2_OUT28,ADMISSIONS_N"

        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        '加title條件區
        Dim tdstyle As String = " colspan='" & str_COLUMN2.Split(",").Length & "'" '依資料欄
        Dim myTable As String = ""

        myTable = ""
        myTable &= "<table border='0' cellspacing='0' cellpadding='0' align='center' style='width:100%;border-collapse@collapse;'>"
        myTable &= "<tr><td " & tdstyle & ">" & "查詢條件如下：" & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練機構：" & center.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練職類：" & TB_career_id.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "班級名稱：" & tb_classname.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "開訓日期：" & start_date.Text & "~" & end_date.Text & "</td></tr>"

        myTable &= "<tr><td " & tdstyle & ">" & "報名日期：" & redate_start.Text & "~" & redate_end.Text & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "開班狀態：" & NotOpen.SelectedItem.Text & "</td></tr>"
        myTable &= "</table>"
        '加title條件區

        'myTable = ""
        myTable &= "<table border='1' cellspacing='0' cellpadding='0' rules='all' style='width:100%;border-collapse@collapse;'>"
        'Common.RespWrite(Me, myTable)

        'HEAD
        'myTable = ""
        myTable &= "<tr>"
        For Each COL1 As String In str_COLUMN1.Split(",")
            myTable &= String.Concat("<td>", COL1, "</td>")
        Next
        myTable &= "</tr>"
        'Common.RespWrite(Me, myTable)

        iClassCnt = If(CPdt IsNot Nothing, CPdt.Rows.Count, 0) '班級數(總)
        iStudCnt1 = 0 '訓練人數(總)
        'iStudCnt2 = 0 '報名人數(總)
        Dim iStudCnt1_1 As Integer = 0
        Dim iStudCnt1_2 As Integer = 0
        Dim iStudCnt2_IN As Integer = 0
        Dim iStudCnt2_OUT As Integer = 0
        If CPdt IsNot Nothing Then
            For i As Integer = 0 To CPdt.Rows.Count - 1
                Dim dr As DataRow = CPdt.Rows(i)
                'myTable = ""
                myTable &= If(i Mod 2 = 1, "<tr style='background-color:#EEEEEE;'>", "<tr>")
                For Each COL2 As String In str_COLUMN2.Split(",")
                    Select Case COL2
                        Case "SEQNO"
                            myTable &= String.Concat("<td align=""center"">", CStr(i + 1), "</td>")
                        Case "OCID", "TNUM", "TYPECNT1", "TYPECNT2", "TYPE2_IN28", "TYPE2_OUT28"
                            myTable &= String.Concat("<td align=""center"">", dr(COL2), "</td>")
                        Case Else
                            myTable &= String.Concat("<td>", dr(COL2), "</td>")
                    End Select
                    'myTable &= String.Concat("<td>", COL1, "</td>")
                Next

                'myTable &= "<td align=""center"">" & CStr(i + 1) & "</td>"
                'myTable &= "<td>" & dr("OrgName") & "</td>"
                ''myTable &= "<td>" & dr("ClassID") & "</td>"
                'myTable &= "<td align=""center"">" & dr("OCID") & "</td>"
                'myTable &= "<td>" & dr("S2FDATE") & "</td>"
                'myTable &= "<td>" & dr("CLASSCNAME") & "</td>"
                'myTable &= "<td>" & dr("TrainName") & "</td>"
                'myTable &= "<td align=""center"">" & dr("TNum") & "</td>" ' "訓練人數" 
                'myTable &= "<td align=""center"">" & dr("typeCnt1") & "</td>" '報名人數1
                'myTable &= "<td align=""center"">" & dr("typeCnt2") & "</td>" '報名人數2
                myTable &= "</tr>"
                'Common.RespWrite(Me, myTable)
                iStudCnt1 += If(Convert.ToString(dr("TNUM")) <> "", Val(dr("TNUM")), 0)
                iStudCnt1_1 += If(Convert.ToString(dr("TYPECNT1")) <> "", Val(dr("TYPECNT1")), 0)
                iStudCnt1_2 += If(Convert.ToString(dr("TYPECNT2")) <> "", Val(dr("TYPECNT2")), 0)
                iStudCnt2_IN += If(Convert.ToString(dr("TYPE2_IN28")) <> "", Val(dr("TYPE2_IN28")), 0)
                iStudCnt2_OUT += If(Convert.ToString(dr("TYPE2_OUT28")) <> "", Val(dr("TYPE2_OUT28")), 0)
            Next
        End If
        'myTable = ""
        myTable &= "</table>"

        '加title條件區
        'myTable = ""
        myTable &= "<table border='0' cellspacing='0' cellpadding='0' align='center' style='width:100%;border-collapse@collapse;'>"
        myTable &= "<tr><td " & tdstyle & ">" & "統計數量：" & "</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "班級：" & iClassCnt & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "訓練人數：" & iStudCnt1 & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "報名人數1：" & iStudCnt1_1 & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "報名人數2：" & iStudCnt1_2 & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "內網報名人數：" & iStudCnt2_IN & "筆</td></tr>"
        myTable &= "<tr><td " & tdstyle & ">" & "外網報名人數：" & iStudCnt2_OUT & "筆</td></tr>"
        myTable &= "</table>"
        'Common.RespWrite(Me, myTable)
        '加title條件區
        Dim strHTML As String = ""
        strHTML &= (myTable)

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        'DG_Classinfo2.AllowPaging = True
        'DG_Classinfo2.Columns(Cst_功能欄位).Visible = True
        'Call TIMS.CloseDbConn(objconn)
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Me.start_date.Text = TIMS.ClearSQM(Me.start_date.Text)
        Me.end_date.Text = TIMS.ClearSQM(Me.end_date.Text)
        Me.redate_start.Text = TIMS.ClearSQM(Me.redate_start.Text)
        Me.redate_end.Text = TIMS.ClearSQM(Me.redate_end.Text)
        If start_date.Text <> "" AndAlso Not TIMS.IsDate1(start_date.Text) Then
            Errmsg &= "開訓日期的起始日 不是正確的日期格式(yyyy/MM/dd)" & vbCrLf
        End If
        If end_date.Text <> "" AndAlso Not TIMS.IsDate1(end_date.Text) Then
            Errmsg &= "開訓日期的結束日 不是正確的日期格式(yyyy/MM/dd)" & vbCrLf
        End If
        If redate_start.Text <> "" AndAlso Not TIMS.IsDate1(redate_start.Text) Then
            Errmsg &= "報名日期的起始日 不是正確的日期格式(yyyy/MM/dd)" & vbCrLf
        End If
        If redate_end.Text <> "" AndAlso Not TIMS.IsDate1(redate_end.Text) Then
            Errmsg &= "報名日期的結束日 不是正確的日期格式(yyyy/MM/dd)" & vbCrLf
        End If
        If Errmsg = "" Then
            Me.start_date.Text = TIMS.Cdate3(Me.start_date.Text)
            Me.end_date.Text = TIMS.Cdate3(Me.end_date.Text)
            Me.redate_start.Text = TIMS.Cdate3(Me.redate_start.Text)
            Me.redate_end.Text = TIMS.Cdate3(Me.redate_end.Text)
        End If
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢鈕
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        flagExportExcel = False ' 非 匯出Excel

        Dim flag28 As Boolean = (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1)

        btnSave1.Visible = False
        btnBack1.Visible = False
        If flag28 Then
            btnSave1.Visible = True
            btnBack1.Visible = True
            '產投28
            '報名人數1＝網路報名人數-報名者自行取消人數     +現場報名人數-網路審核失敗人數                                                
            '報名人數2＝網路報名人數-報名者自行取消人數     +現場報名人數
            Call Search2()
        Else
            'DG_Classinfo.Columns(9).Visible = False '甄試人數
            'DG_Classinfo.Columns(10).Visible = False '開訓人數
            '非產投查詢
            Call Search1()
        End If
    End Sub

    '匯出EXCEL
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnexport1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Call ChkValue1()

        Dim flag28 As Boolean = (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1)

        '產投28
        If flag28 Then
            Call Export2() : Exit Sub
        Else
            'TIMS
            Call Export1() : Exit Sub
        End If
    End Sub

    '取消。
    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        '重新查詢。
        Call Search2()
    End Sub

    Sub SaveData1()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE CLASS_CLASSINFO" & vbCrLf
        sql &= " SET ADMISSIONS=@ADMISSIONS" & vbCrLf
        sql &= " WHERE OCID=@OCID" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim uCmd As New SqlCommand(sql, objconn)

        For Each eItem As DataGridItem In DG_Classinfo2.Items
            Dim Label99 As Label = eItem.FindControl("Label99")
            Dim ddlADMISSIONS As DropDownList = eItem.FindControl("ddlADMISSIONS")
            Dim hid_OCID As HiddenField = eItem.FindControl("hid_OCID")
            If ddlADMISSIONS.SelectedValue <> "" AndAlso hid_OCID.Value <> "" Then
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("ADMISSIONS", SqlDbType.Int).Value = Val(ddlADMISSIONS.SelectedValue)
                    .Parameters.Add("OCID", SqlDbType.Int).Value = Val(hid_OCID.Value)
                    '.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
                End With
            End If
        Next

    End Sub

    '儲存
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click

        Call SaveData1()

        Common.MessageBox(Me, "儲存成功")

        '重新查詢。
        Call Search2()
    End Sub

End Class
