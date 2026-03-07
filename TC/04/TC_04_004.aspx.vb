Partial Class TC_04_004
    Inherits AuthBasePage

    '產業人才投資方案，審核計畫專用

    'Dim Auth_Relship As DataTable

    'select * from view_function where name ='各類課程明細表'
    '477
    '綜合查詢統計表	SD/15/SD_15_012.aspx	
    '各類課程明細表 SD/15/SD_15_009.aspx	
    'v_Depot04
    'view_Depot06
    'UPDATE Plan_Depot
    'SELECT * FROM KEY_BUSINESS WHERE DepID = '12' 
    'SELECT KID,KNAME FROM V_DEPOT12 ORDER BY KID
    Dim strYears As String = "" '2014 / 2015'(經費分類代碼。)

    'Const Cst_EmptySelValue As String = "==無==" 'TIMS.cst_ddl_PleaseChoose3
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
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        objconn = DbAccess.GetConnection()
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End

        'Select Case strYears
        '    Case "2014"
        '    Case "2015"
        'End Select
        '(經費分類代碼。)

        strYears = "2014" '2014年  顯示層級。
        If sm.UserInfo.Years >= "2015" Then strYears = "2015" '2015年 不顯示層級。
        Dim flag2017 As Boolean = (sm.UserInfo.Years >= "2017") '2017轉程式執行 'If sm.UserInfo.Years >= "2017" Then flag2017 = True
        If Not flag2017 AndAlso TIMS.sUtl_ChkTest() Then flag2017 = True
        If flag2017 Then
            '2017轉程式執行
            Dim url1 As String = "TC_04_004_17.aspx?ID=" & Request("ID")
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Me.LabTMID.Text = "訓練業別"
        End If
        PageControler1.PageDataGrid = Me.DataGrid1


        If Not Me.IsPostBack Then
            '有不區分
            OrgKind2 = TIMS.Get_RblSearchPlan(Me, OrgKind2)
            Common.SetListItem(OrgKind2, "A")

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            Call Create1()

            DataGridTable1.Visible = False '預設搜尋資料不顯示

            panelSearch.Visible = True '搜尋功能啟動
            PanelEdit1.Visible = False '修改功能關閉
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        'Button1.Attributes("onclick") = "return SavaData();"
    End Sub

    '初始載入
    Sub Create1()
        '課程分類
        Dim sql As String = ""
        ddlDepot12.Items.Clear()
        sql = "SELECT KID,KNAME FROM V_DEPOT12 ORDER BY KID" '19:其他類
        DbAccess.MakeListItem(ddlDepot12, sql, objconn)
        ddlDepot12.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        ddlKID04.Items.Clear()
        ddlKID06.Items.Clear()
        ddlKID10.Items.Clear()
        Select Case strYears
            Case "2014"
                '2014 四大新興智慧型產業
                sql = "SELECT KID,KNAME FROM V_DEPOT04 ORDER BY KID"
                DbAccess.MakeListItem(ddlKID04, sql, objconn)
                ddlKID04.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

                '2014 六大新興產業
                sql = "SELECT KID,KNAME FROM V_DEPOT06 ORDER BY KID"
                DbAccess.MakeListItem(ddlKID06, sql, objconn)
                ddlKID06.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

                '2014 十大重點服務業
                sql = "SELECT KID,KNAME FROM V_DEPOT10 ORDER BY KID"
                DbAccess.MakeListItem(ddlKID10, sql, objconn)
                ddlKID10.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))
            Case "2015" '(9.10.11)
                '2015 四大新興智慧型產業
                sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='09' ORDER BY KID"
                DbAccess.MakeListItem(ddlKID04, sql, objconn)
                ddlKID04.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

                '2015 六大新興產業
                sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='10' ORDER BY KID"
                DbAccess.MakeListItem(ddlKID06, sql, objconn)
                ddlKID06.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

                '2015 十大重點服務業
                sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='11' ORDER BY KID"
                DbAccess.MakeListItem(ddlKID10, sql, objconn)
                ddlKID10.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))
        End Select

        '轄區重點產業【排除停用】
        ddlDEPOT13.Items.Clear()
        Select Case Convert.ToString(sm.UserInfo.DistID)
            Case "000" '署(局)
                sql = "SELECT SEQNO,KNAME ,KID FROM KEY_BUSINESS WHERE DEPID='13' AND Status IS NULL ORDER BY KID"
            Case Else
                sql = "SELECT SEQNO,KNAME ,KID FROM KEY_BUSINESS WHERE DEPID='13' AND DISTID ='" & sm.UserInfo.DistID & "' AND Status IS NULL ORDER BY KID"
        End Select
        DbAccess.MakeListItem(ddlDEPOT13, sql, objconn)
        ddlDEPOT13.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        '生產力4.0
        ddlKID14.Items.Clear()
        sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='14' ORDER BY KID"
        DbAccess.MakeListItem(ddlKID14, sql, objconn)
        ddlKID14.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))
    End Sub

    'SQL PageDataTable LIST
    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select ip.Years" & vbCrLf ' top 100
        sql &= " ,ip.DistName" & vbCrLf
        sql &= " ,oo.OrgName" & vbCrLf
        sql &= " ,pp.planid,pp.comidno,pp.SeqNo " & vbCrLf
        sql &= " ,vt.JobName" & vbCrLf
        Select Case strYears
            Case "2014"
                sql &= " ,vg.GovClassN" & vbCrLf
            Case "2015"
                sql &= " ,vg2.GCODE2 GovClassN" & vbCrLf
        End Select
        sql &= " ,dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE) CLASSNAME" & vbCrLf

        sql &= " ,pp.GCID" & vbCrLf
        sql &= " ,pp.GCID2" & vbCrLf
        sql &= " ,pp.GCID3" & vbCrLf

        sql &= " ,dd.AppResult" & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID04 ,d4.KID) KID04 " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID06 ,d6.KID) KID06 " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID10 ,d10.KID) KID10 " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',K4.KNAME ,D4.KNAME) D4KNAME " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',K6.KNAME ,D6.KNAME) D6KNAME " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',K10.KNAME ,D10.KNAME) D10KNAME " & vbCrLf

        '課程分類 (view_depot12 vd12)
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID12 ,d12.KID) KID12 " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',K12.KNAME  ,d12.KNAME) D12KNAME " & vbCrLf
        '轄區重點產業
        sql &= " ,dd.SEQNOd13 ,d13.KNAME D13KNAME" & vbCrLf
        '生產力4.0
        sql &= " ,dd.KID14 ,K14.KNAME D14KNAME" & vbCrLf

        sql &= " FROM dbo.PLAN_PLANINFO pp " & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip on ip.planid =pp.planid " & vbCrLf
        sql &= " JOIN dbo.Auth_Relship ar ON ar.RID=pp.RID " & vbCrLf '--/* 業務確認 */
        sql &= " JOIN dbo.Org_OrgInfo oo ON oo.OrgID=ar.OrgID  " & vbCrLf '--/* 機構確認 */
        sql &= " JOIN dbo.Org_OrgPlanInfo oop ON oop.RSID=ar.RSID " & vbCrLf '--/* 機構資料確認 */
        sql &= " LEFT JOIN dbo.VIEW_TRAINTYPE vt on vt.TMID=pp.TMID" & vbCrLf

        sql &= " LEFT JOIN dbo.Plan_VerReport pvr ON pp.PlanID = pvr.PlanID AND  pp.ComIDNO = pvr.ComIDNO AND pp.SeqNo = pvr.SeqNo " & vbCrLf
        sql &= " LEFT JOIN dbo.Plan_VerRecord pvrd ON pp.PlanID = pvrd.PlanID AND  pp.ComIDNO = pvrd.ComIDNO AND pp.SeqNo = pvrd.SeqNo " & vbCrLf
        sql &= " LEFT JOIN dbo.Plan_Depot dd ON pp.PlanID = dd.PlanID AND  pp.ComIDNO = dd.ComIDNO AND pp.SeqNo = dd.SeqNo " & vbCrLf
        Select Case strYears
            Case "2014"
                sql &= " left join view_GovClassCast vg on vg.GCID =pp.GCID" & vbCrLf

                sql &= " left join VIEW_DEPOT06 d4  ON d4.GCID =pp.GCID" & vbCrLf
                sql &= " left join VIEW_DEPOT07 d6  ON d6.GCID =pp.GCID" & vbCrLf
                sql &= " left join VIEW_DEPOT08 d10 ON d10.GCID =pp.GCID" & vbCrLf

                sql &= " left join V_DEPOT04 K4  ON K4.KID =dd.KID04 " & vbCrLf
                sql &= " left join V_DEPOT06 K6  ON K6.KID =dd.KID06 " & vbCrLf
                sql &= " left join V_DEPOT10 K10 ON K10.KID =dd.KID10 " & vbCrLf
            Case "2015" '(9.10.11)
                sql &= " left join v_GovClassCast2 vg2 on vg2.GCID2 =pp.GCID2" & vbCrLf

                sql &= " left join VIEW_DEPOT09 d4  ON d4.GCID2 =pp.GCID2" & vbCrLf
                sql &= " left join VIEW_DEPOT10 d6  ON d6.GCID2 =pp.GCID2" & vbCrLf
                sql &= " left join VIEW_DEPOT11 d10 ON d10.GCID2 =pp.GCID2" & vbCrLf

                sql &= " left join (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='09') K4  ON K4.KID =dd.KID04 " & vbCrLf
                sql &= " left join (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='10') K6  ON K6.KID =dd.KID06 " & vbCrLf
                sql &= " left join (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='11') K10 ON K10.KID =dd.KID10 " & vbCrLf
        End Select
        '課程分類 (view_depot12 vd12)
        sql &= " LEFT JOIN VIEW_DEPOT12 d12 ON d12.GCID2 =pp.GCID2" & vbCrLf
        sql &= " LEFT JOIN (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='12') K12 ON K12.KID =dd.KID12 " & vbCrLf
        '轄區重點產業
        sql &= " LEFT JOIN (SELECT SEQNO SEQNOd13,KNAME FROM KEY_BUSINESS WHERE DEPID='13') d13 ON d13.SEQNOd13 =dd.SEQNOd13 " & vbCrLf
        '生產力4.0
        sql &= " LEFT JOIN (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='14') K14 ON K14.KID =dd.KID14 " & vbCrLf

        sql &= " where 1=1" & vbCrLf
        sql &= " AND pvr.IsApprPaper='Y'" & vbCrLf '--正式
        sql &= " and ip.PlanKind=2 " & vbCrLf '計畫種類:1.自辦／2.委外

        'Sql += " AND pvr.SecResult='Y' --複審結果通過" & vbCrLf
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0" '署(局)
            Case Else '非署(局)
                sql &= " and ip.PlanID ='" & sm.UserInfo.PlanID & "' " & vbCrLf '--登入委訓計畫
                sql &= " and ip.DistID ='" & sm.UserInfo.DistID & "' " & vbCrLf '--登入轄區
        End Select
        sql &= " and ip.TPlanID ='" & sm.UserInfo.TPlanID & "' " & vbCrLf '--登入計畫
        sql &= " and ip.Years ='" & sm.UserInfo.Years & "' " & vbCrLf '--登入年度

        '【中心(分署)】業務權限
        sql &= " and EXISTS ( SELECT 'x' FROM Auth_Relship x WHERE x.relship like '" & sm.UserInfo.RelShip & "%' and x.RID =pp.RID ) " & vbCrLf '--業務權限

        If Me.ViewState("jobValue") <> "" Then
            sql &= " and pp.TMID = " & Me.ViewState("jobValue") & vbCrLf
        End If
        If Me.ViewState("trainValue") <> "" Then
            sql &= " and pp.TMID = " & Me.ViewState("trainValue") & vbCrLf
        End If
        '通俗職類
        If Me.ViewState("cjobValue") <> "" Then
            sql &= " and pp.CJOB_UNKEY = " & Me.ViewState("cjobValue") & "" & vbCrLf
        End If
        If Me.ViewState("ClassName") <> "" Then
            sql &= " and pp.ClassName like '%" & Me.ViewState("ClassName") & "%'" & vbCrLf
        End If
        If Me.ViewState("CyclType") <> "" Then
            sql &= " and pp.CyclType='" & Me.ViewState("CyclType") & "'" & vbCrLf
        End If
        If Me.ViewState("STDate1") <> "" Then
            sql &= " and pp.STDate >=" & TIMS.To_date(Me.ViewState("STDate1")) & vbCrLf
        End If
        If Me.ViewState("STDate2") <> "" Then
            sql &= " and pp.STDate <=" & TIMS.To_date(Me.ViewState("STDate2")) & vbCrLf
        End If
        If Me.ViewState("FDDate1") <> "" Then
            sql &= " and pp.FDDate >=" & TIMS.To_date(Me.ViewState("FDDate1")) & vbCrLf
        End If
        If Me.ViewState("FDDate2") <> "" Then
            sql &= " and pp.FDDate <=" & TIMS.To_date(Me.ViewState("FDDate2")) & vbCrLf
        End If
        If Me.ViewState("OrgKind2") <> "" Then
            sql &= " and oo.OrgKind2 ='" & Me.ViewState("OrgKind2") & "'" & vbCrLf
        End If
        If Me.ViewState("sqlSecResult") <> "" Then
            sql += Me.ViewState("sqlSecResult") & vbCrLf
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable1.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    'SQL return datarow
    Function search2(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String) As DataRow
        Dim rst As DataRow = Nothing

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select ip.Years" & vbCrLf
        sql &= " ,ip.DistName" & vbCrLf
        sql &= " ,oo.OrgName" & vbCrLf
        sql &= " ,pp.planid,pp.comidno,pp.SeqNo " & vbCrLf
        sql &= " ,vt.JobName" & vbCrLf
        sql &= " ,vt.TrainName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,CONVERT(varchar, pp.STDate, 111) STDate" & vbCrLf
        sql &= " ,CONVERT(varchar, pp.FDDate, 111) FDDate" & vbCrLf
        sql &= " ,pp.THours" & vbCrLf
        'Sql += " ,vt.JobName" & vbCrLf
        Select Case strYears
            Case "2014"
                sql &= " ,vg.GovClassN" & vbCrLf
            Case "2015"
                sql &= " ,vg2.GCODE2 GovClassN" & vbCrLf
        End Select
        sql &= " ,pp.GCID" & vbCrLf
        sql &= " ,pp.GCID2" & vbCrLf

        sql &= " ,dd.AppResult" & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID04 ,d4.KID) KID04 " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID06 ,d6.KID) KID06 " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID10 ,d10.KID) KID10 " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',K4.KNAME ,D4.KNAME) D4KNAME " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',K6.KNAME ,D6.KNAME) D6KNAME " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',K10.KNAME ,D10.KNAME) D10KNAME " & vbCrLf
        '課程分類
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID12 ,d12.KID) KID12 " & vbCrLf
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',K12.KNAME  ,d12.KNAME) D12KNAME " & vbCrLf
        '轄區重點產業
        sql &= " ,dd.SEQNOd13 ,d13.KNAME D13KNAME" & vbCrLf
        '生產力4.0
        sql &= " ,dd.KID14 ,K14.KNAME D14KNAME" & vbCrLf

        sql &= " FROM dbo.Plan_PlanInfo pp   " & vbCrLf
        sql &= " JOIN dbo.view_Plan ip on ip.planid =pp.planid " & vbCrLf
        sql &= " JOIN dbo.Auth_Relship ar ON ar.RID=pp.RID " & vbCrLf '--/* 業務確認 */
        sql &= " JOIN dbo.Org_OrgInfo oo ON oo.OrgID=ar.OrgID " & vbCrLf '--/* 機構確認 */
        sql &= " JOIN dbo.Org_OrgPlanInfo oop ON oop.RSID=ar.RSID " & vbCrLf '--/* 機構資料確認 */
        sql &= " LEFT JOIN dbo.VIEW_TRAINTYPE vt on vt.TMID=pp.TMID" & vbCrLf

        sql &= " LEFT JOIN dbo.Plan_VerReport pvr ON pp.PlanID = pvr.PlanID AND  pp.ComIDNO = pvr.ComIDNO AND pp.SeqNo = pvr.SeqNo " & vbCrLf
        sql &= " LEFT JOIN dbo.Plan_VerRecord pvrd ON pp.PlanID = pvrd.PlanID AND  pp.ComIDNO = pvrd.ComIDNO AND pp.SeqNo = pvrd.SeqNo " & vbCrLf
        sql &= " LEFT JOIN dbo.Plan_Depot dd ON pp.PlanID = dd.PlanID AND  pp.ComIDNO = dd.ComIDNO AND pp.SeqNo = dd.SeqNo " & vbCrLf

        Select Case strYears
            Case "2014"
                sql &= " left join view_GovClassCast vg on vg.GCID =pp.GCID" & vbCrLf

                sql &= " left join VIEW_DEPOT06 d4  ON d4.GCID =pp.GCID" & vbCrLf
                sql &= " left join VIEW_DEPOT07 d6  ON d6.GCID =pp.GCID" & vbCrLf
                sql &= " left join VIEW_DEPOT08 d10 ON d10.GCID =pp.GCID" & vbCrLf

                sql &= " left join V_DEPOT04 K4  ON K4.KID =dd.KID04 " & vbCrLf
                sql &= " left join V_DEPOT06 K6  ON K6.KID =dd.KID06 " & vbCrLf
                sql &= " left join V_DEPOT10 K10 ON K10.KID =dd.KID10 " & vbCrLf
            Case "2015"
                sql &= " left join v_GovClassCast2 vg2 on vg2.GCID2 =pp.GCID2" & vbCrLf

                sql &= " left join VIEW_DEPOT09 d4  ON d4.GCID2 =pp.GCID2" & vbCrLf
                sql &= " left join VIEW_DEPOT10 d6  ON d6.GCID2 =pp.GCID2" & vbCrLf
                sql &= " left join VIEW_DEPOT11 d10 ON d10.GCID2 =pp.GCID2" & vbCrLf

                sql &= " left join (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='09') K4  ON K4.KID =dd.KID04 " & vbCrLf
                sql &= " left join (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='10') K6  ON K6.KID =dd.KID06 " & vbCrLf
                sql &= " left join (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='11') K10 ON K10.KID =dd.KID10 " & vbCrLf
        End Select
        sql &= " LEFT JOIN VIEW_DEPOT12 d12 ON d12.GCID2 =pp.GCID2" & vbCrLf
        sql &= " LEFT JOIN (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='12') K12 ON K12.KID =dd.KID12 " & vbCrLf
        '轄區重點產業
        sql &= " LEFT JOIN (SELECT SEQNO SEQNOd13,KNAME FROM KEY_BUSINESS WHERE DEPID='13') d13 ON d13.SEQNOd13 =dd.SEQNOd13 " & vbCrLf
        '生產力4.0
        sql &= " LEFT JOIN (SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='14') K14 ON K14.KID =dd.KID14 " & vbCrLf

        sql &= " where 1=1" & vbCrLf
        sql &= " AND pvr.IsApprPaper='Y'" & vbCrLf '--正式
        sql &= " and ip.PlanKind=2 " & vbCrLf '計畫種類:1.自辦／2.委外

        sql &= " and pp.PlanID ='" & PlanID & "' " & vbCrLf
        sql &= " and pp.ComIDNO ='" & ComIDNO & "' " & vbCrLf
        sql &= " and pp.SeqNo ='" & SeqNo & "' " & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then rst = dt.Rows(0)

        Return rst
    End Function

    '設定 ViewState Value
    Sub sUtl_ViewStateValue()
        '訓練業別
        Me.ViewState("jobValue") = IIf(Me.jobValue.Value <> "", Me.jobValue.Value, "")
        '訓練職類
        Me.ViewState("trainValue") = IIf(Me.trainValue.Value <> "", Me.trainValue.Value, "")
        '通俗職類
        Me.ViewState("cjobValue") = IIf(Me.cjobValue.Value <> "", Me.cjobValue.Value, "")
        '班別名稱
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        Me.ViewState("ClassName") = IIf(Trim(ClassName.Text) <> "", Trim(ClassName.Text).Replace("'", "''"), "")

        Me.ViewState("jobValue") = TIMS.ClearSQM(Me.ViewState("jobValue"))
        Me.ViewState("trainValue") = TIMS.ClearSQM(Me.ViewState("trainValue"))
        Me.ViewState("cjobValue") = TIMS.ClearSQM(Me.ViewState("cjobValue"))
        Me.ViewState("ClassName") = TIMS.ClearSQM(Me.ViewState("ClassName"))

        Me.ViewState("CyclType") = TIMS.FmtCyclType(CyclType.Text) 'TIMS.ClearSQM(Me.ViewState("CyclType"))

        Me.ViewState("STDate1") = IIf(Me.STDate1.Text <> "", Me.STDate1.Text, "")
        Me.ViewState("STDate2") = IIf(Me.STDate2.Text <> "", Me.STDate2.Text, "")
        Me.ViewState("FDDate1") = IIf(Me.FDDate1.Text <> "", Me.FDDate1.Text, "")
        Me.ViewState("FDDate2") = IIf(Me.FDDate2.Text <> "", Me.FDDate2.Text, "")

        Me.ViewState("STDate1") = TIMS.ClearSQM(Me.ViewState("STDate1"))
        Me.ViewState("STDate2") = TIMS.ClearSQM(Me.ViewState("STDate2"))
        Me.ViewState("FDDate1") = TIMS.ClearSQM(Me.ViewState("FDDate1"))
        Me.ViewState("FDDate2") = TIMS.ClearSQM(Me.ViewState("FDDate2"))

        Me.ViewState("OrgKind2") = ""
        If OrgKind2.SelectedValue <> "" Then
            Select Case OrgKind2.SelectedValue
                Case "G", "W"
                    Me.ViewState("OrgKind2") = OrgKind2.SelectedValue
            End Select
        End If
        Me.ViewState("OrgKind2") = TIMS.ClearSQM(Me.ViewState("OrgKind2"))

        '業務權限(中心)
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        If Me.ViewState("jobValue") <> "" Then Me.ViewState("trainValue") = ""

        Me.ViewState("sqlSecResult") = ""
        '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
        Select Case PlanMode.SelectedValue
            Case "S" '審核中的
                Me.ViewState("sqlSecResult") = " AND pvr.SecResult IS NULL"
            Case "Y" '已通過
                Me.ViewState("sqlSecResult") = " AND pvr.SecResult='Y'"
            Case "R" '退件修正
                Me.ViewState("sqlSecResult") = " AND pvr.SecResult in ('R','N')"
        End Select

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        Call sUtl_ViewStateValue()  '設定 ViewState Value
        Call Search1()
    End Sub

    'update plan_planinfo class_classinfo 
    Sub sUtl_UpdatePlaninfo(ByVal sSearchW As String)
        Dim PlanID As String = TIMS.GetMyValue(sSearchW, "PlanID")
        Dim ComIDNO As String = TIMS.GetMyValue(sSearchW, "ComIDNO")
        Dim SeqNo As String = TIMS.GetMyValue(sSearchW, "SeqNo")
        If PlanID = "" Then Exit Sub
        If ComIDNO = "" Then Exit Sub
        If SeqNo = "" Then Exit Sub
        'select * from Key_TrainType where TMID=554
        'select * from ID_GOVCLASSCAST2 where GCID2=1160
        Dim sql As String = ""
        sql = ""
        sql &= " UPDATE CLASS_CLASSINFO "
        sql &= " SET TMID=554"
        sql &= " WHERE 1=1"
        sql &= " AND PlanID=@PlanID"
        sql &= " AND ComIDNO=@ComIDNO"
        sql &= " AND SeqNo=@SeqNo"
        Dim uCmd2 As New SqlCommand(sql, objconn)
        TIMS.OpenDbConn(objconn)
        With uCmd2
            .Parameters.Clear()
            .Parameters.Add("PlanID", SqlDbType.VarChar).Value = PlanID
            .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = ComIDNO
            .Parameters.Add("SeqNo", SqlDbType.VarChar).Value = SeqNo
            '.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(uCmd2.CommandText, objconn, uCmd2.Parameters)
        End With

        sql = ""
        sql &= " UPDATE PLAN_PLANINFO "
        sql &= " SET TMID=554,GCID2=1160"
        sql &= " WHERE 1=1"
        sql &= " AND PlanID=@PlanID"
        sql &= " AND ComIDNO=@ComIDNO"
        sql &= " AND SeqNo=@SeqNo"
        Dim uCmd As New SqlCommand(sql, objconn)
        With uCmd
            .Parameters.Clear()
            .Parameters.Add("PlanID", SqlDbType.VarChar).Value = PlanID
            .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = ComIDNO
            .Parameters.Add("SeqNo", SqlDbType.VarChar).Value = SeqNo
            '.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
        End With
    End Sub

    'INSERT Plan_Depot (SAVE)
    Sub sUtl_ChkOkSaveData(ByVal sSearchW As String)
        Dim sql As String = ""
        '確認
        Const cst_其他 As String = "19" 'KID12 課程分類 '19:其他類 ddlDepot12
        'sSearchW = e.CommandArgument
        Dim PlanID As String = TIMS.GetMyValue(sSearchW, "PlanID")
        Dim ComIDNO As String = TIMS.GetMyValue(sSearchW, "ComIDNO")
        Dim SeqNo As String = TIMS.GetMyValue(sSearchW, "SeqNo")
        Dim KID04 As String = TIMS.GetMyValue(sSearchW, "KID04")
        Dim KID06 As String = TIMS.GetMyValue(sSearchW, "KID06")
        Dim KID10 As String = TIMS.GetMyValue(sSearchW, "KID10")
        Dim KID12 As String = TIMS.GetMyValue(sSearchW, "KID12") '19:其他類
        Dim SEQNOD13 As String = TIMS.GetMyValue(sSearchW, "SEQNOD13")
        Dim KID14 As String = TIMS.GetMyValue(sSearchW, "KID14")
        If PlanID = "" Then Exit Sub
        If ComIDNO = "" Then Exit Sub
        If SeqNo = "" Then Exit Sub

        sql = "DELETE Plan_Depot WHERE PlanID ='" & PlanID & "' and ComIDNO='" & ComIDNO & "' and SeqNo ='" & SeqNo & "'"
        DbAccess.ExecuteNonQuery(sql, objconn)

        sql = "" & vbCrLf
        sql &= " INSERT INTO Plan_Depot(" & vbCrLf
        sql &= " PlanID" & vbCrLf
        sql &= " ,ComIDNO" & vbCrLf
        sql &= " ,SeqNo" & vbCrLf
        sql &= " ,KID04" & vbCrLf
        sql &= " ,KID06" & vbCrLf
        sql &= " ,KID10" & vbCrLf
        sql &= " ,KID12" & vbCrLf
        sql &= " ,SEQNOD13" & vbCrLf
        sql &= " ,KID14" & vbCrLf
        sql &= " ,AppResult" & vbCrLf
        sql &= " ,ModifyAcct" & vbCrLf
        sql &= " ,ModifyDate" & vbCrLf
        sql &= " ) VALUES (" & vbCrLf
        sql &= " @PlanID" & vbCrLf
        sql &= " ,@ComIDNO" & vbCrLf
        sql &= " ,@SeqNo" & vbCrLf
        sql &= " ,@KID04" & vbCrLf
        sql &= " ,@KID06" & vbCrLf
        sql &= " ,@KID10" & vbCrLf
        sql &= " ,@KID12" & vbCrLf
        sql &= " ,@SEQNOD13" & vbCrLf
        sql &= " ,@KID14" & vbCrLf
        sql &= " ,'Y'" & vbCrLf
        sql &= " ,@ModifyAcct" & vbCrLf
        sql &= " ,getdate()" & vbCrLf
        sql &= " ) " & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)
        TIMS.OpenDbConn(objconn)
        With iCmd
            .Parameters.Clear()
            .Parameters.Add("@PlanID", SqlDbType.VarChar).Value = PlanID
            .Parameters.Add("@ComIDNO", SqlDbType.VarChar).Value = ComIDNO
            .Parameters.Add("@SeqNo", SqlDbType.VarChar).Value = SeqNo
            .Parameters.Add("@KID04", SqlDbType.VarChar).Value = IIf(KID04 <> "", KID04, Convert.DBNull) 'KID04
            .Parameters.Add("@KID06", SqlDbType.VarChar).Value = IIf(KID06 <> "", KID06, Convert.DBNull) 'KID06
            .Parameters.Add("@KID10", SqlDbType.VarChar).Value = IIf(KID10 <> "", KID10, Convert.DBNull) 'KID10
            .Parameters.Add("@KID12", SqlDbType.VarChar).Value = IIf(KID12 <> "", KID12, Convert.DBNull) 'KID12
            .Parameters.Add("@SEQNOD13", SqlDbType.VarChar).Value = IIf(SEQNOD13 <> "", SEQNOD13, Convert.DBNull) 'SEQNOD13
            .Parameters.Add("@KID14", SqlDbType.VarChar).Value = IIf(KID14 <> "", KID14, Convert.DBNull) 'KID14
            .Parameters.Add("@ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            '.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
        End With

        If KID12 = cst_其他 Then
            '若選其他 異動 PLAN_PLANINFO
            Call sUtl_UpdatePlaninfo(sSearchW)
        End If
    End Sub

    '清除
    Sub sUtl_ClearPanelEdit1()
        Call ClearTrainDesc()

        hidPlanID.Value = ""
        hidComIDNO.Value = ""
        hidSeqNO.Value = ""

        ddlKID06.SelectedIndex = -1
        ddlKID04.SelectedIndex = -1
        ddlKID10.SelectedIndex = -1
        ddlDepot12.SelectedIndex = -1 '19:其他類
        ddlDEPOT13.SelectedIndex = -1
        ddlKID14.SelectedIndex = -1

        lbYears.Text = ""
        lbDistName.Text = ""
        lbOrgName.Text = ""
        lbClassName.Text = ""
        lbSFTDate.Text = ""
        lbTHours.Text = ""
        lbJobName.Text = ""
        lbGovClassN.Text = ""
    End Sub

    '顯示
    Sub sUtl_ShowPanelEdit1(ByVal sSearchW As String)
        panelSearch.Visible = False '搜尋功能關閉
        PanelEdit1.Visible = True '修改功能啟動

        '修改
        'sSearchW = e.CommandArgument
        Dim PlanID As String = TIMS.GetMyValue(sSearchW, "PlanID")
        Dim ComIDNO As String = TIMS.GetMyValue(sSearchW, "ComIDNO")
        Dim SeqNo As String = TIMS.GetMyValue(sSearchW, "SeqNo")

        Dim KID04 As String = TIMS.GetMyValue(sSearchW, "KID04")
        Dim KID06 As String = TIMS.GetMyValue(sSearchW, "KID06")
        Dim KID10 As String = TIMS.GetMyValue(sSearchW, "KID10")
        Dim KID12 As String = TIMS.GetMyValue(sSearchW, "KID12")
        Dim SEQNOD13 As String = TIMS.GetMyValue(sSearchW, "SEQNOD13")
        Dim KID14 As String = TIMS.GetMyValue(sSearchW, "KID14")
        If PlanID = "" Then Exit Sub
        If ComIDNO = "" Then Exit Sub
        If SeqNo = "" Then Exit Sub

        hidPlanID.Value = PlanID
        hidComIDNO.Value = ComIDNO
        hidSeqNO.Value = SeqNo
        If KID04 <> "" Then Common.SetListItem(ddlKID04, KID04)
        If KID06 <> "" Then Common.SetListItem(ddlKID06, KID06)
        If KID10 <> "" Then Common.SetListItem(ddlKID10, KID10)
        'ddlDepot12.SelectedIndex = -1
        If KID12 <> "" Then Common.SetListItem(ddlDepot12, KID12) '19:其他類
        If SEQNOD13 <> "" Then Common.SetListItem(ddlDEPOT13, SEQNOD13)
        If KID14 <> "" Then Common.SetListItem(ddlKID14, KID14)

        Call ShowTrainDesc(PlanID, ComIDNO, SeqNo)

        Dim dr As DataRow = search2(PlanID, ComIDNO, SeqNo)
        If Not dr Is Nothing Then
            lbYears.Text = dr("Years").ToString
            lbDistName.Text = dr("DistName").ToString
            lbOrgName.Text = dr("OrgName").ToString
            lbClassName.Text = dr("ClassName").ToString
            lbSFTDate.Text = Convert.ToString(dr("STDate")) & "~" & Convert.ToString(dr("FDDate"))
            lbTHours.Text = dr("THours").ToString
            lbJobName.Text = dr("JobName").ToString
            lbGovClassN.Text = dr("GovClassN").ToString
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Select Case e.CommandName
            Case "CHKOK"
                '確認'SAVE
                Dim sSearchW As String = ""
                sSearchW = e.CommandArgument
                Call sUtl_ChkOkSaveData(sSearchW)
                Call Search1()

            Case "Edit"
                '修改
                Dim sSearchW As String = ""
                sSearchW = e.CommandArgument
                Call sUtl_ClearPanelEdit1()
                Call sUtl_ShowPanelEdit1(sSearchW)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbD6KNAME As Label = e.Item.FindControl("lbD6KNAME")
                Dim lbD10KNAME As Label = e.Item.FindControl("lbD10KNAME")
                Dim lbD4KNAME As Label = e.Item.FindControl("lbD4KNAME")
                Dim lbD12KNAME As Label = e.Item.FindControl("lbD12KNAME")
                Dim lbD13KNAME As Label = e.Item.FindControl("lbD13KNAME")
                Dim lbD14KNAME As Label = e.Item.FindControl("lbD14KNAME")

                Dim BtnCHKOK As Button = e.Item.FindControl("BtnCHKOK")
                Dim BtnEdit As Button = e.Item.FindControl("BtnEdit")

                '序號
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex

                Const Cst_emptyTxt As String = "無"
                lbD6KNAME.Text = IIf(Convert.ToString(drv("D6KNAME")) <> "", Convert.ToString(drv("D6KNAME")), Cst_emptyTxt)
                lbD10KNAME.Text = IIf(Convert.ToString(drv("D10KNAME")) <> "", Convert.ToString(drv("D10KNAME")), Cst_emptyTxt)
                lbD4KNAME.Text = IIf(Convert.ToString(drv("D4KNAME")) <> "", Convert.ToString(drv("D4KNAME")), Cst_emptyTxt)
                '課程分類
                lbD12KNAME.Text = IIf(Convert.ToString(drv("D12KNAME")) <> "", Convert.ToString(drv("D12KNAME")), Cst_emptyTxt)
                lbD13KNAME.Text = IIf(Convert.ToString(drv("D13KNAME")) <> "", Convert.ToString(drv("D13KNAME")), Cst_emptyTxt)
                lbD14KNAME.Text = IIf(Convert.ToString(drv("D14KNAME")) <> "", Convert.ToString(drv("D14KNAME")), Cst_emptyTxt)

                Dim cmdArg As String = ""
                cmdArg = ""
                cmdArg &= "&PlanID=" & Convert.ToString(drv("PlanID"))
                cmdArg &= "&ComIDNO=" & Convert.ToString(drv("ComIDNO"))
                cmdArg &= "&SeqNo=" & Convert.ToString(drv("SeqNo"))
                cmdArg &= "&KID04=" & Convert.ToString(drv("KID04"))
                cmdArg &= "&KID06=" & Convert.ToString(drv("KID06"))
                cmdArg &= "&KID10=" & Convert.ToString(drv("KID10"))
                cmdArg &= "&KID12=" & Convert.ToString(drv("KID12"))
                cmdArg &= "&SEQNOD13=" & Convert.ToString(drv("SEQNOD13"))
                cmdArg &= "&KID14=" & Convert.ToString(drv("KID14"))

                BtnCHKOK.CommandArgument = cmdArg
                BtnEdit.CommandArgument = cmdArg

                BtnEdit.Enabled = True
                If Convert.ToString(drv("AppResult")) <> "Y" Then '未確認
                    BtnCHKOK.Enabled = True '使用確認鈕
                Else
                    BtnCHKOK.Enabled = False '停止確認鈕
                    TIMS.Tooltip(BtnCHKOK, "已確認")
                End If

        End Select
    End Sub

    '儲存
    Private Sub btnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim cmdArg As String = ""
        cmdArg = ""
        cmdArg &= "&PlanID=" & hidPlanID.Value
        cmdArg &= "&ComIDNO=" & hidComIDNO.Value
        cmdArg &= "&SeqNo=" & hidSeqNO.Value
        cmdArg &= "&KID04=" & ddlKID04.SelectedValue
        cmdArg &= "&KID06=" & ddlKID06.SelectedValue
        cmdArg &= "&KID10=" & ddlKID10.SelectedValue
        cmdArg &= "&KID12=" & ddlDepot12.SelectedValue '19:其他類
        cmdArg &= "&SEQNOD13=" & ddlDEPOT13.SelectedValue
        cmdArg &= "&KID14=" & ddlKID14.SelectedValue
        Call sUtl_ChkOkSaveData(cmdArg)
        Call Search1()
    End Sub

    '回上一頁
    Private Sub btnBack1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack1.Click
        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉
    End Sub

#Region "S"
    '課程大綱 清理
    Sub ClearTrainDesc()
        Dim sql As String = ""
        Dim dt As DataTable
        sql = ""
        sql &= " SELECT * "
        sql &= " FROM Plan_TrainDesc WHERE 1<>1"
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.DefaultView.Sort = "STrainDate asc,PName asc,PTDID asc"
        dt = TIMS.dv2dt(dt.DefaultView)
        With Datagrid3
            .DataSource = dt
            .DataKeyField = "PTDID"
            .DataBind()
        End With
    End Sub

    '課程大綱
    Sub ShowTrainDesc(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String)
        Dim sql As String = ""
        Dim dt As DataTable
        sql = ""
        sql &= " SELECT * "
        sql &= " FROM Plan_TrainDesc "
        sql &= " WHERE 1=1"
        sql &= " AND PlanID='" & PlanID & "' "
        sql &= " AND ComIDNO='" & ComIDNO & "' "
        sql &= " AND SeqNO='" & SeqNo & "' "
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.DefaultView.Sort = "STrainDate asc,PName asc,PTDID asc"
        dt = TIMS.dv2dt(dt.DefaultView)
        With Datagrid3
            .DataSource = dt
            .DataKeyField = "PTDID"
            .DataBind()
        End With
    End Sub

    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                'SHOW
                Dim drv As DataRowView = e.Item.DataItem
                Dim PHourLabel As Label = e.Item.FindControl("PHourLabel")
                Dim lbContText As Label = e.Item.FindControl("lbContText")
                Dim drpClassification1 As DropDownList = e.Item.FindControl("drpClassification1")

                PHourLabel.Text = Convert.ToString(drv("PHour")) '時數
                lbContText.Text = Convert.ToString(drv("PCont")) '內容
                If Convert.ToString(drv("Classification1")) <> "" Then
                    Common.SetListItem(drpClassification1, drv("Classification1").ToString)
                End If
        End Select
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
#End Region

End Class

