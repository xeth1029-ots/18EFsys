Partial Class SD_05_008
    Inherits AuthBasePage

    '一併修正。AC_03_001_add、SD_05_008
    'ResultStud_2 【非署(局)屬】

    'CheckPrint
    'ResultStud1 (列印學員空白資料卡) ResultStud*.jrxml
    'ResultStud 【署(局)屬】(資料列印)

    'ResultStud507 (201607適用)【署(局)屬】(列印學員空白資料卡)
    'ResultStud10507 (201607適用)【署(局)屬】(資料列印)

    'ResultStud_Title.jrxml '(列印封面) ResultStud*.jrxml

    'FROM STUD_RESULTSTUDDATA /STUD_DATALID
    'select count(1) cnt from STUD_DATALID where ocid is null-- 非署(局)屬
    'select count(1) cnt from STUD_DATALID where ocid is not null -- 署(局)屬
    'select count(1) cnt from STUD_DATALID where ocid =dlid

    'SELECT max(modifydate) x,COUNT(1) CNT from STUD_RESULTSTUDDATA
    'SELECT COUNT(1) CNT FROM STUD_RESULTIDENTDATA
    'SELECT COUNT(1) CNT FROM STUD_RESULTTWELVEDATA

    'SELECT * FROM KEY_IDENTITY
    'Stud_ResultTwelveData 2014不要這個。
    Const cst_search As String = "search"
    Const cst_Juzhu As String = "1" 'cst_Juzhu 1 署(局)屬 
    Const cst_NonJuzhu As String = "2" 'cst_NonJuzhu 2 非署(局)屬

    Const cst_printFN1 As String = "ResultStud_Title" '列印封面
    Const cst_printFN2 As String = "ResultStud" '勾選要列印的學員【署(局)屬】
    Const cst_printFN3 As String = "ResultStud10507" '勾選要列印的學員【署(局)屬】 'cst_NewSuySD20160701
    Const cst_printFN4 As String = "ResultStud_2" '勾選要列印的學員【非署(局)屬】

    Const cst_printFN5 As String = "ResultStud1" '列印學員空白資料卡【署(局)屬】
    Const cst_printFN6 As String = "ResultStud507" '列印學員空白資料卡【署(局)屬】 'cst_NewSuySD20160701

    Const cst_dg2列印 As Integer = 0
    Const cst_dg2學號 As Integer = 1
    Const cst_dg2姓名 As Integer = 2
    Const cst_dg2填寫狀態 As Integer = 3
    Const cst_dg2功能 As Integer = 4

    Dim tt As String = ""
    Dim Days1 As Integer
    Dim Days2 As Integer
    'Dim FunDr As DataRow
    Dim dtArc As DataTable
    Dim dtUnit As DataTable

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1

        Call TIMS.Get_SysDays(Days1, Days2)
        dtArc = TIMS.Get_Auth_REndClass(Me, objconn)
        dtUnit = TIMS.Get_dtUnit(objconn)

        If Not IsPostBack Then
            msg.Text = ""
            LabelMsg1.Text = ""
            LabelMsg2.Text = ""

            'Call SHOW_TPlan()
            TPlan = TIMS.Get_TPlan(TPlan, , 1)
            UnitCode = TIMS.Get_OthUnitCode(UnitCode, 1, objconn)

            SearchTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
            Button6.Style.Item("display") = "none"
            StudentTable.Style.Item("display") = "none"
            'FTDate2.Text = Now.Date
        End If

        '列印學員空白資料卡
        Me.Button12.Visible = False
        TR1_1.Style("display") = "none"
        TR1_2.Style("display") = "none"
        TR1_3.Style("display") = "none"
        TR1_4.Style("display") = "none"
        TR2_1.Style("display") = "none"
        TR2_2.Style("display") = "none"

        Dim v_RadioButtonList1 As String = TIMS.GetListValue(RadioButtonList1)
        Select Case v_RadioButtonList1' RadioButtonList1.SelectedValue
            Case cst_Juzhu '署(局)屬 
                TR1_1.Style("display") = TIMS.cst_inline1 '"inline"
                TR1_2.Style("display") = TIMS.cst_inline1 '"inline" '署(局)
                TR1_3.Style("display") = TIMS.cst_inline1 '"inline"
                TR1_4.Style("display") = TIMS.cst_inline1 '"inline"
                TR2_1.Style("display") = "none"
                TR2_2.Style("display") = "none"
                Me.Button12.Visible = True
            Case cst_NonJuzhu '非署(局)屬
                TR1_1.Style("display") = "none"
                TR1_2.Style("display") = "none" '非署(局)
                TR1_3.Style("display") = "none"
                TR1_4.Style("display") = "none"
                TR2_1.Style("display") = TIMS.cst_inline1 '"inline"
                TR2_2.Style("display") = TIMS.cst_inline1 '"inline"
        End Select

        Button1.Attributes("onclick") = "javascript:return search()"
        RadioButtonList1.Attributes("onclick") = "choose(null)"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button11.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');showFrame();"
            center.Style("CURSOR") = "hand"
            HistoryRID.Attributes("onclick") = "showFrame();"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '結訓學員匯出功能
        But1.Enabled = If(sm.UserInfo.RID = "A", True, False) ' True
        TIMS.Tooltip(But1, If(sm.UserInfo.RID = "A", "", "權限不足，請使用署登入"), True)

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PageControler1.Visible = False
            If sm.UserInfo.LID <> 0 Then
                RadioButtonList1.Items(0).Selected = True '署(局)屬
                RadioButtonList1.Enabled = False
                Page.RegisterStartupScript("aaaa", "<script>choose(null);</script>")
                MainTr.Style("display") = "none"
            Else
                MainTr.Style("display") = TIMS.cst_inline1 '"inline" '非局屬(開啟)
            End If
            Common.SetListItem(TPlan, sm.UserInfo.TPlanID)

            Call USE_KEEPSEARCH()
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            ' TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If
    End Sub

    Sub USE_KEEPSEARCH()
        If Session(cst_search) Is Nothing Then Return
        If Convert.ToString(Session(cst_search)) = "" Then Return

        Dim s_SEARCH1 As String = Convert.ToString(Session(cst_search))

        Session(cst_search) = Nothing
        Dim MyValue As String = ""
        MyValue = TIMS.GetMyValue(s_SEARCH1, "RadioButtonList1")
        Common.SetListItem(RadioButtonList1, MyValue)

        center.Text = TIMS.GetMyValue(s_SEARCH1, "center")
        RIDValue.Value = TIMS.GetMyValue(s_SEARCH1, "RIDValue")
        MyValue = TIMS.GetMyValue(s_SEARCH1, "TPlan")
        Common.SetListItem(TPlan, MyValue)
        TMID1.Text = TIMS.GetMyValue(s_SEARCH1, "TMID1")
        OCID1.Text = TIMS.GetMyValue(s_SEARCH1, "OCID1")
        TMIDValue1.Value = TIMS.GetMyValue(s_SEARCH1, "TMIDValue1")
        OCIDValue1.Value = TIMS.GetMyValue(s_SEARCH1, "OCIDValue1")

        FTDate1.Text = TIMS.GetMyValue(s_SEARCH1, "FTDate1")
        FTDate2.Text = TIMS.GetMyValue(s_SEARCH1, "FTDate2")

        MyValue = TIMS.GetMyValue(s_SEARCH1, "UnitCode")
        Common.SetListItem(UnitCode, MyValue)
        MyValue = TIMS.GetMyValue(s_SEARCH1, "PageIndex")
        If MyValue <> "" AndAlso IsNumeric(MyValue) Then
            MyValue = CInt(MyValue)
            PageControler1.PageIndex = MyValue
        End If

        Dim flagC1 As Boolean = False 'True'已經使用班級查詢動作。
        Dim flag_studlist As Boolean = False  '執行查詢。
        MyValue = TIMS.GetMyValue(s_SEARCH1, "StudentTable")
        If MyValue = "inline" Then flag_studlist = True

        MyValue = TIMS.GetMyValue(s_SEARCH1, "OCID")
        If MyValue <> "" Then
            ShowStudData(MyValue, "") '已查詢
            flag_studlist = False '不動作
            flagC1 = True '已經使用班級查詢動作。
        End If
        MyValue = TIMS.GetMyValue(s_SEARCH1, "DLID")
        If Not flagC1 AndAlso MyValue <> "" Then
            '未使用班級查詢動作。
            ShowStudData("", MyValue) '已查詢
            flag_studlist = False '不動作
        End If

        Dim v_RadioButtonList1 As String = TIMS.GetListValue(RadioButtonList1)
        If flag_studlist Then
            'Button1_Click(sender, e) '執行查詢動作。
            Call ShowClass(v_RadioButtonList1)
            Page.RegisterStartupScript("asddq", "<script>choose(1);</script>")
        Else
            Page.RegisterStartupScript("asddq", "<script>choose(2);</script>")
        End If
        'Session(cst_search) = Nothing
    End Sub

    ''' <summary>顯示班級資訊 依-rblJuzhu</summary>
    ''' <param name="rblJuzhu"></param>
    Sub ShowClass(ByVal rblJuzhu As String)
        'Dim sql As String
        'rblJuzhu: 1@cst_Juzhu '署(局)屬 2@cst_NonJuzhu '非署(局)屬
        Select Case rblJuzhu 'RadioButtonList1.SelectedValue
            Case cst_Juzhu '"1" '署(局)屬
                Button6.Style.Item("display") = "none"

                RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
                '若2個都空白就不尋找
                If OCIDValue1.Value = "" AndAlso RIDValue.Value = "" Then Exit Sub

                RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

                Dim sql As String = ""
                sql &= " WITH WM1 AS ( SELECT p.DLID" & vbCrLf
                sql &= " ,COUNT(distinct cs.socid) Total" & vbCrLf '計算學員 (1個學員計算1次)
                sql &= " ,COUNT(p.SubNO) num" & vbCrLf '計算填寫數 (重複計算)
                sql &= " FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
                sql &= " JOIN CLASS_CLASSINFO cc on cc.ocid =cs.ocid" & vbCrLf
                sql &= " JOIN ID_PLAN ip ON ip.PlanID=cc.PlanID" & vbCrLf
                sql &= " JOIN VIEW_RIDNAME f ON cc.RID=f.RID" & vbCrLf
                sql &= " LEFT JOIN STUD_RESULTSTUDDATA p on cs.socid =p.socid" & vbCrLf
                sql &= " where cs.StudStatus NOT IN (2,3)" & vbCrLf
                If OCIDValue1.Value <> "" Then
                    sql &= " AND cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
                End If
                sql &= " AND ip.Years ='" & sm.UserInfo.Years & "'" & vbCrLf
                '署(局)屬
                If sm.UserInfo.RID = "A" Then
                    '依登入計畫,年度
                    sql &= " AND ip.TPlanID ='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                Else
                    '依登入計畫
                    'sql += " AND ip.PlanID ='" & sm.UserInfo.PlanID & "'" & vbCrLf
                    If RIDValue.Value.Length > 1 Then
                        '有指定機構
                        sql &= " AND cc.RID='" & RIDValue.Value & "'" & vbCrLf
                    Else
                        '未指定機構,依登入計畫
                        sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                        sql &= " AND f.RELSHIP like '" & Relship & "%'" & vbCrLf
                    End If
                End If
                sql &= " GROUP BY p.DLID" & vbCrLf
                sql &= " )" & vbCrLf

                sql &= " select f.OrgName" & vbCrLf
                sql &= " ,f.RID" & vbCrLf
                sql &= " ,f.Relship" & vbCrLf
                sql &= " ,a.OCID" & vbCrLf
                sql &= " ,a.CyclType" & vbCrLf
                sql &= " ,a.LevelType" & vbCrLf
                sql &= " ,a.ClassCName" & vbCrLf
                sql &= " ,a.FTDate" & vbCrLf
                sql &= " ,ISNULL(d.Total,0) Total" & vbCrLf
                sql &= " ,c.DLID" & vbCrLf
                sql &= " ,ISNULL(d.num,0) num" & vbCrLf
                sql &= " ,e.ClassID " & vbCrLf
                sql &= " FROM Class_ClassInfo a " & vbCrLf
                sql &= " JOIN ID_Class e ON a.CLSID=e.CLSID " & vbCrLf
                sql &= " JOIN ID_Plan ip ON ip.PlanID=a.PlanID" & vbCrLf
                sql &= " JOIN VIEW_RIDNAME f ON a.RID=f.RID AND f.Relship like '" & Relship & "%'" & vbCrLf
                sql &= " LEFT JOIN STUD_DATALID c ON a.OCID=c.OCID" & vbCrLf
                sql &= " LEFT JOIN WM1 d ON c.DLID=d.DLID" & vbCrLf
                sql &= " WHERE ip.Years ='" & sm.UserInfo.Years & "'" & vbCrLf
                '署(局)屬
                If sm.UserInfo.RID = "A" Then
                    '依登入計畫,年度
                    sql &= " AND ip.TPlanID ='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                Else
                    '依登入計畫
                    'sql += " AND ip.PlanID ='" & sm.UserInfo.PlanID & "'" & vbCrLf
                    If RIDValue.Value.Length > 1 Then
                        '有指定機構
                        sql &= " AND a.RID='" & RIDValue.Value & "'" & vbCrLf
                    Else
                        '未指定機構,依登入計畫
                        sql &= " AND a.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                        sql &= " AND f.RELSHIP like '" & Relship & "%'" & vbCrLf
                    End If
                End If
                ''署(局)屬
                'If TPlan.SelectedValue <> "" Then
                '    sql &= " AND ip.TPlanID ='" & TPlan.SelectedValue & "'" & vbCrLf
                'End If
                If OCIDValue1.Value <> "" Then
                    sql &= " and a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
                End If

                If FTDate1.Text <> "" Then
                    sql &= " AND a.FTDate>= " & TIMS.To_date(FTDate1.Text) & vbCrLf
                End If
                If FTDate2.Text <> "" Then
                    sql &= " AND a.FTDate<= " & TIMS.To_date(FTDate2.Text) & vbCrLf 'convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf
                End If
                If cjobValue.Value <> "" Then
                    If sm.UserInfo.RID <> "A" Then
                        sql &= " and a.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                    End If
                    sql &= " and a.CJOB_UNKEY=" & cjobValue.Value & vbCrLf
                End If
                sql &= " ORDER BY a.FTDATE"

                'SELECT B.*
                'from Class_ClassInfo a 
                'JOIN ID_Class e ON a.CLSID=e.CLSID 
                'JOIN ID_Plan ip ON ip.PlanID=a.PlanID AND ip.Years='2011' 
                'JOIN view_RIDName f ON a.RID=f.RID AND f.Relship like 'A/G/G1459/%'
                'JOIN Stud_DataLid c ON a.OCID=c.OCID
                'JOIN Stud_ResultStudData   d ON c.DLID=d.DLID 
                'JOIN Class_StudentsOfClass  b ON D.SOCID=b.SOCID 
                'da.SelectCommand.Parameters.Clear()
                'msg.Text = "查無資料!!"
                'DataGrid1.Style.Item("display") = "none"
                'PageControler1.Visible = False
                'Label3.Visible = False

                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

                msg.Text = "查無資料!!"
                DataGrid1.Style.Item("display") = "none"
                PageControler1.Visible = False
                Label3.Visible = False

                If dt.Rows.Count > 0 Then
                    msg.Text = ""
                    DataGrid1.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    PageControler1.Visible = True
                    Label3.Visible = True

                    'Label3.Visible = True
                    Label3.Text = "紅色【*】表示:此班級有學員還沒填寫問卷或是填寫的問卷作答有遺漏"
                    Label3.ForeColor = Color.Red
                    DataGrid1.PageSize = 10
                    DataGrid1.AllowPaging = True

                    DataGrid1.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    DataGrid1.Columns(1).Visible = True
                    DataGrid1.DataKeyField = "OCID"

                    PageControler1.PageDataTable = dt
                    PageControler1.PrimaryKey = "OCID"
                    PageControler1.Sort = "ClassID,CyclType"
                    PageControler1.ControlerLoad()
                End If

            Case Else 'cst_NonJuzhu 非署(局)屬
                '非署(局)屬
                Button6.Visible = True
                Button6.Style.Item("display") = TIMS.cst_inline1 '"inline"

                ClassName.Text = TIMS.ClearSQM(ClassName.Text)

                Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " WITH WM1 AS (" & vbCrLf
                sql &= " select a.DLID" & vbCrLf
                'sql &= " ,COUNT(a.SUBNO) num" & vbCrLf '填寫學員數。
                sql &= " ,COUNT(dbo.fn_CHK_RESULTSTUDDATA(a.DLID,a.SUBNO,1)) num" & vbCrLf '填寫學員數。
                sql &= " from STUD_RESULTSTUDDATA a" & vbCrLf
                sql &= " JOIN STUD_DATALID b on a.dlid=b.dlid" & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= " and b.OCID is null " & vbCrLf
                sql &= " and b.unitcode='" & UnitCode.SelectedValue & "'" & vbCrLf
                If ClassName.Text <> "" Then
                    sql &= " and b.classname like '%'+'" + ClassName.Text + "'+'%'" & vbCrLf
                End If
                If FTDate1.Text <> "" Then
                    sql &= " AND b.resultdate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf '
                End If
                If FTDate2.Text <> "" Then
                    sql &= " AND b.resultdate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf 'convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf
                End If
                sql &= " group by a.dlid" & vbCrLf
                sql &= " )" & vbCrLf

                sql &= " select b.OCID" & vbCrLf
                sql &= " ,b.unitcode orgname" & vbCrLf
                sql &= " ,b.classname classcname" & vbCrLf
                sql &= " ,b.DLID" & vbCrLf
                sql &= " ,ISNULL(a.num,0) num" & vbCrLf
                sql &= " ,b.resultcount total" & vbCrLf
                sql &= " ,b.resultdate ftdate" & vbCrLf
                sql &= " ,0 cycltype"
                sql &= " ,0 leveltype"
                sql &= " from STUD_DATALID b " & vbCrLf
                sql &= " left join WM1 a on a.dlid=b.dlid" & vbCrLf
                'sql += " 	select DLID,COUNT(SUBNO) num " & vbCrLf '填寫學員數。
                'sql += " 	from STUD_RESULTSTUDDATA " & vbCrLf
                'sql += " 	group by dlid) a on a.dlid=b.dlid " & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= " and b.ocid is null " & vbCrLf
                sql &= " and b.unitcode='" & UnitCode.SelectedValue & "'" & vbCrLf
                'If ClassName.Text <> "" Then ClassName.Text = TIMS.ClearSQM(ClassName.Text)
                If ClassName.Text <> "" Then
                    sql &= " and b.classname like '%'+'" + ClassName.Text + "'+'%'" & vbCrLf
                End If
                If FTDate1.Text <> "" Then
                    sql &= " AND b.resultdate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf '
                End If
                If FTDate2.Text <> "" Then
                    sql &= " AND b.resultdate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf 'convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf
                End If
                sql &= " order by b.resultdate" & vbCrLf

                Dim dt As DataTable
                dt = DbAccess.GetDataTable(sql, objconn)

                Label3.Visible = False
                msg.Text = "查無資料!!"
                DataGrid1.Style.Item("display") = "none"
                PageControler1.Visible = False
                If dt.Rows.Count > 0 Then
                    'Label3.Visible = False
                    msg.Text = ""
                    DataGrid1.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    PageControler1.Visible = True

                    DataGrid1.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    DataGrid1.DataKeyField = "DLID"
                    DataGrid1.Columns(1).Visible = False

                    'PageControler1.SqlPrimaryKeyDataCreate(Sql, "DLID")
                    PageControler1.PageDataTable = dt
                    PageControler1.PrimaryKey = "DLID"
                    PageControler1.ControlerLoad()
                End If
        End Select

    End Sub

    'show Data
    ''' <summary>
    ''' 顯示班級-學員資訊  依-vOCID/vDLID
    ''' </summary>
    ''' <param name="vOCID"></param>
    ''' <param name="vDLID"></param>
    Sub ShowStudData(ByVal vOCID As String, ByVal vDLID As String)
        StudentTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
        SearchTable.Style.Item("display") = "none"
        ClassTable.Style.Item("display") = "none"
        OCID.Value = If(vOCID <> "", vOCID, "")
        DLID.Value = If(vDLID <> "", vDLID, "")

        If vOCID <> "" Then
            '署(局)屬查詢
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql &= " " & vbCrLf
            sql &= " SELECT a.*" & vbCrLf
            sql &= " ,ISNULL(b.StudentCount,0) StudentCount" & vbCrLf '開訓人數
            sql &= " ,ISNULL(b.TrainCount,0) TrainCount" & vbCrLf '結訓人數
            sql &= " ,ISNULL(b.LeaveCount,0) LeaveCount " & vbCrLf '離退人數
            sql &= " FROM Class_ClassInfo a" & vbCrLf
            sql &= " LEFT JOIN (	" & vbCrLf
            sql &= " 	SELECT cs.OCID" & vbCrLf
            sql &= " 	,Count(1) StudentCount " & vbCrLf '開訓人數
            sql &= " 	,sum(case when cs.StudStatus NOT IN (2,3) then 1 end ) TrainCount " & vbCrLf '結訓人數
            sql &= " 	,sum(case when cs.StudStatus IN (2,3) then 1 end ) LeaveCount " & vbCrLf '離退人數
            sql &= " 	FROM Class_StudentsOfClass cs" & vbCrLf
            sql &= " 	WHERE cs.OCID='" & vOCID & "' " & vbCrLf
            sql &= " 	Group By cs.OCID) b ON a.OCID=b.OCID" & vbCrLf
            sql &= " where 1=1" & vbCrLf
            sql &= " and a.OCID='" & vOCID & "'" & vbCrLf
            Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
            LabelMsg1.Text = ""
            If dr IsNot Nothing Then
                LabelMsg1.Text = "班別：" & TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))
                LabelMsg1.Text += "(開訓人數:" & dr("StudentCount").ToString & "&nbsp;&nbsp;&nbsp;&nbsp;在結訓人數:" & dr("TrainCount").ToString & "&nbsp;&nbsp;&nbsp;&nbsp;離退訓人數:" & dr("LeaveCount").ToString & ")"
            End If
            '學員基本資料。 
            sql = "" & vbCrLf
            sql &= " SELECT a.Name" & vbCrLf
            sql &= " ,s.SOCID" & vbCrLf '(Class_StudentsOfClass)
            sql &= " ,d.DLID" & vbCrLf '(Stud_ResultStudData)
            sql &= " ,sl.DLID DLID2" & vbCrLf '(Stud_DataLid)
            sql &= " ,d.SubNo" & vbCrLf '(Stud_ResultStudData)
            sql &= " ,s.OCID" & vbCrLf
            sql &= " ,s.StudentID" & vbCrLf
            sql &= " ,s.RejectTDate1" & vbCrLf
            sql &= " ,s.RejectTDate2" & vbCrLf
            sql &= " ,e.FTDate" & vbCrLf
            sql &= " ,d.StudentID StudentID2 " & vbCrLf

            sql &= " FROM STUD_STUDENTINFO a " & vbCrLf
            sql &= " JOIN CLASS_STUDENTSOFCLASS s on a.SID =s.SID and s.OCID ='" & vOCID & "'" & vbCrLf
            sql &= " JOIN CLASS_CLASSINFO e on s.OCID=e.OCID and e.OCID ='" & vOCID & "'" & vbCrLf
            sql &= " LEFT JOIN STUD_DATALID sL on sL.OCID =e.OCID" & vbCrLf '結訓封面卡。
            sql &= " LEFT JOIN STUD_RESULTSTUDDATA d on d.SOCID=s.SOCID" & vbCrLf '學員。
            sql &= " Order BY s.StudentID" & vbCrLf
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

            StudentTable.Style.Item("display") = "none"
            LabelMsg2.Text = ""
            'LabelMsg2.Visible = False
            If dt.Rows.Count > 0 Then
                StudentTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
                'LabelMsg2.Visible = True

                LabelMsg2.Text = "紅色【*】表示:此學員還沒填寫問卷或是填寫的問卷作答有遺漏"
                LabelMsg2.ForeColor = Color.Red

                DataGrid2.DataSource = dt
                DataGrid2.DataKeyField = "SOCID"
                DataGrid2.DataBind()

                'CHK 結訓資料卡封面
                For i As Integer = 0 To dt.Rows.Count - 1
                    'If Me.OCID.Value = "" AndAlso Convert.ToString(dt.Rows(i)("OCID")) <> "" Then
                    '    Me.OCID.Value = Convert.ToString(dt.Rows(i)("OCID")) '有DLID 也有OCID
                    'End If
                    If Convert.ToString(dt.Rows(i)("DLID")) <> "" Then
                        DLID.Value = Convert.ToString(dt.Rows(i)("DLID")) '有DLID 也有OCID
                        Exit For
                    End If
                Next
                If DLID.Value = "" Then
                    Button7.Enabled = False
                    TIMS.Tooltip(Button7, "結訓資料卡封面檔有誤!!", True)
                End If
            End If

        Else
            '非署(局)屬。
            'Me.ViewState("OCID") = ""
            'Me.ViewState("DLID") = vDLID
            'Me.DLID.Value = vDLID
            'Me.OCID.Value = vOCID

            Dim sql As String
            sql = "SELECT * FROM STUD_DATALID WHERE DLID='" & vDLID & "'"
            Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
            LabelMsg1.Text = ""
            If dr IsNot Nothing Then LabelMsg1.Text = String.Concat("班別：", dr("ClassName"))

            '非署(局)屬。
            sql = "" & vbCrLf
            sql &= " SELECT a.DLID" & vbCrLf
            sql &= " ,a.SubNo" & vbCrLf
            sql &= " ,a.StdName Name" & vbCrLf
            sql &= " ,0 OCID" & vbCrLf
            sql &= " ,a.StudentID " & vbCrLf
            sql &= " ,a.SOCID" & vbCrLf
            sql &= " ,dbo.fn_CHK_RESULTSTUDDATA(a.DLID,a.SUBNO,1) CHK_SUBNO2" & vbCrLf
            sql &= " FROM STUD_RESULTSTUDDATA a" & vbCrLf
            sql &= " WHERE a.DLID='" & vDLID & "'"
            sql &= " ORDER BY a.SubNo"
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

            'LabelMsg2.Visible = False
            LabelMsg2.Text = ""
            msg.Text = "查此班無學生資料!!"
            StudentTable.Style.Item("display") = "none"
            SearchTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
            'Me.ViewState("DLID") = ""
            If dt.Rows.Count > 0 Then
                'Label2.Visible = False
                msg.Text = ""
                StudentTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
                SearchTable.Style.Item("display") = "none"
                'Me.ViewState("DLID") = Me.DLID.Value

                DataGrid2.DataSource = dt
                DataGrid2.DataKeyField = "" '有必要清空。
                DataGrid2.DataBind()
            End If
        End If

        'Dim sPath As String = ""
        'sPath = TIMS.Server_Path()
        'If vOCID = "" AndAlso vDLID <> "" Then
        '    '非署(局)屬 【ResultStud_2】
        '    Button7.Attributes("onclick") = "CheckPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=MultiBlock&filename=ResultStud_2&path=" & sPath & "&');return false;"
        'Else
        '    '署(局)屬 【ResultStud】
        '    Button7.Attributes("onclick") = "CheckPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=MultiBlock&filename=ResultStud&path=" & sPath & "&');return false;"
        'End If
        'Button12.Attributes("onclick") = "CheckPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=MultiBlock&filename=ResultStud1&path=" & sPath & "&');return false;"
        '非署(局)屬 (ResultStud_2) '署(局)屬 【ResultStud】 列印結訓學員資料卡 
        Button7.Attributes("onclick") = "return CheckPrint();"
        '列印學員空白資料卡 ResultStud1
        Button12.Attributes("onclick") = "return CheckPrint();"
    End Sub

    ''' <summary> '查詢 (班級。)-顯示班級資訊 依-rblJuzhu </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ClassTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
        'Me.ViewState("RadioButtonList1") = RadioButtonList1.SelectedValue '1:署(局)屬 2:非署(局)屬
        'Me.ViewState("OCID") = ""
        'Me.ViewState("DLID") = ""
        Me.DLID.Value = ""
        Me.OCID.Value = ""
        Call ShowClass(TIMS.GetListValue(RadioButtonList1))
    End Sub

    ''' <summary>
    ''' Class list-command
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'view(查詢。OCID或DLID 學員)/edit(修改封面)/add(新增封面)
        Dim v_RadioButtonList1 As String = TIMS.GetListValue(RadioButtonList1)
        Select Case e.CommandName
            Case "addstd", "viewstd"  '查詢。OCID或DLID 學員
                Me.OCID.Value = ""
                Me.DLID.Value = ""
                Select Case v_RadioButtonList1 'RadioButtonList1.SelectedValue
                    Case cst_Juzhu 'cst_Juzhu 1 署(局)屬 
                        Call ShowStudData(e.CommandArgument, "")
                    Case Else  'cst_NonJuzhu 2 非署(局)屬
                        Call ShowStudData("", e.CommandArgument)
                End Select
            Case "edit" '修改封面。
                Dim ss As String = e.CommandArgument
                Me.OCID.Value = TIMS.GetMyValue(ss, "OCID")
                Me.DLID.Value = TIMS.GetMyValue(ss, "DLID")
                KeepSearchStr()
                TIMS.Utl_Redirect1(Me, e.CommandArgument)
            Case "add" '新增封面 , "edit"
                Dim ss As String = e.CommandArgument
                Me.OCID.Value = TIMS.GetMyValue(ss, "OCID")
                Me.DLID.Value = TIMS.GetMyValue(ss, "DLID")
                KeepSearchStr()
                TIMS.Utl_Redirect1(Me, e.CommandArgument)
            Case "print" '列印


        End Select
    End Sub

    ''' <summary>Class list</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        '署(局)屬
        '新增-沒有封面  導向封面建立  addmy         C頁面   OCID
        '新增-有封面,導向所有學生建立 addmyall      D頁面   DLID
        '查詢-OCID當鍵值搜尋DataGrid2

        '非署(局)屬
        '新增-有封面,導向所有學生建立 addother      D頁面   DLID
        '修改-有封面,導向所有學生建立 edit          C頁面   DLID
        '查詢-DLID當鍵值搜尋DataGrid2
        'view(查詢。OCID或DLID 學員)/edit(修改封面)/add(新增封面)

        'CommandArgument
        '新增。
        Dim sJuzhu As String = Get_sJuzhu() '取得署(局)屬傳遞參數。
        Dim v_RadioButtonList1 As String = TIMS.GetListValue(RadioButtonList1)
        Dim s_MRqID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        Dim sUrl As String = ""

        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1

                Dim drv As DataRowView = e.Item.DataItem
                Dim LinkBtn2_edit As LinkButton = e.Item.FindControl("LinkBtn2_edit") 'edit 可更改封面。'修改 edit (封面)
                Dim LabStart As Label = e.Item.FindControl("LabStart")

                '列印封面按鈕2005/5/25-Melody
                Dim Btn1_view_std As Button = e.Item.FindControl("Btn1_view_std") '查詢 (學員) viewstd
                Dim Btn1_add_std As Button = e.Item.FindControl("Btn1_add_std")   '新增 (學員) addstd
                Dim Btn2_add As Button = e.Item.FindControl("Btn2_add")           '新增 add (封面)
                Dim Btn3_print As Button = e.Item.FindControl("Btn3_print")       '列印 print
                Btn1_view_std.Visible = False
                Btn1_add_std.Visible = False
                Btn2_add.Visible = False
                Btn3_print.Visible = False

                Dim s_CyclType As String = If(IsNumeric(drv("CyclType")) AndAlso Int(drv("CyclType")) <> 0, String.Concat("第", drv("CyclType"), "期"), "")
                Dim s_ClassName As String = String.Concat(drv("ClassCName"), s_CyclType) '班級名稱
                LinkBtn2_edit.Text = s_ClassName

                Select Case v_RadioButtonList1 'RadioButtonList1.SelectedValue
                    Case cst_Juzhu '"1" 'cst_Juzhu 1 署(局)屬 
                        If Convert.ToString("OCID") <> "" Then
                            '檢查是否有未填寫的學員或題目
                            'False:'表示都沒有未填寫的 ／True:'表示有未填寫的
                            Dim fg_NG As Boolean = SD_05_008_D.Check_Qdata(drv("OCID"), "", objconn)
                            If Not fg_NG Then
                                LabStart.Text = ""
                                LabStart.Visible = False
                                'myLink.Text = drv("ClassCName").ToString '班級名稱
                            Else
                                '紅色【*】表示:此班級有學員還沒填寫問卷或是填寫的問卷作答有遺漏
                                LabStart.Text = "*"
                                LabStart.ForeColor = Color.Red
                                LabStart.Visible = True '有未填資料。
                                TIMS.Tooltip(LabStart, "有未填寫的學員或題目")
                                'myLink.Text = drv("ClassCName").ToString '班級名稱
                            End If
                        End If

                        Dim ParentRID As String = TIMS.Get_ParentRID(Convert.ToString(drv("Relship")), objconn)
                        If ParentRID <> "" Then e.Item.Cells(1).Text = String.Concat(ParentRID, "-", drv("OrgName")) 'e.Item.Cells(1).Text

                        '署(局)屬
                        Btn2_add.Visible = If(Convert.ToString(drv("DLID")) <> "", False, True)
                        LinkBtn2_edit.ForeColor = Color.FromName("#336699")
                        If Convert.ToString(drv("DLID")) = "" Then
                            LinkBtn2_edit.ForeColor = Color.Black
                            LinkBtn2_edit.Attributes("onclick") = "alert('封面尚未建立，請按「新增」');return false;"
                        Else
                            sUrl = ""
                            sUrl &= String.Concat("SD_05_008_C.aspx?ID=", s_MRqID, sJuzhu)
                            sUrl &= String.Concat("&Proecess=editmy", "&OCID=", drv("OCID"), "&DLID=", drv("DLID"))
                            LinkBtn2_edit.CommandArgument = sUrl
                        End If
                        Btn1_view_std.Visible = If(Convert.ToString(drv("DLID")) <> "", True, False)


                    Case Else 'cst_NonJuzhu 2 非署(局)屬
                        'UnitCode
                        LabStart.Text = ""               '星號
                        LabStart.Visible = False
                        'myLink.Text = drv("ClassCName").ToString '班級名稱
                        'UnitCode OrgName
                        e.Item.Cells(1).Text = TIMS.Get_UnitCodeOrgName(Convert.ToString(drv("OrgName")), dtUnit)

                        Btn2_add.Visible = True
                        If Convert.ToString(drv("DLID")) <> "" Then Btn2_add.Visible = False
                        '非署(局)屬
                        LinkBtn2_edit.ForeColor = Color.FromName("#336699")
                        If Convert.ToString(drv("DLID")) = "" Then
                            LinkBtn2_edit.ForeColor = Color.Black
                            LinkBtn2_edit.Attributes("onclick") = "alert('封面尚未建立，請按「新增」');return false;"
                        Else
                            sUrl = ""
                            sUrl &= "SD_05_008_C.aspx?ID=" & s_MRqID 'Request("ID")
                            sUrl &= sJuzhu
                            sUrl &= "&Proecess=editother"
                            sUrl &= "&DLID=" & Convert.ToString(drv("DLID"))
                            LinkBtn2_edit.CommandArgument = sUrl
                        End If
                        If Not Btn2_add.Visible Then
                            Btn1_view_std.Visible = True
                            If drv("num") = 0 Then Btn1_view_std.Visible = False
                        End If
                End Select

                Select Case v_RadioButtonList1
                    Case cst_Juzhu 'cst_Juzhu 1 署(局)屬 
                        Btn1_view_std.CommandArgument = CStr(drv("OCID")) 'DataGrid1.DataKeys(e.Item.ItemIndex)
                    Case Else  'cst_NonJuzhu 2 非署(局)屬
                        Btn1_view_std.CommandArgument = CStr(drv("DLID")) 'DataGrid1.DataKeys(e.Item.ItemIndex)
                End Select
                'Select Case RadioButtonList1.SelectedValue
                '    Case "1" '署(局)屬
                '    Case Else '非署(局)屬
                'End Select

                Select Case v_RadioButtonList1
                    Case cst_Juzhu '"1"  'cst_Juzhu 1 署(局)屬 【addmy, addmyall】
                        If drv("DLID").ToString = "" Then  '沒有封面(新增封面)
                            sUrl = ""
                            sUrl &= String.Concat("SD_05_008_C.aspx?ID=", s_MRqID, sJuzhu)
                            sUrl &= String.Concat("&Proecess=addmy", "&OCID=", drv("OCID"))
                            Btn2_add.CommandArgument = sUrl
                        Else
                            '有封面(新增學員)
                            sUrl = ""
                            sUrl &= String.Concat("SD_05_008_D.aspx?ID=", s_MRqID, sJuzhu)
                            sUrl &= String.Concat("&Proecess=addmyall", "&DLID=", drv("DLID"), "&OCID=", drv("OCID"))
                            Btn1_add_std.CommandArgument = sUrl
                        End If

                    Case Else 'cst_NonJuzhu 2 非署(局)屬 【addother】 '有封面(新增學員)
                        sUrl = ""
                        sUrl &= String.Concat("SD_05_008_D.aspx?ID=", s_MRqID, sJuzhu)
                        sUrl &= String.Concat("&Proecess=addother", "&DLID=", drv("DLID")) 'DataGrid1.DataKeys(e.Item.ItemIndex)
                        Btn1_add_std.CommandArgument = sUrl
                End Select

                'Select Case RadioButtonList1.SelectedValue
                '    Case cst_Juzhu '"1"  'cst_Juzhu 1 署(局)屬 【addmy, addmyall】
                '    Case Else 'cst_NonJuzhu 2 非局屬 (addother)
                'End Select

                If v_RadioButtonList1 = cst_Juzhu Then
                    Dim flag_full As Boolean = False 'true:滿了 false:未滿
                    If Convert.ToString(drv("total")) = 0 Then flag_full = True 'true:滿了 false:未滿

                    If Convert.ToString(drv("num")) <> 0 AndAlso Convert.ToString(drv("num")) = Convert.ToString(drv("total")) Then
                        Btn1_add_std.Visible = False
                        TIMS.Tooltip(e.Item, "填寫總數與學員相符，停用新增功能")
                        flag_full = True 'true:滿了 false:未滿
                    End If
                    If Not flag_full Then
                        '班級人數尚未填寫(未相符)。
                        '登入者要是非署(局)，且結續超過兩週則不給予新增資料，有結訓班級權限者不在此限
                        If Not TIMS.Check_Auth_RendClass(drv("OCID"), dtArc) Then
                            If sm.UserInfo.RID <> "A" Then
                                Select Case sm.UserInfo.RoleID
                                    Case 0, 1
                                        If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days2 Then
                                            Btn1_add_std.Visible = False
                                            Btn2_add.Visible = False
                                            TIMS.Tooltip(e.Item, "超過結訓日期" & Days2 & "天，停用新增功能..")
                                        End If
                                    Case Else
                                        If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days1 Then
                                            Btn1_add_std.Visible = False
                                            Btn2_add.Visible = False
                                            TIMS.Tooltip(e.Item, "超過結訓日期" & Days1 & "天，停用新增功能.")
                                        End If
                                End Select
                            End If
                        Else
                            If sm.UserInfo.RID <> "A" Then
                                Select Case sm.UserInfo.RoleID
                                    Case 0, 1
                                        '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成75天
                                        '暫時先改這樣,以後還會再改
                                        If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days2 Then
                                            Btn1_add_std.Visible = False
                                            Btn2_add.Visible = False
                                            TIMS.Tooltip(e.Item, "超過結訓日期" & Days2 & "天，停用新增功能..")
                                        End If

                                    Case Else
                                        '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成60天
                                        '暫時先改這樣,以後還會再改
                                        If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days1 Then
                                            Btn1_add_std.Visible = False
                                            Btn2_add.Visible = False
                                            TIMS.Tooltip(e.Item, "超過結訓日期" & Days1 & "天，停用新增功能.")
                                        End If
                                End Select
                            End If
                        End If
                    End If

                End If

                '(列印封面)
                If Convert.ToString(drv("DLID")) <> "" Then
                    Btn3_print.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_printFN1, "DLID=" & drv("DLID").ToString)
                Else
                    Btn3_print.Attributes("onclick") = "alert('封面尚未建立，請按「新增」');return false;"
                End If
                '20090227 andy edit

                '非署(局)屬 【署(局)屬 無新增btn2功能】
                'cst_Juzhu '"1"  'cst_Juzhu 1 署(局)屬 
                If v_RadioButtonList1 = cst_Juzhu Then
                    Dim vsTitle As String = ""
                    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > 100 Then  '結訓日超過一百天不開放
                        Btn1_add_std.Visible = False
                        Btn2_add.Visible = False
                        vsTitle = "超過結訓日期" & "100天，停用新增功能"
                        TIMS.Tooltip(e.Item, vsTitle)
                    End If
                    '授權設定該班級有設定則開放
                    If Not TIMS.ChkIsEndDate(CStr(drv("OCID")), TIMS.cst_FunID_結訓學員資料卡登錄, dtArc) Then
                        Btn1_add_std.Enabled = True
                        vsTitle = "授權設定該班級有開放"
                        TIMS.Tooltip(e.Item, vsTitle)
                    End If
                    'If TIMS.ChkIsEndDate(Convert.ToString(DataGrid1.DataKeys(e.Item.ItemIndex)), "154", dtArc) = False Then
                    'End If
                    'If ChkIsEndDate(sm.UserInfo.UserID, Convert.ToString(DataGrid1.DataKeys(e.Item.ItemIndex))) = False Then   '授權設定該班級有設定則開放
                    '    btn2.Enabled = True
                    'End If
                End If
        End Select
    End Sub

    ''' <summary> stud list Command </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Select Case e.CommandName
            Case "del"
                Dim sCmdArg As String = e.CommandArgument
                Dim v_DLID As String = TIMS.GetMyValue(sCmdArg, "DLID")
                Dim v_SubNo As String = TIMS.GetMyValue(sCmdArg, "SubNo")
                If sCmdArg = "" OrElse v_DLID = "" OrElse v_SubNo = "" Then
                    Common.MessageBox(Me, "資料刪除範圍過大，停止刪除動作!!")
                    Exit Sub
                End If

                Dim sql As String = ""
                sql = "SELECT COUNT(1) CNT FROM STUD_RESULTSTUDDATA WHERE DLID=@DLID AND SubNo=@SubNo" 'DLIDSubNo & e.CommandArgument
                Call TIMS.OpenDbConn(objconn)
                Dim dt As New DataTable
                Dim sCmd As New SqlCommand(sql, objconn)
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("DLID", SqlDbType.BigInt).Value = Val(v_DLID)
                    .Parameters.Add("SubNo", SqlDbType.BigInt).Value = Val(v_SubNo)
                    dt.Load(.ExecuteReader())
                End With
                If dt.Rows.Count = 0 Then Exit Sub
                Dim dr As DataRow = dt.Rows(0)
                If Val(dr("CNT")) > 100 Then
                    Common.MessageBox(Me, "資料刪除範圍過大，停止刪除動作!!")
                    Exit Sub
                End If

                'SELECT * FROM STUD_RESULTSTUDDATA WHERE ROWNUM <=1
                'SELECT * FROM STUD_RESULTIDENTDATA WHERE ROWNUM <=1
                'SELECT * FROM STUD_RESULTTWELVEDATA WHERE ROWNUM <=1
                Dim s_COLUMN1 As String = "DLID,SUBNO,SOCID,STDNAME,STUDENTID,STDPID,SEX,BIRTHYEAR,BIRTHMONTH,BIRTHDATE,DEGREEID,MILITARYID,Q7,Q8,Q8OTHER,Q9,Q9Y,Q10,Q11,Q11N,MODIFYACCT,MODIFYDATE,Q12A,Q12B,Q12V1,Q12V2,Q12V3,Q12V4,Q12V5"
                sql = ""
                sql &= " UPDATE STUD_RESULTSTUDDATA "
                sql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=getdate() WHERE DLID=@DLID AND SubNo=@SubNo"
                'sql &= e.CommandArgument
                Dim uCmd As New SqlCommand(sql, objconn)

                sql = ""
                sql &= String.Concat(" INSERT INTO STUD_RESULTSTUDDATA_BAK (", s_COLUMN1, ") ")
                sql &= String.Concat(" SELECT ", s_COLUMN1, " FROM STUD_RESULTSTUDDATA WHERE DLID=@DLID AND SubNo=@SubNo")
                'sql &= e.CommandArgument
                Dim iCmd As New SqlCommand(sql, objconn)

                Call TIMS.OpenDbConn(objconn)
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("DLID", SqlDbType.BigInt).Value = Val(v_DLID)
                    .Parameters.Add("SubNo", SqlDbType.BigInt).Value = Val(v_SubNo)
                    .ExecuteNonQuery()
                End With

                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("DLID", SqlDbType.BigInt).Value = Val(v_DLID)
                    .Parameters.Add("SubNo", SqlDbType.BigInt).Value = Val(v_SubNo)
                    .ExecuteNonQuery()
                End With

                Dim parms As New Hashtable
                parms.Add("DLID", Val(v_DLID))
                parms.Add("SubNo", Val(v_SubNo))

                sql = " DELETE STUD_RESULTSTUDDATA WHERE DLID=@DLID AND SubNo=@SubNo" ' e.CommandArgument
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
                'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
                'BY AMU 2009-07-30
                '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
                sql = " DELETE STUD_RESULTIDENTDATA WHERE DLID=@DLID AND SubNo=@SubNo" ' e.CommandArgument
                DbAccess.ExecuteNonQuery(sql, objconn, parms)

                sql = " DELETE STUD_RESULTTWELVEDATA WHERE DLID=@DLID AND SubNo=@SubNo" ' e.CommandArgument
                DbAccess.ExecuteNonQuery(sql, objconn, parms)

                'Common.MessageBox(Me, "刪除成功!")
                If Me.OCID.Value <> "" Then
                    ShowStudData(Me.OCID.Value, "")
                End If
                If Me.OCID.Value = "" AndAlso Me.DLID.Value <> "" Then
                    ShowStudData("", Me.DLID.Value)
                End If

            Case Else
                If e.CommandArgument <> "" Then
                    Call KeepSearchStr()
                    TIMS.Utl_Redirect1(Me, e.CommandArgument)
                Else
                    Common.MessageBox(Me, "(操作錯誤)不提供該功能。")
                    Exit Sub
                End If
        End Select
    End Sub

    ''' <summary> stud list </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Dim sJuzhu As String = Get_sJuzhu() '取得署(局)屬傳遞參數。
        Dim v_RadioButtonList1 As String = TIMS.GetListValue(RadioButtonList1)
        Dim s_MRqID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        Dim sUrl As String = ""

        Dim drv As DataRowView = e.Item.DataItem
        Dim mybutEdit As Button = e.Item.FindControl("Button4")   '修改
        Dim mybutDel As Button = e.Item.FindControl("Button9")    '刪除 
        Dim mybutAdd As Button = e.Item.FindControl("Button5")    '新增

        Dim objControl As HtmlInputCheckBox = e.Item.FindControl("Checkbox2")
        Dim chkbox_all As HtmlInputCheckBox = e.Item.FindControl("chkbox_all")
        Dim LabName As Label = e.Item.FindControl("LabName")
        Dim LabStart2 As Label = e.Item.FindControl("LabStart2")

        Select Case e.Item.ItemType
            Case ListItemType.Header
                chkbox_all.Attributes("onclick") = "ChangeAll(this);"
            Case ListItemType.AlternatingItem, ListItemType.Item, ListItemType.EditItem
                Dim v_RejectTDate1 As String = If(drv("RejectTDate1").ToString <> "", String.Concat("(", FormatDateTime(drv("RejectTDate1"), 2), ")"), "")
                Dim v_RejectTDate2 As String = If(drv("RejectTDate2").ToString <> "", String.Concat("(", FormatDateTime(drv("RejectTDate2"), 2), ")"), "")
                LabName.Text = String.Concat(drv("Name"), v_RejectTDate1, v_RejectTDate2) '姓名
                Dim RejFlag As Boolean = If(v_RejectTDate1 <> "", True, If(v_RejectTDate2 <> "", True, False)) '(有離退日期)

                Select Case v_RadioButtonList1 'RadioButtonList1.SelectedValue
                    Case cst_Juzhu
                        'cst_Juzhu 1 署(局)屬
                        'Dim msg As String = ""
                        Dim drQdata As DataRow = SD_05_008_D.Check_QdataDr(drv("OCID"), Convert.ToString(drv("SOCID")), objconn)

                        e.Item.Cells(cst_dg2學號).Text = Right(Trim(drv("StudentID")), 2)
                        'e.Item.FindControl("Checkbox2")
                        objControl.Value = Right(Trim(drv("StudentID")), 2)
                        'objControl.Disabled = True
                        objControl.Visible = True '(學號)'e.Item.FindControl("Checkbox2")

                        '表示此學員尚未填寫資料 '新增功能
                        e.Item.Cells(cst_dg2填寫狀態).Text = If(drQdata Is Nothing, "否", "是")
                        mybutEdit.Visible = (drQdata IsNot Nothing) '有資料修改
                        mybutAdd.Visible = (drQdata Is Nothing) '沒資料新增
                        mybutDel.Visible = False '不提供刪除

                        mybutAdd.Enabled = (sm.UserInfo.RoleID <> 0) '角色啟用
                        mybutEdit.Enabled = (sm.UserInfo.RoleID <> 0) '角色啟用

                        If drQdata Is Nothing Then '新增
                            If drv("DLID").ToString = "" AndAlso drv("DLID2").ToString = "" Then
                                '異常資料 重新新增
                                sUrl = ""
                                sUrl &= String.Concat("SD_05_008_C.aspx?ID=", s_MRqID, sJuzhu)
                                sUrl &= String.Concat("&Proecess=addmy", "&SOCID=", drv("SOCID"), "&OCID=", drv("OCID"))
                            ElseIf drv("DLID").ToString <> "" Then
                                sUrl = ""
                                sUrl &= String.Concat("SD_05_008_D.aspx?ID=", s_MRqID, sJuzhu)
                                sUrl &= String.Concat("&Proecess=add", "&DLID=", drv("DLID"), "&SOCID=", drv("SOCID"), "&OCID=", drv("OCID"))
                            ElseIf drv("DLID2").ToString <> "" Then
                                sUrl = ""
                                sUrl &= String.Concat("SD_05_008_D.aspx?ID=", s_MRqID, sJuzhu)
                                sUrl &= String.Concat("&Proecess=add", "&DLID=", drv("DLID2"), "&SOCID=", drv("SOCID"), "&OCID=", drv("OCID"))
                            End If
                            If sUrl <> "" Then mybutAdd.CommandArgument = sUrl

                            If RejFlag Then '有離退資料(不開放新增功能)
                                mybutAdd.Enabled = False '(不開放新增功能)
                                TIMS.Tooltip(mybutAdd, "離退訓學員不需填寫結訓學員資料卡!!")
                            End If
                        Else
                            '修改／刪除功能
                            Dim fg_NG As Boolean = SD_05_008_D.Check_Qdata(drv("OCID"), Convert.ToString(drv("SOCID")), objconn)
                            If Not fg_NG Then  '檢查是否有未填寫的學員或題目
                                LabStart2.Visible = False '星號不顯示
                            Else
                                LabStart2.Text = "*"   '星號顯示
                                LabStart2.ForeColor = Color.Red
                                LabStart2.Visible = True
                                TIMS.Tooltip(LabStart2, "有未填寫的題目")
                                mybutEdit.Visible = (drQdata IsNot Nothing) '(可修改)
                                mybutEdit.Enabled = (drQdata IsNot Nothing) '(可修改)
                            End If

                            If TIMS.Check_Auth_RendClass(drv("OCID"), dtArc) Then
                                mybutDel.Visible = True '顯示
                            Else
                                Dim blnFun1 As Boolean = TIMS.Get_FunEnable(Convert.ToString(sm.UserInfo.UserID), 154, "Del", objconn)
                                If blnFun1 Then mybutDel.Visible = True '顯示
                            End If
                            Dim sCmdArg As String = ""
                            TIMS.SetMyValue(sCmdArg, "DLID", Convert.ToString(drv("DLID")))
                            TIMS.SetMyValue(sCmdArg, "SubNo", Convert.ToString(drv("SubNo")))
                            mybutDel.CommandArgument = sCmdArg '"WHERE DLID='" & drv("DLID").ToString & "' and SubNo='" & drv("SubNO").ToString & "'"
                            mybutDel.Attributes("onclick") = "return confirm('確定要刪除此學員填寫的資料卡?');"

                            sUrl = ""
                            sUrl &= String.Concat("SD_05_008_D.aspx?ID=", s_MRqID, sJuzhu)
                            sUrl &= String.Concat("&Proecess=edit", "&DLID=", drv("DLID"), "&SubNo=", drv("SubNo"))
                            mybutEdit.CommandArgument = sUrl

                            If Not TIMS.Check_Auth_RendClass(drv("OCID"), dtArc) Then
                                '沒有 已結訓班級使用權限 則執行此段程式？
                                Select Case sm.UserInfo.RoleID
                                    Case 0, 1
                                        If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days2 Then
                                            mybutAdd.Visible = False
                                            mybutEdit.Enabled = False
                                            TIMS.Tooltip(mybutEdit, String.Concat("結訓日期過", Days2, "天停用修改.."), True)
                                            mybutDel.Enabled = False
                                            TIMS.Tooltip(mybutDel, String.Concat("結訓日期過", Days2, "天停用刪除.."), True)
                                        End If
                                    Case Else
                                        If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days1 Then
                                            mybutAdd.Visible = False
                                            mybutEdit.Enabled = False
                                            TIMS.Tooltip(mybutEdit, String.Concat("結訓日期過", Days1, "天停用修改."), True)
                                            mybutDel.Enabled = False
                                            TIMS.Tooltip(mybutDel, String.Concat("結訓日期過", Days1, "天停用刪除."), True)
                                        End If
                                End Select
                            End If

                            If sm.UserInfo.LID = 0 Then
                                '署(局)登入
                                '且提供修改鈕 才開放 刪除功能
                                If mybutEdit.Visible Then
                                    mybutDel.Enabled = True
                                    mybutDel.Visible = True
                                    mybutDel.Attributes("onclick") = "return confirm('確定要刪除此學員填寫的資料卡?');"
                                End If
                            Else
                                If sm.UserInfo.RoleID <= 1 Then
                                    '管理者權限提供清除功能
                                    '且提供修改鈕,有離退 才開放 刪除功能
                                    If RejFlag AndAlso mybutEdit.Visible Then
                                        mybutDel.Enabled = True
                                        mybutDel.Visible = True
                                        mybutDel.Attributes("onclick") = "return confirm('確定要刪除此學員填寫的資料卡?');"
                                    End If
                                End If
                            End If

                        End If

                        Dim vsTitle As String = ""
                        If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > 100 Then       '結訓日超過一百天不開放
                            vsTitle = String.Concat("超過結訓日期", "100天，停用功能")
                            mybutEdit.Enabled = False
                            mybutAdd.Enabled = False
                            mybutDel.Enabled = False
                            TIMS.Tooltip(mybutEdit, vsTitle)
                            TIMS.Tooltip(mybutAdd, vsTitle)
                            TIMS.Tooltip(mybutDel, vsTitle)
                        End If
                        '授權設定該班級有設定則開放
                        If Not TIMS.ChkIsEndDate(Convert.ToString(drv("OCID")), TIMS.cst_FunID_結訓學員資料卡登錄, dtArc) Then
                            mybutEdit.Enabled = True
                            mybutAdd.Enabled = True
                            mybutDel.Enabled = True
                            vsTitle = "授權設定該班級有開放"
                            TIMS.Tooltip(mybutEdit, vsTitle)
                            TIMS.Tooltip(mybutAdd, vsTitle)
                            TIMS.Tooltip(mybutDel, vsTitle)
                        End If

                        'snoopy可強制刪除
                        Dim fg_snoopy As Boolean = If(Convert.ToString(sm.UserInfo.UserID) <> "" AndAlso LCase(Convert.ToString(sm.UserInfo.UserID).Trim) = "snoopy", True, False)
                        If fg_snoopy Then
                            If Not mybutDel.Visible Then mybutDel.Visible = True
                            If Not mybutDel.Enabled Then mybutDel.Enabled = True
                            mybutDel.Attributes("onclick") = "return confirm('確定要刪除此學員填寫的資料卡?');"
                            TIMS.Tooltip(mybutDel, "目前登入使用者 提供刪除功能,可強制刪除!!")
                        End If
                        If TIMS.sUtl_ChkTest() Then '測試用
                            Dim stp_mybutEdit As String = String.Concat("測試環境，測試功能開啟!!", ",Visible: ", mybutEdit.Visible, ",Enabled: ", mybutEdit.Enabled)
                            Dim stp_mybutAdd As String = String.Concat("測試環境，測試功能開啟!!", ",Visible: ", mybutAdd.Visible, ",Enabled: ", mybutAdd.Enabled)
                            mybutEdit.Visible = True '修改
                            mybutAdd.Visible = (drQdata Is Nothing) '新增
                            mybutEdit.Enabled = True '修改
                            mybutAdd.Enabled = (drQdata Is Nothing) '新增
                            TIMS.Tooltip(mybutEdit, stp_mybutEdit)
                            TIMS.Tooltip(mybutAdd, stp_mybutAdd)
                        End If

                    Case Else
                        'cst_NonJuzhu 2 非署(局)屬
                        LabStart2.Visible = False '星號不顯示
                        e.Item.Cells(cst_dg2學號).Text = Convert.ToString(drv("StudentID"))
                        objControl.Value = Convert.ToString(drv("StudentID"))
                        e.Item.Cells(cst_dg2填寫狀態).Text = If(Convert.ToString(drv("SUBNO")) <> "" AndAlso Convert.ToString(drv("CHK_SUBNO2")) <> "", "是", "否")

                        sUrl = ""
                        sUrl &= String.Concat("SD_05_008_D.aspx?ID=", s_MRqID, sJuzhu)
                        sUrl &= String.Concat("&Proecess=edit", "&DLID=", drv("DLID"), "&SubNo=", drv("SUBNO"))
                        mybutEdit.CommandArgument = sUrl

                        '非署(局)屬 'sCmdArg &= " WHERE DLID='" & drv("DLID").ToString & "' AND SubNo='" & drv("SubNO").ToString & "'"
                        Dim sCmdArg As String = ""
                        TIMS.SetMyValue(sCmdArg, "DLID", Convert.ToString(drv("DLID")))
                        TIMS.SetMyValue(sCmdArg, "SubNo", Convert.ToString(drv("SubNo")))
                        mybutDel.CommandArgument = sCmdArg
                        'TIMS.Tooltip(mybutDel, "非署屬 不開放新增功能!!")

                        mybutAdd.Visible = False
                        TIMS.Tooltip(mybutAdd, "非署屬 不開放新增功能!!")

                        If sm.UserInfo.RoleID <= 1 Then
                            mybutDel.Enabled = True
                            mybutDel.Visible = True
                            mybutDel.Attributes("onclick") = "return confirm('確定要刪除此學員填寫的資料卡?');"
                        Else
                            'mybutDel.Visible = False
                            If sm.UserInfo.LID = 0 Then '署(局)登入
                                mybutDel.Enabled = True
                                mybutDel.Visible = True
                                mybutDel.Attributes("onclick") = "return confirm('確定要刪除此學員填寫的資料卡?');"
                            End If
                        End If

                End Select

        End Select

    End Sub

    ''' <summary>'新增班別封面檔</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call KeepSearchStr()
        Dim sJuzhu As String = Get_sJuzhu() '取得署(局)屬傳遞參數。
        Dim s_MRqID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        Dim sUrl As String = String.Concat("SD_05_008_C.aspx?ID=", s_MRqID, sJuzhu, "&Proecess=addother")
        TIMS.Utl_Redirect1(Me, sUrl)
    End Sub

    '保留搜尋值'Session(cst_search) 
    Sub KeepSearchStr()
        Dim s_SEARCH1 As String = ""
        s_SEARCH1 = "RadioButtonList1=" & RadioButtonList1.SelectedValue
        s_SEARCH1 &= "&center=" & center.Text
        s_SEARCH1 &= "&RIDValue=" & RIDValue.Value
        s_SEARCH1 &= "&TPlan=" & TPlan.SelectedValue
        s_SEARCH1 &= "&TMID1=" & TMID1.Text
        s_SEARCH1 &= "&OCID1=" & OCID1.Text
        s_SEARCH1 &= "&TMIDValue1=" & TMIDValue1.Value
        s_SEARCH1 &= "&OCIDValue1=" & OCIDValue1.Value
        s_SEARCH1 &= "&FTDate1=" & FTDate1.Text
        s_SEARCH1 &= "&FTDate2=" & FTDate2.Text
        s_SEARCH1 &= "&UnitCode=" & UnitCode.SelectedValue
        s_SEARCH1 &= "&ClassName=" & ClassName.Text
        s_SEARCH1 &= String.Concat("&StudentTable=", If(StudentTable.Style.Item("display") = "none", "none", "inline"))
        s_SEARCH1 &= "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        s_SEARCH1 &= "&OCID=" & Me.OCID.Value 'Me.ViewState("OCID")
        s_SEARCH1 &= "&DLID=" & Me.DLID.Value 'Me.ViewState("DLID")
        Session(cst_search) = s_SEARCH1
    End Sub

    '回班別列表
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        StudentTable.Style.Item("display") = "none"
        SearchTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
        ClassTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
        'Me.ViewState("OCID") = ""
        'Me.ViewState("DLID") = ""
        OCID.Value = ""
        DLID.Value = ""
    End Sub

    ''' <summary>結訓學員匯出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub But1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But1.Click
        Dim url1 As String = String.Concat("SD_05_008_E.aspx?ID=", TIMS.Get_MRqID(Me))
        TIMS.Utl_Redirect1(Me, url1)
    End Sub

    '該帳號有賦于班級時(只有一個時)帶出該班級
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    ''' <summary>'取得署(局)屬傳遞參數。【1:署(局)屬 cst_Juzhu /2:非署(局)屬 cst_NonJuzhu】</summary>
    ''' <returns></returns>
    Function Get_sJuzhu() As String
        '預設'cst_NonJuzhu 2 非署(局)屬
        Dim rst As String = String.Concat("&Juzhu=", cst_NonJuzhu)
        Dim v_RadioButtonList1 As String = TIMS.GetListValue(RadioButtonList1)
        Select Case v_RadioButtonList1 'RadioButtonList1.SelectedValue
            Case cst_Juzhu 'cst_Juzhu 1 署(局)屬 
                rst = String.Concat("&Juzhu=", cst_Juzhu)
            Case cst_NonJuzhu  'cst_NonJuzhu 2 非署(局)屬
                rst = String.Concat("&Juzhu=", cst_NonJuzhu)
            Case Else
                TIMS.sUtl_404NOTFOUND(Me, objconn)
        End Select
        Return rst
    End Function

    ''' <summary>列印結訓學員資料卡</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If HidStudentID.Value = "" Then
            Common.MessageBox(Me, "請勾選要列印的學員!")
            Exit Sub
        End If

        DLID.Value = TIMS.ClearSQM(DLID.Value)
        OCID.Value = TIMS.ClearSQM(OCID.Value)
        HidStudentID.Value = TIMS.ClearSQM(HidStudentID.Value)

        Dim sMyValue As String = ""
        TIMS.SetMyValue(sMyValue, "DLID", DLID.Value)
        TIMS.SetMyValue(sMyValue, "OCID", OCID.Value)
        TIMS.SetMyValue(sMyValue, "StudentID", HidStudentID.Value)
        Dim sPrtFN1 As String = ""
        Select Case RadioButtonList1.SelectedValue
            Case cst_Juzhu 'cst_Juzhu 1 署(局)屬 
                Dim drCC As DataRow = TIMS.GetOCIDDate(OCID.Value, objconn)
                'Dim blnPrint2016 As Boolean = False
                sPrtFN1 = cst_printFN2 '署(局)屬
                If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuySD20160701), CDate(drCC("STDATE"))) >= 0 Then
                    'blnPrint2016 = True
                    sPrtFN1 = cst_printFN3 '署(局)屬
                End If

            Case Else  'cst_NonJuzhu 2 非署(局)屬
                sPrtFN1 = cst_printFN4 '非署(局)屬
        End Select
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrtFN1, sMyValue)
    End Sub

    ''' <summary>列印學員空白資料卡 【只有 署(局)屬 才列印空白】</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'Button12.Attributes("onclick") = "CheckPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=MultiBlock&filename=ResultStud1&path=" & sPath & "&');return false;"
        If HidStudentID.Value = "" Then
            Common.MessageBox(Me, "請勾選要列印的學員!")
            Exit Sub
        End If

        DLID.Value = TIMS.ClearSQM(DLID.Value)
        OCID.Value = TIMS.ClearSQM(OCID.Value)
        HidStudentID.Value = TIMS.ClearSQM(HidStudentID.Value)

        Dim sMyValue As String = ""
        TIMS.SetMyValue(sMyValue, "DLID", DLID.Value)
        TIMS.SetMyValue(sMyValue, "OCID", OCID.Value)
        TIMS.SetMyValue(sMyValue, "StudentID", HidStudentID.Value)
        Dim sPrtFN1 As String = ""

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCID.Value, objconn)
        'Dim blnPrint2016 As Boolean = False
        sPrtFN1 = cst_printFN5 '署(局)屬
        If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuySD20160701), CDate(drCC("STDATE"))) >= 0 Then
            'blnPrint2016 = True
            sPrtFN1 = cst_printFN6 '署(局)屬
        End If

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrtFN1, sMyValue)
    End Sub

End Class
