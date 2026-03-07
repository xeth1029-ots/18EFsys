Partial Class CP_07_001
    Inherits AuthBasePage

    '自辦受訓期間學員滿意度匯入功能
    '受訓期間學員滿意度 學員滿意度匯入
    '首頁>>學員動態管理>>訓練成效與滿意度>>受訓期間學員滿意度
    'Select top 10 * from STUD_QUESTRAINING --Stud_QuesTraining
    'Select top 10 * from Stud_ForumRecord

    'Dim strRid As String = ""
    'Dim strOCID As String = ""
    'Dim strSOCID As String = ""
    'Dim isloaded As Boolean
    'Dim RelshipTable As DataTable
    Const cst_SearchStr As String = "CP_07_001_SearchStr"

    Const cst_printFN1 As String = "CP_07_001" '調查表 SOCID
    Const cst_printFN_B As String = "CP_07_001_blank" '空白調查表
    Const cst_printFN_B2 As String = "CP_07_001_blank2" '空白調查表 OCID

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    Const cst_aFILLFORMDATE As Integer = 0 'i填表日期 As Integer = 0
    Const cst_aSIGNER As Integer = 1 '抽訪人員姓名 As Integer = 1 字串長度10
    Const cst_aSTUDID2 As Integer = 2 '學號 字串長度3 或數字
    Const cst_aQ1_1 As Integer = 3 '答案1 As Integer = 3
    Const cst_aQ1_2 As Integer = 4 '答案2 As Integer = 4
    Const cst_aQ1_3 As Integer = 5 '答案3 As Integer = 5

    Const cst_aQ2_1 As Integer = 6 '答案4 As Integer = 6
    Const cst_aQ2_2 As Integer = 7 '答案5 As Integer = 7
    Const cst_aQ2_3 As Integer = 8 '答案6 As Integer = 8
    Const cst_aQ2_4 As Integer = 9 '答案7 As Integer = 9

    Const cst_aQ3_1 As Integer = 10 '答案8 As Integer = 10
    Const cst_aQ3_2 As Integer = 11 '答案9 As Integer = 11
    Const cst_aQ3_3 As Integer = 12 '答案10 As Integer = 12     

    Const cst_aQ4_1 As Integer = 13 '答案11 As Integer = 13 第四部份-1
    Const cst_aQ4_2 As Integer = 14 '答案12 As Integer = 14 第四部份-2
    Const cst_aQ4_3 As Integer = 15 '答案13 As Integer = 15 第四部份-3
    Const cst_aQ4_4 As Integer = 16 '答案14 As Integer = 16 第四部份-4
    Const cst_aSUGGESTION As Integer = 17 '答案15 As Integer = 17 其它意見
    Const cst_iMaxLength1 As Integer = 18 '欄位對應數

    '    OCID numeric   允許     
    'SOCID numeric        
    'RID varchar 10 允許     
    'FILLFORMDATE datetime
    'TYPE numeric   允許     
    'Q1_1 numeric   允許     
    'Q1_2 numeric   允許     
    'Q1_3 numeric   允許     
    'Q2_1 numeric   允許     
    'Q2_2 numeric   允許     
    'Q2_3 numeric   允許     
    'Q3_1 numeric   允許     
    'Q3_2 numeric   允許     
    'Q3_3 numeric   允許     
    'Q4_1 numeric   允許     
    'Q4_2 numeric   允許     
    'Q4_3 numeric   允許     
    'Q4_4 numeric   允許     
    'SUGGESTION nvarchar -1 允許     
    'MODIFYACCT nvarchar 15      
    'MODIFYDATE datetime        
    'SIGNER nvarchar 30 允許     
    'DASOURCE numeric   允許     
    'Q2_4 numeric   允許     


    'SQControl
    'Report:  CP_07_003
    'Report:  CP_07_003_blank

    'Report:  CP_07_001_2010 'old
    'Report:  CP_07_001_blank_2010 'old
    'Report:  CP_07_001 'new
    'Report:  CP_07_001_blank 'new

    'onclick  js@chkOrg, PrintRpt
    'jsValue  = "return PrintRpt(" & Convert.ToString(drv("Type")) & "," & Convert.ToString(drv("socid")) & "," & Convert.ToString(drv("OCID")) & ");"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在-------------------------- Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DG_ClassInfo

        'Dim flag_test_ENVC As Boolean = TIMS.CHK_IS_TEST_ENVC() '檢測為測試環境:true 正式環境為:false
        'trbtnImport1.Visible = flag_test_ENVC '(測試才顯示，正式機暫時不顯示) 20211202 by AMU
        'tr_StudentID.Visible = False
        'Common.RespWrite(Me, "<script>alert('產業人才投資計畫無法使用此功能');top.location.href='../../index.htm';</script>")
        'Response.End()
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Common.MessageBox(Me, "產業人才投資計畫無法使用此功能!!")
            Exit Sub
        End If

        HyperLink1.NavigateUrl = "../../Doc/QuesTrain_v31.zip"

        If Not IsPostBack Then
            cCreate1()

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", , "bt_search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub cCreate1()
        show_table_check(0)
        'PageControler1.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        'strRid = Request("rid")
        'strOCID = Request("ocid")
        'strSOCID = Request("socid")
        'distid.Value = sm.UserInfo.DistID
        'isloaded = False
        bt_search.Attributes.Add("onclick", "return chkOrg();")

        'Me.ViewState("InquireType") = InquireType.SelectedValue.ToString()
        If Session(cst_SearchStr) IsNot Nothing Then
            Dim MyValue As String = ""
            Dim SearchStr As String = Session(cst_SearchStr)
            Session(cst_SearchStr) = Nothing

            MyValue = TIMS.GetMyValue(SearchStr, "pg")
            If MyValue <> "cp07001_sch1" Then Return

            center.Text = TIMS.GetMyValue(SearchStr, "center")
            RIDValue.Value = TIMS.GetMyValue(SearchStr, "RIDValue")
            TMID1.Text = TIMS.GetMyValue(SearchStr, "TMID1")
            OCID1.Text = TIMS.GetMyValue(SearchStr, "OCID1")
            TMIDValue1.Value = TIMS.GetMyValue(SearchStr, "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(SearchStr, "OCIDValue1")
            STDate1.Text = TIMS.GetMyValue(SearchStr, "STDate1")
            STDate2.Text = TIMS.GetMyValue(SearchStr, "STDate2") 'end_date
            Hid_show_table.Value = TIMS.GetMyValue(SearchStr, "Hid_show_table")
            'MyValue = TIMS.GetMyValue(SearchStr, "InquireType")
            'If MyValue <> "" Then Common.SetListItem(InquireType, MyValue)
            MyValue = TIMS.GetMyValue(SearchStr, "PageIndex")
            If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                MyValue = CInt(MyValue)
                PageControler1.PageIndex = MyValue
            End If

            If OCIDValue1.Value <> "" Then
                Call Search2(OCIDValue1.Value)
            Else
                Call Search1()
            End If

        End If

    End Sub

    ''' <summary>
    ''' 狀態改變顯示或隱藏
    ''' </summary>
    ''' <param name="iType"></param>
    Sub show_table_check(ByRef iType As Integer)
        '0 重新載入
        '1 班級查詢
        '2 班級學員查詢
        '3 班級學員查詢 回上頁
        'tr_bt_search.Visible = True
        msg1.Text = ""
        tb_DG_ClassInfo.Visible = False
        msg2.Text = ""
        ClassLabel1.Text = ""
        ClassLabel2.Text = ""
        tb_StudentTable.Visible = False

        Hid_show_table.Value = iType

        If iType = 0 Then Return
        If iType = 1 Then Return

        If iType = 2 Then
            Frametable3.Visible = False
            Return
        End If
        If iType = 3 Then
            Frametable3.Visible = True
            tb_DG_ClassInfo.Visible = True
            Return
        End If
    End Sub

    '查詢SQL
    Private Sub Search1()
        show_table_check(1)

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DG_ClassInfo)

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Common.MessageBox(Me, "產業人才投資計畫無法使用此功能!!")
            Exit Sub
        End If

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)

        Dim parms As New Hashtable
        parms.Clear()
        If RIDValue.Value <> "" Then parms.Add("RID", RIDValue.Value) '&= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
        If OCIDValue1.Value <> "" Then parms.Add("OCID", OCIDValue1.Value) 'sql &= " and cc.OCID='" & OCIDValue1.Value & "' " & vbCrLf
        If sm.UserInfo.LID = "0" Then
            parms.Add("TPlanID", sm.UserInfo.TPlanID)
            parms.Add("Years", sm.UserInfo.Years)
            'sql &= " and cc.TPlanID='" & sm.UserInfo.TPlanID & "' " & vbCrLf
            'sql &= " and cc.Years='" & sm.UserInfo.Years & "' " & vbCrLf
        Else
            parms.Add("PlanID", sm.UserInfo.PlanID)
            'sql &= " and cc.PlanID='" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        If STDate1.Text <> "" Then parms.Add("STDate1", TIMS.Cdate2(STDate1.Text)) 'sql &= " and cc.STDate >= " & TIMS.to_date(STDate1.Text)
        If STDate2.Text <> "" Then parms.Add("STDate2", TIMS.Cdate2(STDate2.Text)) 'sql &= " and cc.STDate <= " & TIMS.to_date(STDate2.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.ORGNAME" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,cc.TPLANID" & vbCrLf
        sql &= " ,cc.YEARS" & vbCrLf
        sql &= " ,cc.PLANID" & vbCrLf
        sql &= " ,cc.RID" & vbCrLf
        sql &= " ,cc.DISTID" & vbCrLf
        sql &= " ,FORMAT(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,FORMAT(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) CLASSCNAME" & vbCrLf
        sql &= " ,cc.TNUM" & vbCrLf
        sql &= " ,dbo.FN_GET_STDCNT(cc.OCID,1) CNT1/*1:開訓人數*/" & vbCrLf
        sql &= " ,dbo.FN_GET_STDCNT(cc.OCID,5) CNT5/*5:結訓人數*/" & vbCrLf
        sql &= " ,dbo.FN_GET_STDCNT(cc.OCID,6) CNT6/*6:問卷填寫人數*/" & vbCrLf

        sql &= " FROM dbo.VIEW2 cc" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and cc.TPlanID!='28'" & vbCrLf
        'sql &= " and cc.TPlanID='06'" & vbCrLf 'sql &= " and cc.YEARS='2021'" & vbCrLf 'sql &= " and cc.DISTID='001'" & vbCrLf
        If RIDValue.Value <> "" Then sql &= " and cc.RID =@RID" & vbCrLf '" & RIDValue.Value & "'" & vbCrLf
        If OCIDValue1.Value <> "" Then sql &= " and cc.OCID=@OCID" & vbCrLf '='" & OCIDValue1.Value & "' " & vbCrLf
        If sm.UserInfo.LID = "0" Then
            sql &= " and cc.TPlanID=@TPlanID" & vbCrLf '" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql &= " and cc.Years=@Years" & vbCrLf '" & sm.UserInfo.Years & "' " & vbCrLf
        Else
            sql &= " and cc.PlanID=@PlanID" & vbCrLf '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        If STDate1.Text <> "" Then sql &= " and cc.STDate >= @STDate1" & vbCrLf '" & TIMS.to_date(STDate1.Text)
        If STDate2.Text <> "" Then sql &= " and cc.STDate <= @STDate2" & vbCrLf '" & TIMS.to_date(STDate2.Text)
        sql &= " ORDER BY cc.STDate,cc.OCID"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        'PageControler1.Visible = False
        msg1.Text = TIMS.cst_NODATAMsg1
        tb_DG_ClassInfo.Visible = False
        PageControler1.Visible = False
        DG_ClassInfo.Visible = False
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return 'Exit Sub
        End If

        msg1.Text = ""
        tb_DG_ClassInfo.Visible = True
        PageControler1.Visible = True
        DG_ClassInfo.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    '查詢學員LIST[SQL]
    Sub Search2(ByRef v_OCIDVal As String)
        show_table_check(2)

        msg2.Text = "查無學生資料!"
        ClassLabel1.Text = ""
        ClassLabel2.Text = ""
        tb_StudentTable.Visible = False

        v_OCIDVal = TIMS.ClearSQM(v_OCIDVal)
        Dim drCC As DataRow = TIMS.GetOCIDDate(v_OCIDVal, objconn)
        If drCC Is Nothing Then Exit Sub
        ClassLabel1.Text = "班別：" & Convert.ToString(drCC("ClassCName2"))

        Dim parms_stdcnt As New Hashtable
        parms_stdcnt.Clear()
        parms_stdcnt.Add("OCID", v_OCIDVal)
        Dim sql_stdcnt As String = ""
        sql_stdcnt = " SELECT opencount,TrainCount,LeaveCount FROM dbo.V_STUDENTCOUNT WHERE OCID =@OCID"
        Dim dr_SC1 As DataRow = DbAccess.GetOneRow(sql_stdcnt, objconn, parms_stdcnt)
        If dr_SC1 Is Nothing Then Exit Sub
        ClassLabel2.Text = String.Format("(開訓人數:{0}&nbsp;&nbsp;在結訓人數:{1}&nbsp;&nbsp;離退訓人數:{2})", dr_SC1("opencount"), dr_SC1("TrainCount"), dr_SC1("LeaveCount"))

        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("OCID", v_OCIDVal)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cs.STUDENTID" & vbCrLf
        sql &= " ,cs.STUDID2" & vbCrLf
        sql &= " ,cs.NAME +case when cs.REJECTTDATE1 is not null then concat('(',format(cs.REJECTTDATE1,'yyyy/MM/dd'),')')" & vbCrLf
        sql &= "  when cs.REJECTTDATE2 is not null then  concat('(',format(cs.REJECTTDATE2,'yyyy/MM/dd'),')')  else '' end STDNAME" & vbCrLf
        sql &= " ,cs.PLANNAME" & vbCrLf
        sql &= " ,cs.DISTNAME" & vbCrLf
        sql &= " ,cs.CLASSCNAME2" & vbCrLf
        sql &= " ,cs.OCID" & vbCrLf
        sql &= " ,cs.SOCID" & vbCrLf
        sql &= " ,cs.RID" & vbCrLf
        sql &= " ,cs.PlanID" & vbCrLf
        sql &= " ,cs.STUDSTATUS" & vbCrLf
        sql &= " ,format(cs.REJECTTDATE1,'yyyy/MM/dd') REJECTTDATE1" & vbCrLf
        sql &= " ,format(cs.REJECTTDATE2,'yyyy/MM/dd') REJECTTDATE2" & vbCrLf
        sql &= " ,CASE WHEN d.SOCID IS NULL THEN '否' ELSE '是' END FILLSTATUS_N /*填寫狀態*/" & vbCrLf
        '問卷填寫狀態
        sql &= " ,d.SOCID QUESOCID" & vbCrLf
        '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
        sql &= " ,d.DASOURCE" & vbCrLf
        sql &= " FROM dbo.V_STUDENTINFO cs" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_QUESTRAINING d ON d.SOCID = cs.SOCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cs.OCID =@OCID " & vbCrLf
        'sql &= " AND cs.OCID =132764" & vbCrLf
        sql &= " ORDER BY cs.STUDID2" & vbCrLf

        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt2.Rows.Count = 0 Then Return

        msg2.Text = ""
        tb_StudentTable.Visible = True

        'Session("DTable_Stuednt") = dt
        'DG_stud.DataKeyField = "SOCID"
        DG_stud.DataSource = dt2
        DG_stud.DataBind()
    End Sub

    Function CHK_QUESTRAINING(ByRef OCID As String, ByRef SOCID As String, ByRef tConn As SqlConnection) As Boolean
        Dim rst As Boolean = False
        Dim sql As String = ""
        sql = "SELECT 1 FROM STUD_QUESTRAINING WHERE SOCID=@SOCID and OCID=@OCID"
        Dim sCmd As New SqlCommand(sql, tConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.Int).Value = Val(SOCID)
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCID)
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function
    Sub KeepSearch()
        Dim SearchStr As String = ""
        SearchStr = ""
        'MyValue = TIMS.GetMyValue(SearchStr, "pg")
        'If MyValue <> "cp07001_sch1" Then Return

        SearchStr &= "&pg=cp07001_sch1" ' & center.Text
        SearchStr &= "&center=" & TIMS.ClearSQM(center.Text)
        SearchStr &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        SearchStr &= "&TMID1=" & TIMS.ClearSQM(TMID1.Text)
        SearchStr &= "&OCID1=" & TIMS.ClearSQM(OCID1.Text)
        SearchStr &= "&TMIDValue1=" & TIMS.ClearSQM(TMIDValue1.Value)
        SearchStr &= "&OCIDValue1=" & TIMS.ClearSQM(OCIDValue1.Value)
        SearchStr &= "&STDate1=" & TIMS.ClearSQM(STDate1.Text)
        SearchStr &= "&STDate2=" & TIMS.ClearSQM(STDate2.Text)
        SearchStr &= "&Hid_show_table=" & TIMS.ClearSQM(Hid_show_table.Value)
        'SearchStr &= "&InquireType=" & InquireType.SelectedValue
        SearchStr &= "&PageIndex=" & DG_ClassInfo.CurrentPageIndex + 1

        Session(cst_SearchStr) = SearchStr
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    ''列印空白報表(系統登打)
    'Protected Sub bt_blankRpt_Click(sender As Object, e As EventArgs) Handles bt_blankRpt.Click
    '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        Common.MessageBox(Me, "產業人才投資計畫無法使用此功能!!")
    '        Exit Sub
    '    End If

    '    Call sUtl_printBlank(1)
    'End Sub

    ''列印空白報表(電話訪查)
    'Protected Sub bt_blankRpt2_Click(sender As Object, e As EventArgs) Handles bt_blankRpt2.Click
    '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        Common.MessageBox(Me, "產業人才投資計畫無法使用此功能!!")
    '        Exit Sub
    '    End If

    '    Call sUtl_printBlank(2)
    'End Sub

    '列印空白報表 1:(系統登打) 2:(電話訪查)  
    Sub sUtl_printBlank(ByVal rptType As Integer)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Common.MessageBox(Me, "產業人才投資計畫無法使用此功能!!")
            Exit Sub
        End If

        'rptType: 1:系統登打 2:電話
        Dim sfileName As String = "CP_07_003_blank" '預設2:電話
        Dim MyValue As String = ""
        MyValue &= "&Years=" & sm.UserInfo.Years
        MyValue &= "&DistID=" & sm.UserInfo.DistID
        Select Case rptType
            Case 1
                sfileName = "CP_07_001_blank" '1:系統登打
            Case Else '2
                sfileName = "CP_07_003_blank" '預設 2:電話
        End Select
        If sfileName = "" Then
            Common.MessageBox(Me, "未選擇列印方式~")
            Exit Sub
        End If

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sfileName, MyValue)

        '&path=' + SMpath + '&Years=' + document.getElementById('years').value + '&distid=' + document.getElementById('distid').value);
    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DG_ClassInfo.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnQUIRE As Button = e.Item.FindControl("btnQUIRE") '查詢
                Dim btnPrintB1 As Button = e.Item.FindControl("btnPrintB1")  '列印空白
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                'sql &= " ,dbo.FN_GET_STDCNT(cc.OCID,5) CNT5/*5:結訓人數*/" & vbCrLf

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "DistID", Convert.ToString(drv("DistID")))

                btnQUIRE.CommandArgument = sCmdArg 'DataGrid1.DataKeys(e.Item.ItemIndex)
                btnPrintB1.CommandArgument = sCmdArg

                Dim flag_no_enter As Boolean = If(Convert.ToString(drv("CNT1")) = "0", True, False)
                btnQUIRE.Enabled = If(flag_no_enter, False, True)
                btnPrintB1.Enabled = If(flag_no_enter, False, True)
                If flag_no_enter Then
                    Dim s_t1 As String = "開訓人數為0,暫不開放"
                    TIMS.Tooltip(btnQUIRE, s_t1, True)
                    TIMS.Tooltip(btnPrintB1, s_t1, True)
                End If

        End Select

    End Sub

    Private Sub DG_ClassInfo_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_ClassInfo.ItemCommand
        Dim sCmdArg As String = ""
        sCmdArg = e.CommandArgument
        If sCmdArg = "" Then Return
        Select Case e.CommandName
            Case "QUIRE"
                Dim v_OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
                If v_OCID = "" Then Return
                KeepSearch()
                Call Search2(v_OCID)

            Case "PrintB1"
                Dim v_OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
                Dim v_PlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
                Dim v_DistID As String = TIMS.GetMyValue(sCmdArg, "DistID")

                Dim myValue1 As String = ""
                TIMS.SetMyValue(myValue1, "OCID", v_OCID)
                TIMS.SetMyValue(myValue1, "TPlanID", sm.UserInfo.TPlanID)
                TIMS.SetMyValue(myValue1, "Years", sm.UserInfo.Years)
                TIMS.SetMyValue(myValue1, "PlanID", v_PlanID)
                TIMS.SetMyValue(myValue1, "DistID", v_DistID)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN_B2, myValue1) '空白調查表 OCID

        End Select
    End Sub

    Protected Sub bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        Call Search1()
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        show_table_check(3)
        Search1()
    End Sub

    Private Sub DG_stud_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'btnInsert'btnEdit'btnCheck'btnPrint'btnClear
                Dim btnInsert As Button = e.Item.FindControl("btnInsert") '新增
                Dim btnEdit As Button = e.Item.FindControl("btnEdit")    '修改
                Dim btnCheck As Button = e.Item.FindControl("btnCheck") '查看
                Dim btnPrint As Button = e.Item.FindControl("btnPrint")  '列印
                Dim btnClear As Button = e.Item.FindControl("btnClear") '清除重填

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "rid", Convert.ToString(drv("rid")))
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "socid", Convert.ToString(drv("socid")))
                TIMS.SetMyValue(sCmdArg, "ocid", Convert.ToString(drv("ocid")))
                '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                TIMS.SetMyValue(sCmdArg, "DaSource", Convert.ToString(drv("DaSource")))

                btnInsert.CommandArgument = sCmdArg
                btnEdit.CommandArgument = sCmdArg
                btnCheck.CommandArgument = sCmdArg
                btnPrint.CommandArgument = sCmdArg
                btnClear.CommandArgument = sCmdArg

                'Dim sName As String = Convert.ToString(drv("name"))
                'If Convert.ToString(drv("RejectTDate1")) <> "" Then sName &= "(" & Convert.ToString(drv("RejectTDate1")) & ")"
                'If Convert.ToString(drv("RejectTDate2")) <> "" Then sName &= "(" & Convert.ToString(drv("RejectTDate2")) & ")"
                'e.Item.Cells(1).Text = sName
                Dim flag_have_data As Boolean = False
                If Convert.ToString(drv("QUESOCID")) <> "" Then
                    '已有資料
                    flag_have_data = True
                    'e.Item.Cells(2).Text = "是"
                    btnInsert.Enabled = False '不可新增
                    TIMS.Tooltip(btnInsert, "已有資料不可新增")
                    btnEdit.Enabled = True '可以修改
                    btnCheck.Enabled = True '可以查看
                    btnClear.Enabled = True '可清除重填 
                    '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                    '學員填寫不可列印。
                    btnPrint.Enabled = If(Convert.ToString(drv("DaSource")) <> "1", True, False) '可以列印
                    Dim s_t2 As String = If(Convert.ToString(drv("DaSource")) <> "1", "非學員自行填寫，可列印", "學員自行填寫，不可列印") '可以列印
                    TIMS.Tooltip(btnPrint, s_t2)
                Else
                    '無資料
                    'e.Item.Cells(2).Text = "否"
                    btnInsert.Enabled = True '可新增
                    btnEdit.Enabled = False '不可以修改
                    btnCheck.Enabled = False '不可以查看
                    btnPrint.Enabled = False '不可以列印
                    btnClear.Enabled = False '不可清除重填
                    TIMS.Tooltip(btnEdit, "沒有資料，不可修改")
                    TIMS.Tooltip(btnCheck, "沒有資料，不可查看")
                    TIMS.Tooltip(btnPrint, "沒有資料，不可以列印")
                    TIMS.Tooltip(btnClear, "沒有資料，不可清除重填")
                End If

                '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                '新增「填寫來源」檢核機制，由系統判斷填寫來源為產投報名網或TIMS系統：
                '1.若為報名網，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰，無法逕行修改，僅保留填寫狀態供訓練單位查詢。
                '2.若為TIMS系統，保留訓練單位各項功能(新增、修改、查詢、列印、清除重填)。
                '3.若有訓練單位先於TIMS系統協助填寫，學員後於報名網修正之情形者，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰。
                'PS: 訓練單位登入,若是學員於外網填寫,是不可修改,不可查詢,不可列印,不可清除重填   但若是分署登入,若是學員於外網填寫,是不可修改,可查詢,可列印,不可清除重填
                If Convert.ToString(drv("DaSource")) = "1" Then
                    If sm.UserInfo.LID = 2 Then
                        '委訓單位
                        btnInsert.Enabled = False '不可新增
                        btnEdit.Enabled = False '不可修改
                        btnCheck.Enabled = False '不可查看
                        btnClear.Enabled = False '不可清除重填
                        btnPrint.Enabled = False '不可以列印
                        TIMS.Tooltip(btnInsert, "學員外網填寫，不可新增")
                        TIMS.Tooltip(btnEdit, "學員外網填寫，不可修改")
                        TIMS.Tooltip(btnCheck, "學員外網填寫，不可查看")
                        TIMS.Tooltip(btnClear, "學員外網填寫，不可清除重填")
                        TIMS.Tooltip(btnPrint, "學員外網填寫，不可以列印")
                        btnInsert.CommandArgument = ""
                        btnEdit.CommandArgument = ""
                        btnCheck.Enabled = True '不可查看
                        btnCheck.CommandArgument = ""
                        btnCheck.Attributes("onclick") = "alert('學員外網填寫，不可查看');return false;"
                        btnClear.CommandArgument = ""
                        btnPrint.CommandArgument = ""
                    Else
                        '非委訓單位(署、分署)
                        btnInsert.Enabled = False '不可新增
                        btnEdit.Enabled = False '不可修改
                        btnCheck.Enabled = True '可查看
                        btnClear.Enabled = False '不可清除重填
                        btnPrint.Enabled = True '可列印
                        TIMS.Tooltip(btnInsert, "學員外網填寫，不可新增")
                        TIMS.Tooltip(btnEdit, "學員外網填寫，不可修改")
                        TIMS.Tooltip(btnCheck, "非委訓單位登入，可查看")
                        TIMS.Tooltip(btnClear, "學員外網填寫，不可清除重填")
                        TIMS.Tooltip(btnPrint, "非委訓單位登入，可列印", True)
                        btnInsert.CommandArgument = "" '不可新增
                        btnEdit.CommandArgument = "" '不可修改
                        'but5.Enabled = True '可查看
                        'but5.CommandArgument = ""
                        'but5.Attributes("onclick") = "alert('學員外網填寫，不可查看');return false;"
                        btnClear.CommandArgument = "" '不可清除重填
                    End If
                End If
        End Select
    End Sub

    Private Sub DG_stud_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DG_stud.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Return 'Exit Sub
        Dim rSOCID As String = TIMS.GetMyValue(sCmdArg, "socid")
        Dim rOCID As String = TIMS.GetMyValue(sCmdArg, "ocid")
        Dim rRID As String = TIMS.GetMyValue(sCmdArg, "rid")
        Dim rPlanID As String = TIMS.GetMyValue(sCmdArg, "planid")
        Dim rDaSource As String = TIMS.GetMyValue(sCmdArg, "DaSource")
        If rSOCID = "" Then Return 'Exit Sub
        If rOCID = "" Then Return 'Exit Sub

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        'Request("ID")
        Dim myValue1 As String = ""
        TIMS.SetMyValue(myValue1, "rid", rRID)
        TIMS.SetMyValue(myValue1, "ocid", rOCID)
        TIMS.SetMyValue(myValue1, "socid", rSOCID)
        TIMS.SetMyValue(myValue1, "PlanID", rPlanID)

        'rid,ocid,socid,PlanID/status
        'status : add/edit/check/clear （新增/修改/查看/清除重填）
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Dim myUrl1 As String = "CP_07_001_add.aspx?"
        Select Case e.CommandName
            Case "Insert" '新增
                Call KeepSearch()
                TIMS.SetMyValue(myValue1, "status", "add")
                TIMS.Utl_Redirect1(Me, myUrl1 & myValue1, objconn)
            Case "Edit" '修改
                Call KeepSearch()
                TIMS.SetMyValue(myValue1, "status", "edit")
                TIMS.Utl_Redirect1(Me, myUrl1 & myValue1, objconn)
            Case "Check" '查詢/可查看
                Call KeepSearch()
                TIMS.SetMyValue(myValue1, "status", "check")
                TIMS.Utl_Redirect1(Me, myUrl1 & myValue1, objconn)
            Case "Clear" '清除重填。
                Call KeepSearch()
                TIMS.CloseDbConn(objconn)
                TIMS.SetMyValue(myValue1, "status", "clear")
                'TIMS.Utl_Redirect1(Me, myUrl1 & myValue1)

                Dim sConfirm1 As String = String.Format("此動作會刪除 {0}，是否確定刪除?", "學員受訓期間意見調查表")
                Dim strScript As String = ""
                strScript = "<script language=""javascript"">" + vbCrLf
                strScript += "if (window.confirm('" & sConfirm1 & "')){" + vbCrLf
                strScript += "location.href ='" & myUrl1 & myValue1 & "';}" + vbCrLf
                strScript += "</script>"
                Page.RegisterStartupScript(TIMS.xBlockName, strScript)

            Case "Print" '列印cst_ptPrint
                Call KeepSearch()
                TIMS.CloseDbConn(objconn)
                'TIMS.SetMyValue(myValue1, "status", "edit")

                Dim prtValue1 As String = ""
                'TIMS.SetMyValue(prtValue1, "RID", rRID)
                TIMS.SetMyValue(prtValue1, "OCID", rOCID)
                TIMS.SetMyValue(prtValue1, "SOCID", rSOCID)
                TIMS.SetMyValue(prtValue1, "PlanID", rPlanID)
                'Const cst_printFN1 As String = "CP_07_001" '調查表 SOCID
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, prtValue1)

        End Select
    End Sub

    Protected Sub btnPrintB2_Click(sender As Object, e As EventArgs) Handles btnPrintB2.Click
        'Select Case concat(dbo.FN_GET_ROC_YEAR(ip.YEARS),'年 度 學員受訓期間意見調查表') TITLE1
        Dim myValue1 As String = ""
        TIMS.SetMyValue(myValue1, "TPlanID", sm.UserInfo.TPlanID)
        TIMS.SetMyValue(myValue1, "Years", sm.UserInfo.Years)
        TIMS.SetMyValue(myValue1, "PlanID", sm.UserInfo.PlanID)
        TIMS.SetMyValue(myValue1, "DistID", sm.UserInfo.DistID)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN_B, myValue1) '空白調查表
    End Sub

    ''' <summary>
    ''' '取得一個iCmd
    ''' </summary>
    ''' <param name="tConn"></param>
    ''' <returns></returns>
    Public Shared Function GetxICmd(ByRef tConn As SqlConnection) As SqlCommand
        'Call TIMS.OpenDbConn(tConn)
        Dim i_sql As String = ""
        i_sql = ""
        i_sql &= " INSERT INTO STUD_QUESTRAINING (" & vbCrLf
        i_sql &= " OCID ,SOCID ,RID" & vbCrLf
        i_sql &= " ,FILLFORMDATE" & vbCrLf
        i_sql &= " ,TYPE" & vbCrLf
        i_sql &= " ,Q1_1" & vbCrLf
        i_sql &= " ,Q1_2" & vbCrLf
        i_sql &= " ,Q1_3" & vbCrLf
        i_sql &= " ,Q2_1" & vbCrLf
        i_sql &= " ,Q2_2" & vbCrLf
        i_sql &= " ,Q2_3" & vbCrLf
        i_sql &= " ,Q2_4" & vbCrLf
        i_sql &= " ,Q3_1" & vbCrLf
        i_sql &= " ,Q3_2" & vbCrLf
        i_sql &= " ,Q3_3" & vbCrLf
        i_sql &= " ,Q4_1" & vbCrLf
        i_sql &= " ,Q4_2" & vbCrLf
        i_sql &= " ,Q4_3" & vbCrLf
        i_sql &= " ,Q4_4" & vbCrLf
        i_sql &= " ,SUGGESTION" & vbCrLf 'nvarchar
        i_sql &= " ,SIGNER" & vbCrLf 'nvarchar
        '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統後台()
        i_sql &= " ,DASOURCE" & vbCrLf 'DASOURCE numeric 
        i_sql &= " ,MODIFYACCT" & vbCrLf
        i_sql &= " ,MODIFYDATE" & vbCrLf
        i_sql &= " ) VALUES (" & vbCrLf
        i_sql &= " @OCID ,@SOCID ,@RID" & vbCrLf
        i_sql &= " ,@FILLFORMDATE " & vbCrLf 'GETDATE() " & vbCrLf 'FILLFORMDATE
        i_sql &= " ,@TYPE" & vbCrLf
        i_sql &= " ,@Q1_1" & vbCrLf
        i_sql &= " ,@Q1_2" & vbCrLf
        i_sql &= " ,@Q1_3" & vbCrLf
        i_sql &= " ,@Q2_1" & vbCrLf
        i_sql &= " ,@Q2_2" & vbCrLf
        i_sql &= " ,@Q2_3" & vbCrLf
        i_sql &= " ,@Q2_4" & vbCrLf
        i_sql &= " ,@Q3_1" & vbCrLf
        i_sql &= " ,@Q3_2" & vbCrLf
        i_sql &= " ,@Q3_3" & vbCrLf
        i_sql &= " ,@Q4_1" & vbCrLf
        i_sql &= " ,@Q4_2" & vbCrLf
        i_sql &= " ,@Q4_3" & vbCrLf
        i_sql &= " ,@Q4_4" & vbCrLf
        i_sql &= " ,@SUGGESTION" & vbCrLf
        i_sql &= " ,@SIGNER" & vbCrLf
        '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統後台()
        i_sql &= " ,@DASOURCE" & vbCrLf
        i_sql &= " ,@MODIFYACCT" & vbCrLf
        i_sql &= " ,GETDATE()" & vbCrLf
        i_sql &= " )" & vbCrLf
        Dim iCMD As New SqlCommand(i_sql, tConn)
        Return iCMD
    End Function

    ''' <summary>
    '''  OJT-21100501：自辦在職、接受企業委託 - 受訓期間學員滿意度：新增批次匯入功能
    ''' </summary>
    ''' <param name="obj1"></param>
    ''' <returns></returns>
    Function CHG_VALUE1(ByVal obj1 As Object) As Object
        If obj1 Is Nothing Then Return obj1
        If Convert.ToString(obj1) = "" Then Return obj1
        '匯入的Excel檔填的資料是：
        '非常滿意：5 '滿意：4 '普通：3 '不滿意：2 '非常不滿意：1 
        '(資料庫的值 1、2、3、4、5)
        Dim iVal As Integer = TIMS.GetValue2(obj1)
        Return If(iVal = 5, 1, If(iVal = 4, 2, If(iVal = 3, 3, If(iVal = 2, 4, If(iVal = 1, 5, obj1)))))
    End Function

    ''' <summary>
    ''' 轉換匯入資料變成 DataRow
    ''' </summary>
    ''' <param name="colArray"></param>
    ''' <returns></returns>
    Function ChgImpData(ByRef colArray As Array, ByRef xHash As Hashtable) As DataRow
        Dim vSOCID As String = TIMS.GetMyValue2(xHash, "SOCID")
        'Dim vRID As String = TIMS.GetMyValue2(xHash, "RID")
        Dim dr1 As DataRow = Nothing
        Dim sql As String = ""
        sql = " SELECT * FROM STUD_QUESTRAINING WHERE 1<>1 "
        Dim dtV As DataTable = DbAccess.GetDataTable(sql, objconn) 'SELECT * FROM CLASS_VISITOR3 WHERE 1<>1
        dr1 = dtV.NewRow
        dr1("OCID") = TIMS.ClearSQM(Hid_OCID.Value)
        dr1("SOCID") = vSOCID
        dr1("RID") = sm.UserInfo.RID
        dr1("FILLFORMDATE") = TIMS.Cdate2(colArray(cst_aFILLFORMDATE))
        dr1("TYPE") = 1 'TYPE: 1.系統登打 2.電話訪查 3.報名網寫入

        dr1("Q1_1") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ1_1)))
        dr1("Q1_2") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ1_2)))
        dr1("Q1_3") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ1_3)))

        dr1("Q2_1") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ2_1)))
        dr1("Q2_2") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ2_2)))
        dr1("Q2_3") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ2_3)))
        dr1("Q2_4") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ2_4)))

        dr1("Q3_1") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ3_1)))
        dr1("Q3_2") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ3_2)))
        dr1("Q3_3") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ3_3)))

        dr1("Q4_1") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ4_1)))
        dr1("Q4_2") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ4_2)))
        dr1("Q4_3") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ4_3)))
        dr1("Q4_4") = TIMS.GetNullValue1(CHG_VALUE1(colArray(cst_aQ4_4)))

        dr1("SUGGESTION") = TIMS.GetNullValue1(TIMS.ClearSQM(colArray(cst_aSUGGESTION))) '其它意見
        dr1("SIGNER") = TIMS.ClearSQM(colArray(cst_aSIGNER)) '抽訪人員姓名 As Integer = 1 字串長度10
        'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
        dr1("DASOURCE") = 2

        dr1("MODIFYACCT") = sm.UserInfo.UserID 'TIMS.ClearSQM(colArray(cst_aMODIFYACCT))
        'dr1("MODIFYDATE") = TIMS.ClearSQM(colArray(cst_aMODIFYDATE))
        Return dr1
    End Function


    ''' <summary>
    ''' 檢查輸入資料
    ''' </summary>
    ''' <param name="colArray"></param>
    ''' <returns></returns>
    Function CheckImportData3(ByRef dtStd As DataTable, ByRef colArray As Array, ByRef xHash As Hashtable) As String
        Dim Reason As String = ""
        '欄位對應數
        If colArray.Length < cst_iMaxLength1 Then
            Reason &= String.Format("欄位對應有誤 ,{0}/{1}<BR>", colArray.Length, cst_iMaxLength1)
            Return Reason
        End If

        colArray(cst_aSTUDID2) = TIMS.ClearSQM(colArray(cst_aSTUDID2))
        If TIMS.IsNumeric2(colArray(cst_aSTUDID2)) Then colArray(cst_aSTUDID2) = TIMS.AddZero(Val(colArray(cst_aSTUDID2)), 3)

        Reason &= ChkValue1("填表日期", colArray(cst_aFILLFORMDATE), cst_st日期必填)
        Reason &= ChkValue1("抽訪人員姓名", colArray(cst_aSIGNER), cst_st字串, 10)
        Reason &= ChkValue1("學號", colArray(cst_aSTUDID2), cst_st數字必填, 10)
        Reason &= ChkValue1("答案1", colArray(cst_aQ1_1), cst_st數字必填)
        Reason &= ChkValue1("答案2", colArray(cst_aQ1_2), cst_st數字必填)
        Reason &= ChkValue1("答案3", colArray(cst_aQ1_3), cst_st數字必填)

        Reason &= ChkValue1("答案4", colArray(cst_aQ2_1), cst_st數字必填)
        Reason &= ChkValue1("答案5", colArray(cst_aQ2_2), cst_st數字必填)
        Reason &= ChkValue1("答案6", colArray(cst_aQ2_3), cst_st數字必填)
        'Reason &= ChkValue1("答案7", colArray(cst_aQ2_4), cst_st數字必填)

        Reason &= ChkValue1("答案8", colArray(cst_aQ3_1), cst_st數字必填)
        Reason &= ChkValue1("答案9", colArray(cst_aQ3_2), cst_st數字必填)
        Reason &= ChkValue1("答案10", colArray(cst_aQ3_3), cst_st數字必填)

        Reason &= ChkValue1("答案11", colArray(cst_aQ4_1), cst_st數字必填)
        Reason &= ChkValue1("答案12", colArray(cst_aQ4_2), cst_st數字必填)
        Reason &= ChkValue1("答案13", colArray(cst_aQ4_3), cst_st數字必填)
        Reason &= ChkValue1("答案14", colArray(cst_aQ4_4), cst_st數字必填)

        Reason &= ChkValue1("答案1", colArray(cst_aQ1_1), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案2", colArray(cst_aQ1_2), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案3", colArray(cst_aQ1_3), cst_st整數, 1, 5)

        Reason &= ChkValue1("答案4", colArray(cst_aQ2_1), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案5", colArray(cst_aQ2_2), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案6", colArray(cst_aQ2_3), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案7", colArray(cst_aQ2_4), cst_st整數, 1, 5)

        Reason &= ChkValue1("答案8", colArray(cst_aQ3_1), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案9", colArray(cst_aQ3_2), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案10", colArray(cst_aQ3_3), cst_st整數, 1, 5)

        Reason &= ChkValue1("答案11", colArray(cst_aQ4_1), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案12", colArray(cst_aQ4_2), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案13", colArray(cst_aQ4_3), cst_st整數, 1, 5)
        Reason &= ChkValue1("答案14", colArray(cst_aQ4_4), cst_st整數, 1, 5)

        Reason &= ChkValue1("其它意見", colArray(cst_aSUGGESTION), cst_st字串, 500)

        If Reason <> "" Then Return Reason

        Dim s_find As String = String.Format("STUDID2='{0}'", colArray(cst_aSTUDID2))
        If dtStd.Select(s_find).Length = 0 Then Reason &= String.Format("({0})學號有誤，該班無此學號", colArray(cst_aSTUDID2))
        hid_SOCIDvalue.Value = If(dtStd.Select(s_find).Length > 0, dtStd.Select(s_find)(0)("SOCID"), "")
        TIMS.SetMyValue2(xHash, "SOCID", hid_SOCIDvalue.Value)

        Dim flag_hava As Boolean = CHK_QUESTRAINING(Hid_OCID.Value, hid_SOCIDvalue.Value, objconn)
        If flag_hava Then Reason &= String.Format("學號({0})，該班已填資料，不可重新匯入", colArray(cst_aSTUDID2))

        Return Reason
    End Function

    '新增iCmd
    Sub Savedata3(ByRef iCmd As SqlCommand, ByVal dr1ARY As DataRow)
        With iCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = dr1ARY("OCID")
            .Parameters.Add("SOCID", SqlDbType.Int).Value = dr1ARY("SOCID")
            .Parameters.Add("RID", SqlDbType.VarChar).Value = dr1ARY("RID") 'sm.UserInfo.RID
            .Parameters.Add("FILLFORMDATE", SqlDbType.DateTime).Value = dr1ARY("FILLFORMDATE")
            .Parameters.Add("TYPE", SqlDbType.Int).Value = dr1ARY("TYPE") 'TYPE: 1.系統登打 2.電話訪查 3.報名網寫入

            .Parameters.Add("Q1_1", SqlDbType.Int).Value = dr1ARY("Q1_1")
            .Parameters.Add("Q1_2", SqlDbType.Int).Value = dr1ARY("Q1_2")
            .Parameters.Add("Q1_3", SqlDbType.Int).Value = dr1ARY("Q1_3")

            .Parameters.Add("Q2_1", SqlDbType.Int).Value = dr1ARY("Q2_1")
            .Parameters.Add("Q2_2", SqlDbType.Int).Value = dr1ARY("Q2_2")
            .Parameters.Add("Q2_3", SqlDbType.Int).Value = dr1ARY("Q2_3")
            .Parameters.Add("Q2_4", SqlDbType.Int).Value = dr1ARY("Q2_4")

            .Parameters.Add("Q3_1", SqlDbType.Int).Value = dr1ARY("Q3_1")
            .Parameters.Add("Q3_2", SqlDbType.Int).Value = dr1ARY("Q3_2")
            .Parameters.Add("Q3_3", SqlDbType.Int).Value = dr1ARY("Q3_3")

            .Parameters.Add("Q4_1", SqlDbType.Int).Value = dr1ARY("Q4_1")
            .Parameters.Add("Q4_2", SqlDbType.Int).Value = dr1ARY("Q4_2")
            .Parameters.Add("Q4_3", SqlDbType.Int).Value = dr1ARY("Q4_3")
            .Parameters.Add("Q4_4", SqlDbType.Int).Value = dr1ARY("Q4_4")

            .Parameters.Add("SUGGESTION", SqlDbType.NVarChar).Value = dr1ARY("SUGGESTION")
            .Parameters.Add("SIGNER", SqlDbType.NVarChar).Value = dr1ARY("SIGNER")
            '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統後台()
            .Parameters.Add("DASOURCE", SqlDbType.Int).Value = dr1ARY("DASOURCE")

            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
            'dt.Load(.ExecuteReader())
            .ExecuteNonQuery()
            'rst = .ExecuteScalar()
        End With
    End Sub

    ''' <summary>
    ''' 匯入 學員滿意度匯入 自辦受訓期間學員滿意度匯入功能
    ''' </summary>
    ''' <param name="FullFileName1"></param>
    Sub Sub_XLSImp3(ByRef FullFileName1 As String, ByRef dtStd As DataTable, ByRef drCC As DataRow)
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "FullFileName1", FullFileName1)
        TIMS.SetMyValue2(htSS, "FirstCol", "學號") '"身分證字號" '任1欄位名稱(必填)
        Dim Reason As String = ""
        '上傳檔案/取得內容
        Dim dt_xls As DataTable = TIMS.Get_File1data(File1, Reason, htSS, flag_File1_xls, flag_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        '儲存錯誤資料的DataTable
        Dim dtWrong As New DataTable
        Dim drWrong As DataRow = Nothing
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("STUDID2"))
        dtWrong.Columns.Add(New DataColumn("Reason"))

        Dim iRowIndex As Integer = 1
        Dim sReason As String = "" '儲存錯誤的原因
        Dim iCmd As SqlCommand = GetxICmd(objconn)
        Try
            '有資料
            sReason = ""
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                If iRowIndex <> 0 Then
                    Dim xHash As New Hashtable
                    TIMS.SetMyValue2(xHash, "OCID", Hid_OCID.Value)
                    Dim colArray As Array = dt_xls.Rows(i).ItemArray
                    sReason = CheckImportData3(dtStd, colArray, xHash) '依據匯入檔判斷錯誤

                    If sReason = "" Then
                        Dim dr1ARY As DataRow = ChgImpData(colArray, xHash) 'out@dr1
                        If dr1ARY IsNot Nothing Then Call Savedata3(iCmd, dr1ARY) '無錯誤存檔 '匯入資料
                    End If

                    If sReason <> "" Then
                        '錯誤資料，填入錯誤資料表
                        drWrong = dtWrong.NewRow
                        dtWrong.Rows.Add(drWrong)
                        drWrong("Index") = iRowIndex 'Index 第幾筆錯誤
                        Dim s_STUDID2 As String = If(colArray.Length > cst_aSTUDID2, colArray(cst_aSTUDID2), "查無此欄位")
                        drWrong("STUDID2") = s_STUDID2
                        drWrong("Reason") = sReason '原因
                    End If
                End If
                iRowIndex += 1
            Next
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, "儲存失敗!!")
            Exit Sub
        End Try

        '判斷匯出資料是否有誤
        'Dim explain, explain2 As String
        Dim explain As String = ""
        Dim explain2 As String = ""
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

        explain2 = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

        If dtWrong.Rows.Count > 0 Then
            Session("MyWrongTable") = dtWrong
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('CP_07_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            Exit Sub
        End If
        If sReason <> "" Then
            Common.MessageBox(Me, explain & sReason)
            Exit Sub
        End If
        Common.MessageBox(Me, explain)
        Exit Sub
    End Sub

    ''' <summary>
    ''' 學員滿意度匯入
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Btn_XlsImport_Click(sender As Object, e As EventArgs) Handles Btn_XlsImport.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "未選擇 職類/班別 無法匯入!!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "查無班級資料，請重新查詢!!")
            Exit Sub
        End If
        If Convert.ToString(drCC("TPLANID")) <> sm.UserInfo.TPlanID Then
            Common.MessageBox(Me, "班級資料計畫選擇有誤，請重新查詢!!")
            Exit Sub
        End If

        Hid_OCID.Value = OCIDValue1.Value
        Dim dtStd As DataTable = TIMS.Get_StudData2(Hid_OCID.Value, objconn)
        If dtStd Is Nothing Then
            Common.MessageBox(Me, "班級資料學員有誤，請重新查詢!!")
            Exit Sub
        End If
        If dtStd.Rows.Count = 0 Then
            Common.MessageBox(Me, "班級資料學員數有誤，請重新查詢!!")
            Exit Sub
        End If

        Dim sMyFileName As String = ""
        Dim sErrMsg As String = TIMS.ChkFile1(File1, sMyFileName, flag_File1_xls, flag_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "xls") Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "ods") Then Return
        End If
        Const Cst_FileSavePath As String = "~/CP/07/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        'Call sImport2(FullFileName1)
        Call Sub_XLSImp3(FullFileName1, dtStd, drCC)
    End Sub

    Const cst_st數字必填 As String = "數字必填"
    Const cst_st日期必填 As String = "日期必填"
    Const cst_st字串必填 As String = "字串必填"
    'Const cst_st文字必填 As String = "文字必填"
    Const cst_st字串 As String = "字串"
    Const cst_st整數 As String = "整數"

    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String) As String
        Return ChkValue1(fN1, vN1, sType, 0)
    End Function

    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String, ByVal iSize As Integer) As String
        Dim rst As String = ""
        Select Case sType
            Case cst_st字串
                If Convert.ToString(vN1) <> "" Then
                    If Convert.ToString(vN1).Length > iSize Then rst &= fN1 & "字串長度 必須小於等於" & iSize & "字數<br>"
                End If
            Case cst_st字串必填
                If Convert.ToString(vN1) <> "" Then
                    If Convert.ToString(vN1).Length > iSize Then rst &= fN1 & "字串長度 必須小於等於" & iSize & "字數<br>"
                Else
                    rst &= fN1 & " 為必填資料<br>"
                End If
            Case cst_st數字必填
                If Convert.ToString(vN1) <> "" Then
                    If Not IsNumeric(vN1) Then rst &= fN1 & "必需為數字<br>"
                Else
                    rst &= fN1 & "必須填寫<br>"
                End If
            Case cst_st日期必填
                If Convert.ToString(vN1) <> "" Then
                    If Not IsDate(vN1) Then
                        rst &= fN1 & "必須是西元年格式(yyyy/MM/dd)<br>"
                    Else
                        If CDate(vN1) < "1900/1/1" Or CDate(vN1) > "2100/1/1" Then rst &= fN1 & "範圍有誤<br>"
                    End If
                Else
                    rst &= fN1 & "必須填寫<br>"
                End If
        End Select
        Return rst
    End Function

    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String, ByVal iMin As Integer, ByVal iMax As Integer) As String
        Dim rst As String = ""
        Select Case sType
            Case cst_st整數
                If Convert.ToString(vN1) = "" Then Return rst
                If Not TIMS.IsInt(vN1) Then
                    rst &= fN1 & " 必須為整數數字<br>"
                    Return rst
                End If
                If Val(vN1) < Val(iMin) Then
                    rst &= fN1 & " 數字超過範圍<br>"
                    Return rst
                End If
                If Val(vN1) > Val(iMax) Then
                    rst &= fN1 & " 數字超過範圍<br>"
                    Return rst
                End If
        End Select
        Return rst
    End Function

End Class
