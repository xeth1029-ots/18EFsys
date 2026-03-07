Imports System.IO

Partial Class TC_01_027
    Inherits AuthBasePage

    'TC_01_027-產投使用
    'TC_01_007-非產投/TIMS使用
    Const cst_printFN1 As String = "TC_01_007"
    Const cst_printFN2 As String = "Teach"

    Dim sMemo As String = "" '(查詢原因)
    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False
    Dim ff3 As String = ""
    'colArray(
    Const cst_i計劃階層 As Integer = 0
    Const cst_i講師代碼 As Integer = 1
    Const cst_i講師姓名 As Integer = 2
    Const cst_i講師英文姓名 As Integer = 3
    Const cst_i身份別 As Integer = 4
    Const cst_i身分證字號 As Integer = 5
    Const cst_i出生日期 As Integer = 6
    Const cst_i性別 As Integer = 7
    Const cst_i主要職類 As Integer = 8 'GCODE2->TMID
    Const cst_i職稱 As Integer = 9
    Const cst_i內外聘 As Integer = 10
    Const cst_i師資別 As Integer = 11
    Const cst_i最高學歷 As Integer = 12
    Const cst_i畢業狀況 As Integer = 13
    Const cst_i學校名稱 As Integer = 14
    Const cst_i科系名稱 As Integer = 15
    Const cst_i聯絡電話 As Integer = 16
    Const cst_i行動電話 As Integer = 17
    Const cst_i電子郵件 As Integer = 18
    Const cst_i郵遞區號前3碼 As Integer = 19
    Const cst_i郵遞區號6碼 As Integer = 20
    Const cst_i通訊地址 As Integer = 21
    Const cst_i服務單位名稱 As Integer = 22
    Const cst_i年資 As Integer = 23
    Const cst_i服務部門 As Integer = 24
    Const cst_i服務單位電話 As Integer = 25
    Const cst_i服務單位傳真 As Integer = 26
    Const cst_i服務單位郵遞區號前3碼 As Integer = 27
    Const cst_i服務單位郵遞區號6碼 As Integer = 28
    Const cst_i服務單位地址 As Integer = 29
    Const cst_i服務單位一 As Integer = 30
    Const cst_i服務單位二 As Integer = 31
    Const cst_i服務單位三 As Integer = 32
    Const cst_i服務年資一 As Integer = 33
    Const cst_i服務年資二 As Integer = 34
    Const cst_i服務年資三 As Integer = 35
    Const cst_i服務職稱一 As Integer = 36
    Const cst_i服務職稱二 As Integer = 37
    Const cst_i服務職稱三 As Integer = 38
    Const cst_i服務期間一起日 As Integer = 39
    Const cst_i服務期間一迄日 As Integer = 40
    Const cst_i服務期間二起日 As Integer = 41
    Const cst_i服務期間二迄日 As Integer = 42
    Const cst_i服務期間三起日 As Integer = 43
    Const cst_i服務期間三迄日 As Integer = 44
    Const cst_i專長一 As Integer = 45
    Const cst_i專長二 As Integer = 46
    Const cst_i專長三 As Integer = 47
    Const cst_i專長四 As Integer = 48
    Const cst_i專長五 As Integer = 49
    Const cst_i譯著 As Integer = 50
    Const cst_i專業證照政府 As Integer = 51 'PROLICENSE1
    Const cst_i專業證照其他 As Integer = 52
    Const cst_i排課使用 As Integer = 53
    Const cst_i講師類別 As Integer = 54
    Const cst_i助教類別 As Integer = 55
    Const cst_i欄位長度 As Integer = 56 '最後欄位數+1

    'EXPORT1XLS
    Const cst_iCol_RIDLevel As Integer = 1
    Const cst_iCol_TeacherID As Integer = 2
    Const cst_iCol_TeachCName As Integer = 3
    Const cst_iCol_TeachEName As Integer = 4
    Const cst_iCol_PassPortNO As Integer = 5
    Const cst_iCol_IDNO As Integer = 6
    Const cst_iCol_Birthday As Integer = 7
    Const cst_iCol_Sex As Integer = 8
    Const cst_iCol_GCODE2 As Integer = 9
    Const cst_iCol_IVID As Integer = 10

    Const cst_iCol_KindEngage As Integer = 11
    Const cst_iCol_KindID As Integer = 12
    Const cst_iCol_DegreeID As Integer = 13
    Const cst_iCol_GraduateStatus As Integer = 14
    Const cst_iCol_SchoolName As Integer = 15
    Const cst_iCol_Department As Integer = 16
    Const cst_iCol_Phone As Integer = 17
    Const cst_iCol_Mobile As Integer = 18
    Const cst_iCol_Email As Integer = 19

    Const cst_iCol_AddressZip As Integer = 20
    Const cst_iCol_AddressZIP6W As Integer = 21
    Const cst_iCol_Address As Integer = 22
    Const cst_iCol_WorkOrg As Integer = 23
    Const cst_iCol_ExpYears As Integer = 24
    Const cst_iCol_ServDept As Integer = 25
    Const cst_iCol_WorkPhone As Integer = 26
    Const cst_iCol_Fax As Integer = 27

    Const cst_iCol_WorkZip As Integer = 28
    Const cst_iCol_WorkZIP6W As Integer = 29
    Const cst_iCol_Workaddr As Integer = 30
    Const cst_iCol_ExpUnit1 As Integer = 31
    Const cst_iCol_ExpUnit2 As Integer = 32
    Const cst_iCol_ExpUnit3 As Integer = 33
    Const cst_iCol_ExpYears1 As Integer = 34
    Const cst_iCol_ExpYears2 As Integer = 35
    Const cst_iCol_ExpYears3 As Integer = 36

    Const cst_iCol_EpINV1 As Integer = 37
    Const cst_iCol_EpINV2 As Integer = 38
    Const cst_iCol_EpINV3 As Integer = 39
    Const cst_iCol_ExpSDate1 As Integer = 40
    Const cst_iCol_ExpEDate1 As Integer = 41
    Const cst_iCol_ExpSDate2 As Integer = 42
    Const cst_iCol_ExpEDate2 As Integer = 43
    Const cst_iCol_ExpSDate3 As Integer = 44
    Const cst_iCol_ExpEDate3 As Integer = 45

    Const cst_iCol_Specialty1 As Integer = 46
    Const cst_iCol_Specialty2 As Integer = 47
    Const cst_iCol_Specialty3 As Integer = 48
    Const cst_iCol_Specialty4 As Integer = 49
    Const cst_iCol_Specialty5 As Integer = 50
    Const cst_iCol_TransBook As Integer = 51

    Const cst_iCol_ProLicense1 As Integer = 52
    Const cst_iCol_ProLicense2 As Integer = 53
    Const cst_iCol_WorkStatus As Integer = 54
    Const cst_iCol_TechType1 As Integer = 55
    Const cst_iCol_TechType2 As Integer = 56

    Dim ID_KindOfTeacher As DataTable = Nothing
    Dim ID_Invest As DataTable = Nothing
    Dim Key_Degree As DataTable = Nothing
    Dim Key_GradState As DataTable = Nothing
    'Dim dtTRAINTYPE As DataTable = Nothing
    Dim dtGOVCLASSCAST3 As DataTable = Nothing '主要職類table y2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection = Nothing

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload

        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        Hyperlink2.NavigateUrl = "../../Doc/ClassTeacher_v18b.zip"

        Dim work2015 As String = TIMS.Utl_GetConfigSet("work2015")
        hidLockTime2.Value = "2"
        If work2015 = "Y" Then hidLockTime2.Value = "1" '啟用鎖定。

        AddHandler Button1.Click, AddressOf SUtl_btnSearchData1 '查詢
        AddHandler Btn_XlsEmport.Click, AddressOf SUtl_btnSearchData1 '匯出 'Protected Sub Btn_XlsEmport_Click(sender As Object, e As EventArgs) Handles Btn_XlsEmport.Click
        AddHandler btndivPwdSubmit.Click, AddressOf SUtl_btnSearchData1 'hidSchBtnNum.value: 1.正常查詢 2.正常匯出

        '啟動個資法。
        'Button1.Attributes("onclick") = "aloader2on();"
        Button1.Attributes.Add("onclick", "return showLoginPwdDiv(1);")
        Button1.CommandName = "Button1"
        Btn_XlsEmport.Attributes.Add("onclick", "return showLoginPwdDiv(2);")
        Btn_XlsEmport.CommandName = "btnxlsemport"
        'btndivPwdSubmit.Attributes("onclick") = "aloader2on();"
        'Button1.Attributes("onclick") = "javascript:return search()"

        If Not IsPostBack Then
            msg.Text = ""
            'panelLoginDiv.Visible = False 'panelLoginDiv.Style.Item("display") = "none"
            labChkMsg.Text = ""
            eMeng.Style("display") = HidVeMeng.Value
            'VeMeng.Text = "none"
            center.Value = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            DropDownList3 = TIMS.Get_Invest(DropDownList3, objconn)

            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            '設定 資料與顯示 狀況
            Call CREATE1(0)
            DataGridTable.Visible = False
        End If

        Dim sql As String = ""
        sql = " SELECT KindID ,KINDNAME ,KINDENGAGE FROM ID_KINDOFTEACHER ORDER BY KINDID "
        ID_KindOfTeacher = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT IVID,InvestName FROM ID_INVEST ORDER BY IVID "
        ID_Invest = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT DEGREEID ,NAME FROM KEY_DEGREE ORDER BY DEGREEID "
        Key_Degree = DbAccess.GetDataTable(sql, objconn)
        '取出dt-畢業狀況代碼-師資資料設定
        Key_GradState = TIMS.Get_GradStateDt2(objconn)
        'sql = "SELECT TMID,BUSID,BUSNAME,JOBID,JOBNAME,TRAINID,TRAINNAME FROM VIEW_TRAINTYPE WHERE TRAINID IS NOT NULL ORDER BY BUSID,JOBID,TMID"
        'dtTRAINTYPE = DbAccess.GetDataTable(sql, objconn)
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " SELECT tt.TMID, tt.BUSID, tt.BUSNAME, tt.JOBID, tt.JOBNAME, tt.TRAINID, tt.TRAINNAME, g.GCODE2, g.CNAME" & vbCrLf
        'sql &= " FROM V_GOVCLASSCAST3 g" & vbCrLf
        'sql &= " JOIN VIEW_TRAINTYPE tt ON tt.TMID = g.TMID AND tt.TRAINID IS NOT NULL" & vbCrLf
        'sql &= " ORDER BY tt.BUSID, tt.JOBID, tt.TMID" & vbCrLf
        'dtGOVCLASSCAST3 = DbAccess.GetDataTable(sql, objconn)
        dtGOVCLASSCAST3 = TIMS.Get_GOVCLASSCAST3dt(dtGOVCLASSCAST3, objconn)

        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button5.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
            HistoryRID.Attributes("onclick") = "ShowFrame();"
            center.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            USE_MySearchStr()
        End If
    End Sub

    Sub USE_MySearchStr()
        If Session("MySearchStr") Is Nothing Then Return
        Dim MyValue As String = ""
        Dim strSession As String = Session("MySearchStr")
        Session("MySearchStr") = Nothing

        center.Value = TIMS.GetMyValue(strSession, "center")
        RIDValue.Value = TIMS.GetMyValue(strSession, "RIDValue")
        'MyValue = TIMS.GetMyValue(strSession, "DropDownList1")
        'If MyValue <> "" Then Common.SetListItem(DropDownList1, MyValue)
        TextBox2.Text = TIMS.GetMyValue(strSession, "TextBox2")
        TextBox3.Text = TIMS.GetMyValue(strSession, "TextBox3")
        MyValue = TIMS.GetMyValue(strSession, "DropDownList2")
        If MyValue <> "" Then Common.SetListItem(DropDownList2, MyValue)
        TextBox4.Text = TIMS.GetMyValue(strSession, "TextBox4")
        MyValue = TIMS.GetMyValue(strSession, "DropDownList3")
        If MyValue <> "" Then Common.SetListItem(DropDownList3, MyValue)
        TB_career_id.Text = TIMS.GetMyValue(strSession, "TB_career_id")
        trainValue.Value = TIMS.GetMyValue(strSession, "trainValue")
        jobValue.Value = TIMS.GetMyValue(strSession, "jobValue")
        MyValue = TIMS.GetMyValue(strSession, "DropDownList4")
        If MyValue <> "" Then
            Common.SetListItem(DropDownList4, MyValue)
            Call Sub_DDL4Sel()
        End If
        MyValue = TIMS.GetMyValue(strSession, "DropDownList1")
        If MyValue <> "" Then Common.SetListItem(DropDownList1, MyValue)
        MyValue = TIMS.GetMyValue(strSession, "PageIndex")
        If MyValue <> "" Then PageControler1.PageIndex = MyValue
        MyValue = TIMS.GetMyValue(strSession, "Button1")
        'Button1_Click(sender, e)
        If MyValue = "True" Then Call gClickSearchButton()
    End Sub

    '設定 資料與顯示 狀況！
    Sub CREATE1(ByVal num As Integer)
        'num 0:第一次呼叫 --請選擇-- 1:內聘 2:外聘
        Select Case num
            Case 0
                DropDownList1.Items.Clear()
                DropDownList1.Items.Add(New ListItem("--請選擇內外聘--", 0))
                tr_techtype12.Visible = False
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then tr_techtype12.Visible = True   '產投  '顯示 講師 助教 (類別) 
            Case Else '1.2
                DropDownList1 = TIMS.Get_KindOfTeacher(DropDownList1, CStr(num), "1", objconn)
        End Select
    End Sub

    Sub GetSearchStr()
        Session("MySearchStr") = Nothing

        Dim v_DropDownList1 As String = TIMS.GetListValue(DropDownList1)
        Dim v_DropDownList2 As String = TIMS.GetListValue(DropDownList2)
        Dim v_DropDownList3 As String = TIMS.GetListValue(DropDownList3)
        Dim v_DropDownList4 As String = TIMS.GetListValue(DropDownList4)

        Dim sMySearchStr As String = ""
        sMySearchStr = "center=" & TIMS.ClearSQM(center.Value)
        sMySearchStr += "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        sMySearchStr += "&DropDownList1=" & v_DropDownList1
        sMySearchStr += "&TextBox2=" & TIMS.ClearSQM(TextBox2.Text)
        sMySearchStr += "&TextBox3=" & TIMS.ClearSQM(TextBox3.Text)
        sMySearchStr += "&DropDownList2=" & v_DropDownList2
        sMySearchStr += "&TextBox4=" & TIMS.ClearSQM(TextBox4.Text)
        sMySearchStr += "&DropDownList3=" & v_DropDownList3
        sMySearchStr += "&TB_career_id=" & TIMS.ClearSQM(TB_career_id.Text)
        sMySearchStr += "&trainValue=" & TIMS.ClearSQM(trainValue.Value)
        sMySearchStr += "&jobValue=" & TIMS.ClearSQM(jobValue.Value)
        sMySearchStr += "&DropDownList4=" & v_DropDownList4
        sMySearchStr += "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        sMySearchStr += "&Button1=" & DataGrid1.Visible
        Session("MySearchStr") = sMySearchStr
    End Sub

    '刪除
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Select Case e.CommandName
            Case "edit"
                Call GetSearchStr()
                Dim url1 As String = $"TC_01_027_add.aspx?proecess=edit&serial={e.CommandArgument}&ID={TIMS.Get_MRqID(Me)}"
                TIMS.Utl_Redirect(Me, objconn, url1)
            Case "print"

            Case "del"
                'e.CommandArgument@TechID
                If Convert.ToString(e.CommandArgument) = "" Then
                    Common.MessageBox(Me, "傳入參數有誤，請重新查詢")
                    Exit Sub
                End If
                'e.CommandArgument@TechID
                Dim sTechID As String = e.CommandArgument
                Dim tmpTeacherName As String = TIMS.Get_TeachCName(sTechID, objconn) 'TIMS.Get_TeacherName(e.CommandArgument)
                If tmpTeacherName = "" Then
                    Common.MessageBox(Me, "查無該師姓名，請重新查詢")
                    Exit Sub
                End If
                If Not gDelTeach_TeacherInfo(e.CommandArgument) Then
                    Common.MessageBox(Me, "使用中，不可刪除")
                    Exit Sub
                End If
                Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) 'rblWorkMode.SelectedValue
                sMemo = $"&動作=刪除&NAME={tmpTeacherName}"
                '寫入Log查詢(SubInsAccountLog1(Auth_Accountlog))
                Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm刪除, v_rblWorkMode, "", sMemo)
                Common.MessageBox(Me, "刪除完成")
                gClickSearchButton() '查詢
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.Cells(0).Text = "序號"
                If Me.cb_CourID.Checked Then e.Item.Cells(0).Text = "匯入用<BR>代碼"
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = DataGrid1.CurrentPageIndex * DataGrid1.PageSize + e.Item.ItemIndex + 1
                If Me.cb_CourID.Checked Then e.Item.Cells(0).Text = CStr(drv("TechID"))
                'Dim row() As DataRow
                If Convert.ToString(drv("KindID")) <> "" Then
                    ff3 = "KindID='" & Convert.ToString(drv("KindID")) & "'"
                    If ID_KindOfTeacher.Select(ff3).Length <> 0 Then e.Item.Cells(4).Text = ID_KindOfTeacher.Select(ff3)(0)("KindName")
                End If
                Dim strKindEngage As String = Convert.ToString(drv("KindEngage"))
                Select Case Convert.ToString(drv("KindEngage"))
                    Case "1"
                        strKindEngage = "內聘(專任)"
                    Case "2"
                        strKindEngage = "外聘(兼任)"
                End Select
                e.Item.Cells(5).Text = strKindEngage
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim lbtDel As LinkButton = e.Item.FindControl("lbtDel")
                Dim lbtPrt As LinkButton = e.Item.FindControl("lbtPrt")
                '修改/檢視
                'Dim but As Button = e.Item.Cells(6).FindControl("Button3")
                lbtEdit.CommandArgument = Convert.ToString(drv("TechID"))
                lbtEdit.Text = "檢視"
                If sm.UserInfo.RID = drv("RID") Then
                    lbtEdit.Text = "修改"
                ElseIf Len(sm.UserInfo.RID.ToString) = 1 Then
                    lbtEdit.Text = "修改"
                End If
                'If FunDr("Mod") = "1" Then lbtEdit.Enabled = True
                'lbtEdit.Enabled = False
                'If au.blnCanMod Then lbtEdit.Enabled = True
                '刪除鈕
                'Dim btndelete As Button = e.Item.FindControl("btndelete")
                lbtDel.CommandArgument = Convert.ToString(drv("TechID"))
                lbtDel.Visible = False
                If sm.UserInfo.LID <= 1 Then
                    lbtDel.Attributes("onclick") = "javascript:return confirm('此動作會刪除師資資料，是否確定刪除?');"
                    lbtDel.Visible = True
                End If
                '列印師資資料
                'but = e.Item.Cells(6).FindControl("Button4")  '列印師資資料
                lbtPrt.CommandArgument = Convert.ToString(drv("TechID"))
                lbtPrt.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_printFN2, "TechID=" & drv("TechID") & "")
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then lbtPrt.Visible = False '產投  '不顯示 列印師資資料
        End Select
    End Sub

    '刪除 (含檢查使用狀況動作)
    Function gDelTeach_TeacherInfo(ByVal TechID As String) As Boolean
        Dim Rst As Boolean = False ' 刪除 有異常
        Dim flagCanDelete As Boolean = True '可以刪除
        Dim sql As String = ""
        'Dim dr As DataRow
        Dim dt As DataTable
        If TechID.Trim <> "" Then
            TechID = TechID.Trim
            If IsNumeric(TechID) Then
                sql = "" & vbCrLf
                If flagCanDelete Then
                    '開班老師檔(產業人才投資方案)
                    sql = ""
                    sql &= " SELECT DISTINCT 'x1' x FROM CLASS_TEACHER WHERE TechID = '" & TechID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then flagCanDelete = False '有資料，不可以刪除 
                End If
                If flagCanDelete Then
                    '不預告實地抽查訪視記錄檔
                    sql = ""
                    sql &= " SELECT DISTINCT 'x21' x FROM CLASS_UNEXPECTVISITOR WHERE TechID = '" & TechID & "'" & vbCrLf
                    sql &= " UNION" & vbCrLf
                    sql &= " SELECT DISTINCT 'x22' x FROM CLASS_UNEXPECTVISITOR WHERE TechID2 = '" & TechID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then flagCanDelete = False '有資料，不可以刪除 
                End If
                If flagCanDelete Then
                    '-班級申請老師檔(產學訓)
                    sql = ""
                    sql &= " SELECT DISTINCT 'x3' x FROM PLAN_TEACHER WHERE TechID = '" & TechID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then flagCanDelete = False '有資料，不可以刪除 
                End If
                If flagCanDelete Then
                    '計畫訓練內容簡介(95年度)(97產學訓課程大綱)
                    'Sql += " union" & vbCrLf
                    sql = ""
                    sql &= " SELECT DISTINCT 'x4' x FROM Plan_TrainDesc WHERE TechID = '" & TechID & "'" & vbCrLf
                    sql &= " UNION" & vbCrLf
                    sql &= " SELECT DISTINCT 'x5' x FROM Plan_TrainDesc WHERE TechID2 = '" & TechID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then flagCanDelete = False '有資料，不可以刪除 
                End If
                If flagCanDelete Then
                    '排課資訊
                    sql = "" & vbCrLf
                    sql &= " WITH x6 AS (SELECT DISTINCT 'x6' x FROM MVIEW_CLASS_SCHEDULE WHERE TechID = '" & TechID & "')" & vbCrLf
                    sql &= " ,x7 AS (SELECT DISTINCT 'x7' x FROM MVIEW_CLASS_SCHEDULE WHERE TechID2 = '" & TechID & "')" & vbCrLf
                    sql &= " SELECT * FROM x6 UNION SELECT * from x7" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then flagCanDelete = False '有資料，不可以刪除 
                End If
                If flagCanDelete Then
                    '排課資訊
                    sql = "" & vbCrLf
                    sql &= " WITH x6 AS (SELECT DISTINCT 'x6' x FROM VIEW_CLASS_SCHEDULE WHERE TechID = '" & TechID & "')" & vbCrLf
                    sql &= " ,x7 AS (SELECT DISTINCT 'x7' x FROM VIEW_CLASS_SCHEDULE WHERE TechID2 = '" & TechID & "')" & vbCrLf
                    sql &= " SELECT * FROM x6 UNION SELECT * FROM x7" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If dt.Rows.Count > 0 Then flagCanDelete = False '有資料，不可以刪除 
                End If
                If flagCanDelete Then
                    '無使用資料 '可以刪除 
                    Try
                        sql = " DELETE TEACH_TEACHERINFO WHERE TechID = '" & TechID & "' "
                        DbAccess.ExecuteNonQuery(sql, objconn)
                        Rst = True ' 刪除 完成
                    Catch ex As Exception
                        Throw ex '刪除失敗
                    End Try
                End If
            End If
        End If
        Return Rst
    End Function

    '設定 資料與顯示 狀況
    Private Sub DropDownList4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DropDownList4.SelectedIndexChanged
        Call Sub_DDL4Sel() '設定 資料與顯示 狀況
    End Sub

    '設定 資料與顯示 狀況
    Sub Sub_DDL4Sel()
        Select Case DropDownList4.SelectedValue
            Case "1", "2"
                Call CREATE1(DropDownList4.SelectedValue)
            Case Else
                Call CREATE1(0)
        End Select
    End Sub

    '檢查輸入資料
    Function CheckImportData(ByRef colArray As Array) As String
        Dim Reason As String = ""
        Dim sql As String = ""
        Dim strCol1 As String = ""

        If colArray.Length < cst_i欄位長度 Then
            Reason += "欄位對應有誤，資料欄位不足" & cst_i欄位長度 & "個欄位<BR>"
        Else
            colArray = TIMS.ChangeColArray(colArray)
            If colArray(cst_i講師代碼).ToString = "" Then
                Reason += "講師代碼必須填寫<Br>"
            Else
                If (colArray(cst_i講師代碼).ToString).Length > 10 Then Reason += "講師代碼不符合<BR>"
            End If
            If colArray(cst_i講師姓名).ToString = "" Then Reason += "講師姓名必須填寫<Br>"
            If colArray(cst_i講師英文姓名).ToString <> "" Then
                colArray(cst_i講師英文姓名) = TIMS.ChangeIDNO(colArray(cst_i講師英文姓名), " ") '講師英文姓名
                If (colArray(cst_i講師英文姓名).ToString).Length > 30 Then Reason += "講師英文姓名 過長應小於等於30字字數<BR>"
            End If
            If colArray(cst_i身份別).ToString = "" Then
                Reason += "身份別必須填寫<Br>"
            Else
                Select Case colArray(cst_i身份別).ToString
                    Case "1", "2"
                    Case Else
                        Reason += "身份別必須輸入1或2(1.本國,2.外籍)<Br>"
                End Select
            End If
            If colArray(cst_i身分證字號).ToString = "" Then
                Reason += "身分證必須填寫<Br>"
            Else
                If (colArray(cst_i身分證字號).ToString.Length <> 10) Then Reason += "身分證字數不符合<BR>"
            End If
            If Reason <> "" Then Return Reason '上述有錯誤離開
            colArray(cst_i身分證字號) = TIMS.ClearSQM(colArray(cst_i身分證字號))
            colArray(cst_i身分證字號) = TIMS.ChangeIDNO(colArray(cst_i身分證字號))
            Select Case colArray(cst_i身份別).ToString
                Case "1" '本國
                    If Not TIMS.CheckIDNO(colArray(cst_i身分證字號)) Then
                        Reason += "身分證號碼有誤<BR>"
                    End If
                Case "2" '外藉
                    Dim nsIDNO As String = colArray(cst_i身分證字號)
                    '2:居留證 4:居留證2021
                    Dim flag2 As Boolean = TIMS.CheckIDNO2(nsIDNO, 2)
                    Dim flag4 As Boolean = TIMS.CheckIDNO2(nsIDNO, 4)
                    If Not flag2 AndAlso Not flag4 Then
                        Reason += "身份別為外藉，居留證號有誤<BR>"
                    End If

                Case Else
                    Reason += "身份別必須輸入1或2(1.本國,2.外籍)<Br>"
            End Select

            RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
            Dim vRID As String = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
            Dim vIDNO As String = TIMS.ClearSQM(colArray(cst_i身分證字號).ToString)
            sql = ""
            sql &= " SELECT COUNT(1) CNT"
            sql &= " FROM TEACH_TEACHERINFO"
            sql &= " WHERE RID = '" & vRID & "' AND IDNO = '" & vIDNO & "' "
            Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
            If CInt(dr("CNT").ToString) > 0 Then Reason += "同一計劃，有相同身份號碼，重複輸入<BR>"
            If Convert.ToString(colArray(cst_i出生日期)) = "" Then
                'Reason += "出生日期必須填寫<Br>"
            Else
                If IsDate(Convert.ToString(colArray(cst_i出生日期))) = False Then
                    Reason += "出生日期必須是西元年格式(yyyy/mm/dd)<BR>"
                Else
                    Try
                        colArray(cst_i出生日期) = CDate(Convert.ToString(colArray(cst_i出生日期))).ToString("yyyy/MM/dd")
                        If CDate(Convert.ToString(colArray(cst_i出生日期))) < "1900/1/1" OrElse CDate(Convert.ToString(colArray(cst_i出生日期))) > "2100/1/1" Then Reason += "出生日期範圍有誤<BR>"
                    Catch ex As Exception
                        Reason += "出生日期必須是西元年格式(yyyy/mm/dd)<BR>"
                    End Try
                End If
            End If
            If colArray(cst_i性別).ToString = "" Then
                Reason += "必須填寫性別<BR>"
            Else
                colArray(cst_i性別) = UCase(colArray(cst_i性別))
                Select Case colArray(cst_i性別).ToString
                    Case "M", "F"
                    Case Else
                        Reason += "性別代號只能是M或者是F<BR>"
                End Select
            End If


            'If colArray(cst_i主要職類).ToString = "" Then
            '    Reason += "主要職類必須填寫<Br>"
            'Else
            '    ff = "TMID='" & colArray(cst_i主要職類) & "'"
            '    If dtTRAINTYPE.Select(ff).Length = 0 Then
            '        Reason += "主要職類不在鍵詞範圍內，請確認<BR>"
            '    End If
            'End If
            colArray(cst_i主要職類) = TIMS.ClearSQM(colArray(cst_i主要職類))
            If colArray(cst_i主要職類).ToString = "" Then
                Reason += "主要職類必須填寫<Br>"
            Else
                ff3 = "GCODE2='" & colArray(cst_i主要職類) & "'"
                If dtGOVCLASSCAST3.Select(ff3).Length = 0 Then Reason += "主要職類不在鍵詞範圍內，請確認<BR>"
            End If

            If colArray(cst_i職稱).ToString = "" Then
                Reason += "職稱填寫是否正確<Br>"
            Else
                If RIDValue.Value.Length = 1 Then '分署(中心)以上單位
                    If IsNumeric(colArray(cst_i職稱)) = False Then
                        Reason += "職稱必需為數字<BR>"
                    Else
                        colArray(cst_i職稱) = TIMS.ChangeIDNO(colArray(cst_i職稱))
                        If Len(colArray(cst_i職稱).ToString) < 2 Then colArray(cst_i職稱) = "0" & colArray(cst_i職稱).ToString
                        ff3 = "IVID='" & colArray(cst_i職稱) & "'"
                        If ID_Invest.Select(ff3).Length = 0 Then Reason += "職稱不在鍵詞範圍內，請確認<BR>"
                    End If
                Else
                    '委訓單位(輸入INVEST)
                End If
            End If

            If colArray(cst_i內外聘).ToString = "" Then
                Reason += "內外聘必須填寫<Br>"
            Else
                If IsNumeric(colArray(cst_i內外聘)) = False Then
                    Reason += "內外聘 必需為數字(1.內聘(專任)  2.外聘(兼任)(委訓單位))<BR>"
                Else
                    Select Case Convert.ToString(Val(colArray(cst_i內外聘)))
                        Case "1", "2"
                            colArray(cst_i內外聘) = Convert.ToString(Val(colArray(cst_i內外聘)))
                        Case Else
                            Reason += "內外聘必需為數字(1.內聘(專任)  2.外聘(兼任)(委訓單位))<BR>"
                    End Select
                End If
            End If
            If colArray(cst_i師資別).ToString = "" Then
                Reason += "師資別必須填寫<Br>"
            Else
                If IsNumeric(colArray(cst_i師資別)) = False Then
                    Reason += "師資別必需為數字<BR>"
                Else
                    If Reason = "" Then
                        If RIDValue.Value.Length = 1 Then '分署(中心)以上單位
                            ff3 = "KindID='" & colArray(cst_i師資別) & "' AND KindEngage='" & Convert.ToString(CInt(colArray(cst_i內外聘))) & "'"
                            If ID_KindOfTeacher.Select(ff3).Length = 0 Then Reason += "師資別不在鍵詞範圍內，請確認<BR>"
                        Else
                            '委訓單位只能輸入130講師
                            If Convert.ToString(colArray(cst_i師資別)) <> "130" Then Reason += "委訓單位: 師資別只能輸入代碼:130(講師)<BR>"
                        End If
                    End If
                End If
            End If
            If colArray(cst_i最高學歷).ToString = "" Then
                Reason += "最高學歷必須填寫<Br>"
            Else
                If Len(colArray(cst_i最高學歷).ToString) < 2 Then colArray(cst_i最高學歷) = "0" & colArray(cst_i最高學歷).ToString
                ff3 = "DegreeID='" & colArray(cst_i最高學歷) & "'"
                If Not Key_Degree.Select(ff3).Length > 0 Then Reason += "最高學歷不在鍵詞範圍內，請確認<BR>"
            End If
            If colArray(cst_i畢業狀況).ToString = "" Then
                Reason += "畢業狀況必須填寫<Br>"
            Else
                If Len(colArray(cst_i畢業狀況).ToString) < 2 Then colArray(cst_i畢業狀況) = "0" & colArray(cst_i畢業狀況).ToString
                ff3 = "GradID='" & colArray(cst_i畢業狀況) & "'"
                If Not Key_GradState.Select(ff3).Length > 0 Then Reason += "畢業狀況不在鍵詞範圍內，請確認<BR>"
            End If
            colArray(cst_i學校名稱) = TIMS.ClearSQM(colArray(cst_i學校名稱))
            If colArray(cst_i學校名稱).ToString <> "" Then
                If (colArray(cst_i學校名稱).ToString.Length > 30) Then Reason += "學校名稱必須小於等於中文25字<BR>"
            Else
                Reason += "學校名稱必須填寫<Br>"
            End If
            colArray(cst_i科系名稱) = TIMS.ClearSQM(colArray(cst_i科系名稱))
            If colArray(cst_i科系名稱).ToString <> "" Then
                If (colArray(cst_i科系名稱).ToString.Length > 25) Then Reason += "科系名稱必須小於等於中文25字<BR>"
            Else
                Reason += "科系名稱必須填寫<Br>"
            End If
            If colArray(cst_i聯絡電話).ToString = "" Then
                Reason += "聯絡電話必須填寫<Br>"
            Else
                If (colArray(cst_i聯絡電話).ToString.Length > 15) Then Reason += "聯絡電話必須小於等於15字字數<BR>"
            End If
            If colArray(cst_i行動電話).ToString <> "" Then
                If (colArray(cst_i行動電話).ToString.Length > 20) Then Reason += "行動電話必須小於等於20字字數<BR>"
            End If
            If colArray(cst_i電子郵件).ToString <> "" Then
                If (colArray(cst_i電子郵件).ToString.Length > 64) Then Reason += "E_mail必須小於等於64字字數<BR>"
            End If
            If colArray(cst_i郵遞區號前3碼).ToString = "" Then
                Reason += "通訊地址郵遞區號前3碼必須填寫<BR>"
            Else
                If IsNumeric(colArray(cst_i郵遞區號前3碼)) = False Then
                    Reason += "通訊地址郵遞區號前3碼必須為數字<BR>"
                Else
                    If Len(Convert.ToString(colArray(cst_i郵遞區號前3碼)).Trim) <> 3 Then Reason += "通訊地址郵遞區號前3碼必須為3碼<BR>"
                End If
            End If

            colArray(cst_i郵遞區號6碼) = TIMS.ClearSQM(colArray(cst_i郵遞區號6碼))
            Dim s_tmpzip6w1 As String = colArray(cst_i郵遞區號6碼)
            If s_tmpzip6w1 = "" Then
                Reason += "通訊地址郵遞區號6碼必須填寫<BR>"
            Else
                If Not IsNumeric(s_tmpzip6w1) Then
                    Reason += "通訊地址郵遞區號6碼必須為數字<BR>"
                Else
                    Dim ilen As Integer = Len(s_tmpzip6w1)
                    If ilen <> 5 AndAlso ilen <> 6 Then Reason += "通訊地址郵遞區號6碼長度必須為5碼或6碼<BR>"
                End If
            End If
            If colArray(cst_i通訊地址).ToString = "" Then
                Reason += "通訊地址必須填寫<BR>"
            Else
                If (colArray(cst_i通訊地址).ToString.Length > 50) Then Reason += "通訊地址 必須小於等於 50字字數<BR>"
            End If

            If colArray(cst_i服務單位名稱).ToString = "" Then
                Reason += "服務單位名稱必須填寫<BR>"
            Else
                If (colArray(cst_i服務單位名稱).ToString.Length > 50) Then Reason += "服務單位名稱 必須小於等於 50字字數<BR>"
            End If
            If colArray(cst_i年資).ToString <> "" Then
                If IsNumeric(colArray(cst_i年資)) = False Then Reason += "服務年資必須為數字<BR>"
            End If
            If colArray(cst_i服務部門).ToString <> "" Then
                colArray(cst_i服務部門) = Trim(colArray(cst_i服務部門)) '服務部門
                If (colArray(cst_i服務部門).ToString).Length > 50 Then Reason += "服務部門 過長應小於等於50字字數<BR>"
            End If
            If colArray(cst_i服務單位電話).ToString = "" Then
                Reason += "服務單位電話 必須填寫<BR>"
            Else
                If (colArray(cst_i服務單位電話).ToString.Length > 20) Then Reason += "服務單位電話 必須小於等於 20字字數<BR>"
            End If
            If colArray(cst_i服務單位傳真).ToString <> "" Then
                colArray(cst_i服務單位傳真) = Trim(colArray(cst_i服務單位傳真)) '服務單位傳真
                If (colArray(cst_i服務單位傳真).ToString).Length > 20 Then Reason += "服務單位傳真 過長應小於等於20字字數<BR>"
            End If
            If colArray(cst_i服務單位郵遞區號前3碼).ToString <> "" Then
                If IsNumeric(colArray(cst_i服務單位郵遞區號前3碼)) = False Then
                    Reason += "服務單位郵遞區號前3碼必須為數字<BR>"
                Else
                    If Len(Convert.ToString(colArray(cst_i服務單位郵遞區號前3碼)).Trim) <> 3 Then Reason += "服務單位郵遞區號前3碼必須為3碼<BR>"
                End If
            End If

            colArray(cst_i服務單位郵遞區號6碼) = TIMS.ClearSQM(colArray(cst_i服務單位郵遞區號6碼))
            Dim s_tmpzip6w2 As String = colArray(cst_i郵遞區號6碼)
            If s_tmpzip6w2 = "" Then
                Reason += "服務單位郵遞區號6碼必須填寫<BR>"
            Else
                If Not IsNumeric(s_tmpzip6w2) Then
                    Reason += "服務單位郵遞區號6碼必須為數字<BR>"
                Else
                    Dim ilen As Integer = Len(s_tmpzip6w2)
                    If ilen <> 5 AndAlso ilen <> 6 Then Reason += "服務單位郵遞區號6碼長度必須為5碼或6碼<BR>"
                End If
            End If
            If colArray(cst_i服務單位地址).ToString <> "" Then
                colArray(cst_i服務單位地址) = Trim(colArray(cst_i服務單位地址)) '服務單位地址
                If (colArray(cst_i服務單位地址).ToString).Length > 50 Then Reason += "服務單位地址 過長應小於等於50字字數<BR>"
            End If
            'If colArray(30).ToString = "" Then Reason += "服務單位一必須填寫<BR>"
            If colArray(cst_i服務單位一).ToString <> "" Then
                colArray(cst_i服務單位一) = Trim(colArray(cst_i服務單位一))
                If (colArray(cst_i服務單位一).ToString.Length > 50) Then Reason += "服務單位一 必須小於等於 50字字數<BR>"
            End If
            If colArray(cst_i服務單位二).ToString <> "" Then
                colArray(cst_i服務單位二) = Trim(colArray(cst_i服務單位二))
                If (colArray(cst_i服務單位二).ToString.Length > 50) Then Reason += "服務單位二 必須小於等於 50字字數<BR>"
            End If
            If colArray(cst_i服務單位三).ToString <> "" Then
                colArray(cst_i服務單位三) = Trim(colArray(cst_i服務單位三))
                If (colArray(cst_i服務單位三).ToString.Length > 50) Then Reason += "服務單位三 必須小於等於 50字字數<BR>"
            End If
            'If colArray(33).ToString = "" Then Reason += "服務年資一必須填寫<BR>"
            If colArray(cst_i服務年資一).ToString <> "" Then
                If Not IsNumeric(colArray(cst_i服務年資一).ToString) Then
                    Reason += "服務年資一必須填寫整數數字格式<BR>"
                Else
                    If colArray(cst_i服務年資一).ToString.Trim.IndexOf(".") > -1 Then Reason += "服務年資一必須填寫整數數字格式<BR>"
                    colArray(cst_i服務年資一) = CInt(colArray(cst_i服務年資一))
                End If
            End If
            If colArray(cst_i服務年資二).ToString <> "" Then
                If Not IsNumeric(colArray(cst_i服務年資二).ToString) Then
                    Reason += "服務年資二必須填寫整數數字格式<BR>"
                Else
                    If colArray(cst_i服務年資二).ToString.Trim.IndexOf(".") > -1 Then Reason += "服務年資二必須填寫整數數字格式<BR>"
                    colArray(cst_i服務年資二) = CInt(colArray(cst_i服務年資二))
                End If
            End If
            If colArray(cst_i服務年資三).ToString <> "" Then
                If Not IsNumeric(colArray(cst_i服務年資三).ToString) Then
                    Reason += "服務年資三必須填寫整數數字格式<BR>"
                Else
                    If colArray(cst_i服務年資三).ToString.Trim.IndexOf(".") > -1 Then Reason += "服務年資三必須填寫整數數字格式<BR>"
                    colArray(cst_i服務年資三) = CInt(colArray(cst_i服務年資三))
                End If
            End If
            colArray(cst_i服務職稱一) = TIMS.ClearSQM(colArray(cst_i服務職稱一))
            If colArray(cst_i服務職稱一).ToString <> "" Then
                If (colArray(cst_i服務職稱一).ToString.Length > 50) Then Reason += "服務職稱一 必須小於等於 50字字數<BR>"
            End If
            colArray(cst_i服務職稱二) = TIMS.ClearSQM(colArray(cst_i服務職稱二))
            If colArray(cst_i服務職稱二).ToString <> "" Then
                If (colArray(cst_i服務職稱二).ToString.Length > 50) Then Reason += "服務職稱二 必須小於等於 50字字數<BR>"
            End If
            colArray(cst_i服務職稱三) = TIMS.ClearSQM(colArray(cst_i服務職稱三))
            If colArray(cst_i服務職稱三).ToString <> "" Then
                If (colArray(cst_i服務職稱三).ToString.Length > 50) Then Reason += "服務職稱三 必須小於等於 50字字數<BR>"
            End If
            '36~41
            For intCol As Integer = cst_i服務期間一起日 To cst_i服務期間三迄日
                'intCol = ji 'cst_i服務期間一起日
                Select Case intCol
                    Case cst_i服務期間一起日
                        strCol1 = "服務期間一起日"
                    Case cst_i服務期間一迄日
                        strCol1 = "服務期間一迄日"
                    Case cst_i服務期間二起日
                        strCol1 = "服務期間二起日"
                    Case cst_i服務期間二迄日
                        strCol1 = "服務期間二迄日"
                    Case cst_i服務期間三起日
                        strCol1 = "服務期間三起日"
                    Case cst_i服務期間三迄日
                        strCol1 = "服務期間三迄日"
                End Select
                If colArray(intCol).ToString = "" Then
                    'Reason += strCol1 & " 必須填寫<Br>"
                Else
                    If IsDate(colArray(intCol)) = False Then
                        Reason += strCol1 & " 必須是西元年格式(yyyy/mm/dd)<BR>"
                    Else
                        Try
                            colArray(intCol) = CDate(colArray(intCol)).ToString("yyyy/MM/dd")
                            If CDate(colArray(intCol)) < "1900/1/1" Or CDate(colArray(intCol)) > "2100/1/1" Then Reason += strCol1 & " 範圍有誤<BR>"
                        Catch ex As Exception
                            Reason += strCol1 & " 必須是西元年格式(yyyy/MM/dd)<BR>"
                        End Try
                    End If
                End If
            Next
            '42~46
            For intCol As Integer = cst_i專長一 To cst_i專長五
                Select Case intCol
                    Case cst_i專長一
                        strCol1 = "專長一"
                    Case cst_i專長二
                        strCol1 = "專長二"
                    Case cst_i專長三
                        strCol1 = "專長三"
                    Case cst_i專長四
                        strCol1 = "專長四"
                    Case cst_i專長五
                        strCol1 = "專長五"
                End Select
                If colArray(intCol).ToString <> "" Then
                    colArray(intCol) = Trim(colArray(intCol))
                    If colArray(intCol).ToString.Length > 250 Then Reason += strCol1 & " 長度過長，限制為250個字元<BR>"
                End If
            Next
            Dim i_MAX_len As Integer = 300
            i_MAX_len = 100
            If colArray(cst_i譯著).ToString <> "" Then
                colArray(cst_i譯著) = Trim(colArray(cst_i譯著))
                If colArray(cst_i譯著).ToString.Length > i_MAX_len Then Reason += "譯著 長度過長，限制為" & i_MAX_len & "個字元<BR>"
            End If
            i_MAX_len = 200 'PROLICENSE1
            If colArray(cst_i專業證照政府).ToString <> "" Then
                colArray(cst_i專業證照政府) = Trim(colArray(cst_i專業證照政府))
                If colArray(cst_i專業證照政府).ToString.Length > i_MAX_len Then Reason += "專業證照-政府機關辦理相關證照或檢定 長度過長，限制為" & i_MAX_len & "個字元<BR>"
            End If
            i_MAX_len = 200
            If colArray(cst_i專業證照其他).ToString <> "" Then
                colArray(cst_i專業證照其他) = Trim(colArray(cst_i專業證照其他))
                If colArray(cst_i專業證照其他).ToString.Length > i_MAX_len Then Reason += "專業證照-其他證照或檢定 長度過長，限制為" & i_MAX_len & "個字元<BR>"
            End If
            If colArray(cst_i排課使用).ToString = "" Then
                Reason += "排課使用必須填寫<BR>"
            Else
                Select Case colArray(cst_i排課使用).ToString
                    Case "1", "2"
                    Case Else
                        Reason += "排課使用必須輸入1或2(1.是,2.否)<BR>"
                End Select
            End If
            colArray(cst_i講師類別) = TIMS.ClearSQM(Convert.ToString(colArray(cst_i講師類別)))
            If colArray(cst_i講師類別).ToString <> "" Then
                colArray(cst_i講師類別) = UCase(colArray(cst_i講師類別))
                Select Case colArray(cst_i講師類別).ToString
                    Case "Y"
                    Case Else
                        Reason += "講師類別 只能為(Y:是講師)或不填<BR>"
                End Select
            End If
            colArray(cst_i助教類別) = TIMS.ClearSQM(Convert.ToString(colArray(cst_i助教類別)))
            If colArray(cst_i助教類別).ToString <> "" Then
                colArray(cst_i助教類別) = UCase(colArray(cst_i助教類別))
                Select Case colArray(cst_i助教類別).ToString
                    Case "Y"
                    Case Else
                        Reason += "助教類別 只能為(Y:是助教)或不填<BR>"
                End Select
            End If

        End If
        Return Reason
    End Function

    'SQL 查詢
    Function Get_sSearch1() As String
        Dim sql As String = ""
        'Dim SearchStr As String = "" & vbCrLf
        Dim vsKindID As String = ""
        Dim vsWorkStatus As String = ""
        Dim vsIVID As String = ""
        Dim vsKindEngage As String = ""
        'Dim vsjobValue As String = ""
        Dim vsTMID As String = ""
        Dim vsRID As String = ""
        Dim vsTeachCName As String = ""
        Dim vsIDNO As String = ""
        Dim vsTeacherID As String = ""
        Dim vsTechType1 As String = ""
        Dim vsTechType2 As String = ""

        If tr_techtype12.Visible Then
            '產投(顯示) 才存取此功能
            'tr_techtype12.Visible = True '顯示 '講師 助教 (類別) 
            vsTechType1 = If(cb_techtype1.Checked, "Y", "")
            vsTechType2 = If(cb_techtype2.Checked, "Y", "")
        End If

        Dim v_DropDownList1 As String = TIMS.GetListValue(DropDownList1)
        Dim v_DropDownList2 As String = TIMS.GetListValue(DropDownList2)
        Dim v_DropDownList3 As String = TIMS.GetListValue(DropDownList3)
        Dim v_DropDownList4 As String = TIMS.GetListValue(DropDownList4)
        If DropDownList1.SelectedIndex <> 0 AndAlso v_DropDownList1 <> "" Then vsKindID = TIMS.ClearSQM(v_DropDownList1)
        If DropDownList2.SelectedIndex <> 0 AndAlso v_DropDownList2 <> "" Then vsWorkStatus = TIMS.ClearSQM(v_DropDownList2)
        If DropDownList3.SelectedIndex <> 0 AndAlso v_DropDownList3 <> "" Then vsIVID = TIMS.ClearSQM(v_DropDownList3)
        If DropDownList4.SelectedIndex <> 0 AndAlso v_DropDownList4 <> "" Then vsKindEngage = TIMS.ClearSQM(v_DropDownList4)
        'vsjobValue = TIMS.ClearSQM(jobValue.Value)
        vsTMID = TIMS.ClearSQM(trainValue.Value)
        If vsTMID = "" Then vsTMID = TIMS.ClearSQM(jobValue.Value) '試著取得-jobValue
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        vsRID = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        vsTeachCName = TIMS.ClearSQM(TextBox2.Text)
        vsIDNO = TIMS.ChangeIDNO(TIMS.ClearSQM(TextBox3.Text))
        vsTeacherID = TIMS.ClearSQM(TextBox4.Text)

        sql = ""
        sql &= " SELECT *" & vbCrLf
        sql &= " FROM dbo.TEACH_TEACHERINFO" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        '--------------- 查詢開始 '--------------- 
        If vsKindID <> "" Then sql &= " AND KindID = " & vsKindID & vbCrLf
        If vsWorkStatus <> "" Then sql &= " AND WorkStatus = '" & vsWorkStatus & "'" & vbCrLf
        If vsIVID <> "" Then sql &= " AND IVID = '" & vsIVID & "'" & vbCrLf
        If vsKindEngage <> "" Then sql &= " AND KindEngage = " & vsKindEngage & vbCrLf
        If vsTMID <> "" Then sql &= " AND TMID = " & vsTMID & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投
            'tr_techtype12.Visible = True '顯示 '講師 助教 (類別) 
            'If vsTechType1 <> "" Then sql &= " AND TechType1 = '" & vsTechType1 & "'" & vbCrLf  '講師
            If vsTechType1 <> "" And vsTechType2 = "" Then sql &= " AND (TechType1 = '" & vsTechType1 & "' OR KINDID IN (SELECT KINDID FROM ID_KINDOFTEACHER WHERE KINDNAME LIKE '%講師%'))" & vbCrLf  '講師，by:20181121
            'If vsTechType2 <> "" Then sql &= " AND TechType2 = '" & vsTechType2 & "'" & vbCrLf  '助教
            If vsTechType2 <> "" And vsTechType1 = "" Then sql &= " AND (TechType2 = '" & vsTechType2 & "' OR KINDID IN (SELECT KINDID FROM ID_KINDOFTEACHER WHERE KINDNAME LIKE '%助教%'))" & vbCrLf  '助教，by:20181121
        End If
        If vsRID <> "" Then sql &= " AND RID = '" & vsRID & "'" & vbCrLf
        If vsTeachCName <> "" Then sql &= " AND TeachCName LIKE '%" & vsTeachCName & "%'" & vbCrLf 'fix ORA-01722: invalid number
        If vsIDNO <> "" Then sql &= " AND IDNO LIKE '%" & vsIDNO & "%'" & vbCrLf 'fix ORA-01722: invalid number
        If vsTeacherID <> "" Then sql &= " AND TeacherID LIKE '%" & vsTeacherID & "%'" & vbCrLf
        '========== (依照承辦人需求，增加"主要職類關鍵字"欄位，by:20180912)
        Dim tCareerKeyWord As String = TIMS.ChangeIDNO(TIMS.ClearSQM(txtCareerKeyWord.Text))
        If tCareerKeyWord <> "" Then
            'sql &= " AND TMID IN (SELECT TMID FROM VIEW_TRAINTYPE WHERE TRAINNAME LIKE '%" + tCareerKeyWord + "%')" & vbCrLf
            sql &= " AND TMID IN (SELECT TMID FROM VIEW_TRAINTYPE WHERE JOBNAME LIKE '%" + tCareerKeyWord + "%' OR TRAINNAME LIKE '%" + tCareerKeyWord + "%')" & vbCrLf  '(依照承辦人需求,調整查詢條件，by:2080920)
        End If
        '===============================================================
        Return sql
    End Function

    '查詢原因-INQUIRY
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        center.Value = TIMS.ClearSQM(center.Value)
        TextBox2.Text = TIMS.ClearSQM(TextBox2.Text)
        TextBox3.Text = TIMS.ClearSQM(TextBox3.Text)
        TextBox4.Text = TIMS.ClearSQM(TextBox4.Text)
        TB_career_id.Text = TIMS.ClearSQM(TB_career_id.Text)
        Dim v_DropDownList1 As String = TIMS.GetListValue(DropDownList1)
        Dim v_DropDownList2 As String = TIMS.GetListValue(DropDownList2)
        Dim v_DropDownList3 As String = TIMS.GetListValue(DropDownList3)
        Dim v_DropDownList4 As String = TIMS.GetListValue(DropDownList4)

        If center.Value <> "" Then RstMemo &= String.Concat("&訓練機構=", center.Value)
        If TextBox2.Text <> "" Then RstMemo &= String.Concat("&講師姓名=", TextBox2.Text)
        If TextBox3.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", TextBox3.Text)
        If v_DropDownList4 <> "" Then RstMemo &= String.Concat("&內外聘=", v_DropDownList4)
        If v_DropDownList1 <> "" Then RstMemo &= String.Concat("&師資別=", v_DropDownList1)
        If TextBox4.Text <> "" Then RstMemo &= String.Concat("&講師代碼=", TextBox4.Text)
        If v_DropDownList2 <> "" Then RstMemo &= String.Concat("&排課使用=", v_DropDownList2)
        If TB_career_id.Text <> "" Then RstMemo &= String.Concat("&主要職類=", TB_career_id.Text)
        If v_DropDownList3 <> "" Then RstMemo &= String.Concat("&職稱=", v_DropDownList3)
        RstMemo &= String.Concat("&cb_techtype1=", cb_techtype1.Checked)
        RstMemo &= String.Concat("&cb_techtype2=", cb_techtype2.Checked)
        'RstMemo &= String.Concat("&cb_techtype3=", cb_techtype3.Checked)
        'RstMemo &= String.Concat("&cb_techtype4=", cb_techtype4.Checked)
        Return RstMemo
    End Function

    '查詢SQL
    Sub gClickSearchButton()

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim sql As String = Get_sSearch1()
        msg.Text = "查無資料!!"
        DataGridTable.Visible = False

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Dim dt As DataTable = Nothing
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, "查詢時發生錯誤，請重新輸入查詢值!!")
            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg += "/* ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            'Throw ex
            Exit Sub
        End Try

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode)
        Session(TIMS.gcst_rblWorkMode) = v_rblWorkMode 'rblWorkMode.SelectedValue
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "TEACHERID,TEACHCNAME,IDNO,KINDID,KINDENGAGE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, v_rblWorkMode, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then Return
        'If dt.Rows.Count > 0 Then End If

        '寫入Log查詢 SubInsAccountLog1 (Auth_Accountlog)
        'Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm查詢, v_rblWorkMode, "", "")
        For Each dr As DataRow In dt.Rows
            Dim idno As String = TIMS.ChangeIDNO(dr("IDNO").ToString())
            If v_rblWorkMode = TIMS.cst_wmdip1 Then dr("IDNO") = TIMS.strMask(idno, 1)
        Next

        msg.Text = ""
        DataGridTable.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "TechID"
        PageControler1.Sort = "TeacherID"
        PageControler1.ControlerLoad()
    End Sub

    '新增
    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox2.Text = TIMS.ClearSQM(TextBox2.Text)
        TextBox3.Text = TIMS.ClearSQM(TextBox3.Text)
        Call GetSearchStr()
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Dim sUrl1 As String = "TC_01_027_add.aspx?ID=" & rqMID
        sUrl1 &= "&proecess=Insert"
        sUrl1 &= "&TeachCName=" & TextBox2.Text
        sUrl1 &= "&TeachIDNO=" & TextBox3.Text
        TIMS.Utl_Redirect(Me, objconn, sUrl1)
    End Sub

    '列印排課匯入用的講師代碼
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, "RID=" & RIDValue.Value)
    End Sub

    Sub SImportFile2(ByRef FullFileName1 As String)
        File2.PostedFile.SaveAs(FullFileName1)  '上傳檔案

        Dim dt_xls As DataTable = Nothing
        Dim Reason As String = "" '儲存錯誤的原因
        '取得內容
        If (flag_File1_xls) Then
            dt_xls = TIMS.GetDataTable_XlsFile(FullFileName1, "", Reason, "計劃階層", "講師代碼", "講師姓名")
            If Reason <> "" Then
                Common.MessageBox(Me, "無法匯入!!" & Reason)
                Exit Sub
            End If
        End If
        If (flag_File1_ods) Then
            dt_xls = TIMS.GetDataTable_ODSFile(FullFileName1)
        End If
        '刪除檔案
        'IO.File.Delete(FullFileName1)
        TIMS.MyFileDelete(FullFileName1)
        Reason = TIMS.Chk_DTXLS1(dt_xls, flag_File1_xls, flag_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        Dim iRowIndex As Integer = 1

        'xls 方式 讀取寫入資料庫
        If dt_xls.Rows.Count > 0 Then '有資料
            '建立錯誤資料格式Table
            dtWrong.Columns.Add(New DataColumn("Index"))
            dtWrong.Columns.Add(New DataColumn("TeacherID"))
            dtWrong.Columns.Add(New DataColumn("Name"))
            dtWrong.Columns.Add(New DataColumn("IDNO"))
            dtWrong.Columns.Add(New DataColumn("Reason"))
            Reason = ""
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                If iRowIndex <> 0 Then
                    Dim colArray As Array = dt_xls.Rows(i).ItemArray
                    Reason = CheckImportData(colArray)

                    If Reason <> "" Then
                        '錯誤資料，填入錯誤資料表
                        Dim drWrong As DataRow = Nothing
                        drWrong = dtWrong.NewRow
                        dtWrong.Rows.Add(drWrong)
                        drWrong("Index") = iRowIndex
                        If colArray.Length > cst_i講師代碼 Then drWrong("TeacherID") = colArray(cst_i講師代碼)
                        If colArray.Length > cst_i講師姓名 Then drWrong("Name") = colArray(cst_i講師姓名)
                        If colArray.Length > cst_i身分證字號 Then drWrong("IDNO") = colArray(cst_i身分證字號)
                        drWrong("Reason") = Reason

                    Else '匯入資料
                        Dim sql As String = ""
                        Dim dr As DataRow = Nothing
                        Dim dt As DataTable = Nothing
                        Dim da As SqlDataAdapter = Nothing
                        Dim tConn As SqlConnection = DbAccess.GetConnection()
                        Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
                        Dim vRID As String = TIMS.ClearSQM(RIDValue.Value)
                        Dim vIDNO As String = TIMS.ClearSQM(colArray(cst_i身分證字號).ToString)
                        Dim vTeacherID As String = TIMS.ClearSQM(colArray(cst_i講師代碼).ToString)
                        Try
                            sql = ""
                            sql &= " SELECT * FROM TEACH_TEACHERINFO"
                            sql &= " WHERE RID = '" & vRID & "' AND IDNO = '" & vIDNO & "' AND TeacherID = '" & vTeacherID & "' "
                            dt = DbAccess.GetDataTable(sql, da, trans)
                            If dt.Rows.Count = 0 Then
                                Dim iTECHID As Integer = DbAccess.GetNewId(trans, "TEACH_TEACHERINFO_TECHID_SEQ,TEACH_TEACHERINFO,TECHID")
                                dr = dt.NewRow()
                                dt.Rows.Add(dr)
                                dr("TECHID") = iTECHID 'TEACH_TEACHERINFO_TECHID_SEQ
                            Else
                                dr = dt.Rows(0)
                            End If
                            dr("RID") = RIDValue.Value '機構
                            dr("TeacherID") = colArray(cst_i講師代碼).ToString '講師代碼
                            dr("TeachCName") = colArray(cst_i講師姓名).ToString '講師姓名
                            If colArray(cst_i講師英文姓名).ToString <> "" Then dr("TeachEName") = colArray(cst_i講師英文姓名).ToString '講師英文姓名
                            Select Case colArray(cst_i身份別).ToString
                                Case "1", "2"
                                    dr("PassPortNO") = colArray(cst_i身份別).ToString '身份別
                                Case Else
                                    dr("PassPortNO") = "2" '身份別
                            End Select
                            dr("IDNO") = colArray(cst_i身分證字號).ToString '身分證號碼
                            dr("Birthday") = If(colArray(cst_i出生日期).ToString <> "", colArray(cst_i出生日期).ToString, Convert.DBNull) '出生日期
                            Select Case Convert.ToString(colArray(cst_i性別))
                                Case "M", "F"
                                    dr("Sex") = colArray(cst_i性別).ToString '性別
                            End Select

                            'If colArray(cst_i主要職類).ToString = "" Then
                            '    Reason += "主要職類必須填寫<Br>"
                            'Else
                            '    ff = "TMID='" & colArray(cst_i主要職類) & "'"
                            '    If dtTRAINTYPE.Select(ff).Length = 0 Then
                            '        Reason += "主要職類不在鍵詞範圍內，請確認<BR>"
                            '    End If
                            'End If
                            ff3 = "GCODE2='" & colArray(cst_i主要職類) & "'"
                            Dim TMID_VAL As String = ""
                            If dtGOVCLASSCAST3.Select(ff3).Length > 0 Then
                                TMID_VAL = dtGOVCLASSCAST3.Select(ff3)(0)("TMID")
                                'dr("TMID") = TMID_VAL   '職類代碼
                                'Reason += "主要職類不在鍵詞範圍內，請確認<BR>"
                            End If
                            dr("TMID") = If(TMID_VAL <> "", TMID_VAL, Convert.DBNull) '職類代碼

                            If RIDValue.Value.Length = 1 Then
                                'SELECT distinct IVID FROM TEACH_TEACHERINFO where 1=1 order by 1
                                dr("IVID") = colArray(cst_i職稱).ToString '職稱代碼
                            Else
                                'SELECT distinct INVEST,trim(INVEST) FROM TEACH_TEACHERINFO where 1=1 and INVEST!=trim(INVEST) order by 1
                                'update TEACH_TEACHERINFO set INVEST=trim(INVEST) where 1=1 and INVEST!=trim(INVEST)
                                'SELECT distinct INVEST FROM TEACH_TEACHERINFO where 1=1 order by 1
                                dr("INVEST") = colArray(cst_i職稱).ToString '職稱代碼
                            End If
                            dr("KindEngage") = colArray(cst_i內外聘).ToString '內外聘
                            dr("KindID") = colArray(cst_i師資別).ToString '師資別
                            dr("DegreeID") = colArray(cst_i最高學歷).ToString '最高學歷
                            dr("GraduateStatus") = colArray(cst_i畢業狀況).ToString '畢業狀況
                            If colArray(cst_i學校名稱).ToString <> "" Then dr("SchoolName") = colArray(cst_i學校名稱).ToString '學校名稱
                            If colArray(cst_i科系名稱).ToString <> "" Then dr("Department") = colArray(cst_i科系名稱).ToString '科系名稱
                            dr("Phone") = colArray(cst_i聯絡電話).ToString '聯絡電話
                            If Convert.ToString(colArray(cst_i行動電話)) <> "" Then dr("Mobile") = colArray(cst_i行動電話).ToString '行動電話
                            If colArray(cst_i電子郵件).ToString <> "" Then dr("Email") = colArray(cst_i電子郵件).ToString 'E_Mail
                            dr("AddressZip") = colArray(cst_i郵遞區號前3碼).ToString '通訊地址Zip
                            dr("AddressZIP6W") = Val(colArray(cst_i郵遞區號6碼)) 'colArray(20).ToString '通訊地址Zip後2碼
                            dr("Address") = colArray(cst_i通訊地址).ToString '通訊地址

                            dr("WorkOrg") = colArray(cst_i服務單位名稱).ToString '服務單位名稱
                            If Convert.ToString(colArray(cst_i年資)) <> "" Then dr("ExpYears") = colArray(cst_i年資).ToString '服務年資
                            If colArray(cst_i服務部門).ToString <> "" Then dr("ServDept") = colArray(cst_i服務部門).ToString '服務部門
                            dr("WorkPhone") = colArray(cst_i服務單位電話).ToString '服務單位電話
                            If colArray(cst_i服務單位傳真).ToString <> "" Then dr("Fax") = colArray(cst_i服務單位傳真).ToString '服務單位傳真

                            If colArray(cst_i服務單位郵遞區號前3碼).ToString <> "" Then dr("WorkZip") = colArray(cst_i服務單位郵遞區號前3碼) '服務單位地址Zip
                            If colArray(cst_i服務單位郵遞區號6碼).ToString <> "" Then dr("WorkZIP6W") = Val(colArray(cst_i服務單位郵遞區號6碼)) '服務單位地址Zip後2碼
                            If colArray(cst_i服務單位地址).ToString <> "" Then dr("Workaddr") = colArray(cst_i服務單位地址).ToString '服務單位地址

                            If colArray(cst_i服務單位一).ToString <> "" Then dr("ExpUnit1") = colArray(cst_i服務單位一).ToString '服務單位一
                            If colArray(cst_i服務單位二).ToString <> "" Then dr("ExpUnit2") = colArray(cst_i服務單位二).ToString '服務單位二
                            If colArray(cst_i服務單位三).ToString <> "" Then dr("ExpUnit3") = colArray(cst_i服務單位三).ToString '服務單位三
                            If colArray(cst_i服務年資一).ToString <> "" Then
                                dr("ExpYears1") = colArray(cst_i服務年資一).ToString '服務年資一
                                dr("ExpMonths1") = 0
                            End If
                            If colArray(cst_i服務年資二).ToString <> "" Then
                                dr("ExpYears2") = colArray(cst_i服務年資二).ToString '服務年資二
                                dr("ExpMonths2") = 0
                            End If
                            If colArray(cst_i服務年資三).ToString <> "" Then
                                dr("ExpYears3") = colArray(cst_i服務年資三).ToString '服務年資三
                                dr("ExpMonths3") = 0
                            End If
                            If colArray(cst_i服務職稱一).ToString <> "" Then dr("INV1") = colArray(cst_i服務職稱一).ToString '服務職稱一
                            If colArray(cst_i服務職稱二).ToString <> "" Then dr("INV2") = colArray(cst_i服務職稱二).ToString '服務職稱二
                            If colArray(cst_i服務職稱三).ToString <> "" Then dr("INV3") = colArray(cst_i服務職稱三).ToString '服務職稱三
                            If colArray(cst_i服務期間一起日).ToString <> "" Then dr("ExpSDate1") = colArray(cst_i服務期間一起日).ToString '服務單位一起日
                            If colArray(cst_i服務期間一迄日).ToString <> "" Then dr("ExpEDate1") = colArray(cst_i服務期間一迄日).ToString '服務單位一迄日
                            If colArray(cst_i服務期間二起日).ToString <> "" Then dr("ExpSDate2") = colArray(cst_i服務期間二起日).ToString '服務單位二起日
                            If colArray(cst_i服務期間二迄日).ToString <> "" Then dr("ExpEDate2") = colArray(cst_i服務期間二迄日).ToString '服務單位二迄日
                            If colArray(cst_i服務期間三起日).ToString <> "" Then dr("ExpSDate3") = colArray(cst_i服務期間三起日).ToString '服務單位三起日
                            If colArray(cst_i服務期間三迄日).ToString <> "" Then dr("ExpEDate3") = colArray(cst_i服務期間三迄日).ToString '服務單位三迄日
                            Dim xi As Integer = 0
                            xi = 0
                            For ji As Integer = cst_i專長一 To cst_i專長五
                                xi += 1
                                Dim columnName As String = "Specialty" & CStr(xi)
                                dr(columnName) = colArray(ji).ToString '專長一~專長五(42~46)
                            Next
                            If colArray(cst_i譯著).ToString <> "" Then dr("TransBook") = colArray(cst_i譯著).ToString '譯著
                            'If colArray(cst_i專業證照).ToString <> "" Then dr("ProLicense") = colArray(cst_i專業證照).ToString '專業證照
                            dr("ProLicense1") = Convert.DBNull
                            If colArray(cst_i專業證照政府).ToString <> "" Then dr("ProLicense1") = colArray(cst_i專業證照政府).ToString '專業證照-政府
                            dr("ProLicense2") = Convert.DBNull
                            If colArray(cst_i專業證照其他).ToString <> "" Then dr("ProLicense2") = colArray(cst_i專業證照其他).ToString '專業證照-其他
                            If colArray(cst_i排課使用).ToString <> "" Then dr("WorkStatus") = colArray(cst_i排課使用).ToString '排課使用
                            If colArray(cst_i講師類別).ToString <> "" Then dr("TECHTYPE1") = colArray(cst_i講師類別).ToString '講師類別
                            If colArray(cst_i助教類別).ToString <> "" Then dr("TECHTYPE2") = colArray(cst_i助教類別).ToString '助教類別
                            dr("ModifyAcct") = sm.UserInfo.UserID '異動者
                            dr("ModifyDate") = Now() '異動時間

                            DbAccess.UpdateDataTable(dt, da, trans)
                            DbAccess.CommitTrans(trans)
                        Catch ex As Exception
                            DbAccess.RollbackTrans(trans)
                            TIMS.CloseDbConn(tConn)
                            Const cst_errmsg1 As String = "意外錯誤：(請提供詳細資料，並連絡系統管理者協助處理)"
                            Dim strErrmsg As String = ""
                            strErrmsg += "/*  匯入名冊 TC_01_027. Private Sub Btn_XlsImport_Click(ByVal sender As Object */" & vbCrLf
                            strErrmsg += "/*  ex.ToString: */" & vbCrLf
                            strErrmsg += ex.ToString & vbCrLf
                            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                            Call TIMS.WriteTraceLog(strErrmsg)
                            Common.MessageBox(Me, cst_errmsg1)
                            Exit Sub
                            'Throw 'ex
                        End Try
                        Call TIMS.CloseDbConn(tConn)
                    End If
                End If
                iRowIndex += 1
            Next
        End If

        '判斷匯出資料是否有誤
        Dim explain, explain2 As String
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
        explain2 = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"
        If dtWrong.Rows.Count = 0 Then
            If Reason = "" Then
                Common.MessageBox(Me, explain)
            Else
                Common.MessageBox(Me, explain & Reason)
            End If
        Else
            'Session("MyWrongTable") = dtWrong
            Datagrid2.Style.Item("display") = "inline"
            Datagrid2.Visible = True
            Datagrid2.DataSource = dtWrong
            Datagrid2.DataBind()
            Common.MessageBox(Me, "匯入動作完成,但有錯誤資料(無法匯入)請檢示原因!!!")
            For i As Integer = 1 To 100
                If i = 100 Then eMeng.Style.Item("display") = "inline"
                'Page.RegisterStartupScript("", "<script>{window.document.getElementById('eMeng').style.visibility='visible';}</script>")
            Next
            'Page.RegisterStartupScript("", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('TC_01_007_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If
        'Button1_Click(sender, e)
        Call gClickSearchButton()
    End Sub

    '匯入名冊
    Private Sub Btn_XlsImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btn_XlsImport.Click
        '../../Doc/ClassTeacher_v18.zip
        'Dim flag_File1_xls As Boolean = False
        'Dim flag_File1_ods As Boolean = False
        Dim sMyFileName As String = ""
        Dim sErrMsg As String = TIMS.ChkFile1(File2, sMyFileName, flag_File1_xls, flag_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, "xls") Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, "ods") Then Return
        End If

        Const Cst_FileSavePath As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File2.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        Call SImportFile2(FullFileName1)
    End Sub

#Region "匯出教師資料 EXCEL檔"
#End Region

#Region "NO USE"
    '20080603  Andy 新增匯出教師資料 '匯出名冊
    'Sub Export1()
    '    'Dim sql As String
    '    'Dim dt As DataTable
    '    'Dim dr As DataRow
    '    'copy一份sample資料---------------------   Start
    '    'Dim MyFile As System.IO.File
    '    'Dim MyDownload As System.IO.File
    '    Dim strErrmsg As String = ""

    '    center.Value = TIMS.ClearSQM(center.Value)
    '    RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
    '    If center.Value = "" Then
    '        Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
    '        Exit Sub
    '    End If
    '    If RIDValue.Value = "" Then
    '        Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
    '        Exit Sub
    '    End If
    '    'Dim ExpTitle As String = center.Value.ToString() & Format(Date.Now(), "yyyy-M-d")
    '    'Dim ExpTitle As String = center.Value.ToString() & Format(Date.Now(), "yyyy-M")
    '    Dim ExpTitle As String = TIMS.ChangeIDNO(Replace(Replace(Replace(center.Value, ")", ""), "(", ""), "/", ""))
    '    Dim MyPath As String = ""
    '    Dim xlsExtNM1 As String = ".xls"
    '    Dim sFileName As String = String.Concat("~\TC\01\Temp\", ExpTitle, TIMS.GetDateNo(), xlsExtNM1)

    '    MyPath = Server.MapPath(sFileName)
    '    Dim MyFileName As String = String.Concat(ExpTitle, xlsExtNM1)
    '    Const cst_Sample1xls As String = "~\TC\01\Temp\Sample22.xls" ', xlsExtNM1
    '    If Not IO.File.Exists(Server.MapPath(cst_Sample1xls)) Then
    '        Common.MessageBox(Me, "Sample檔案不存在")
    '        Exit Sub
    '    End If
    '    Try
    '        IO.File.Copy(Server.MapPath(cst_Sample1xls), MyPath, True)
    '        '除去sample檔的唯讀屬性
    '        IO.File.SetAttributes(MyPath, IO.FileAttributes.Normal)
    '    Catch ex As Exception
    '        strErrmsg = ""
    '        strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
    '        strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
    '        strErrmsg += ex.ToString & vbCrLf
    '        Common.MessageBox(Me, strErrmsg)
    '        'Exit Sub
    '    End Try
    '    'copy一份sample資料---------------------   End

    '    Dim dt As DataTable = Nothing
    '    Dim sql As String = Get_sSearch1()
    '    dt = DbAccess.GetDataTable(sql, objconn)
    '    If dt.Rows.Count = 0 Then
    '        Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
    '        Exit Sub
    '    End If

    '    '寫入Log查詢 SubInsAccountLog1 (Auth_Accountlog)
    '    Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) '
    '    Dim MRqID As String = TIMS.Get_MRqID(Me)
    '    Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm匯出, v_rblWorkMode, "", "")

    '    Using MyConn As New OleDb.OleDbConnection
    '        MyConn.ConnectionString = TIMS.Get_OleDbStr(MyPath)
    '        Try
    '            MyConn.Open()
    '        Catch ex As Exception
    '            'Dim strErrmsg As String = ""
    '            strErrmsg &= "/* ex.ToString: */" & vbCrLf & ex.ToString & vbCrLf
    '            strErrmsg &= "sql:" & vbCrLf & sql & vbCrLf
    '            strErrmsg &= "conn.ConnectionString:" & MyConn.ConnectionString & vbCrLf
    '            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
    '            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '            Call TIMS.WriteTraceLog(strErrmsg)
    '            'Common.MessageBox(Me, "Excel資料無法開啟連線!")
    '            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
    '            Return 'Exit Sub
    '        End Try

    '        dt.DefaultView.Sort = "TechID"
    '        'Dim ModifyAcct As String ="" '異動者
    '        'Dim ModifyDate As String ="" '異動時間
    '        For Each dr As DataRow In dt.Rows
    '            'Dim RID As String '機構
    '            Dim RIDLevel As String = "" '計畫階層
    '            Dim TeacherID As String = "" '講師代碼
    '            Dim TeachCName As String = "" '講師姓名
    '            Dim TeachEName As String = "" '講師英文姓名
    '            Dim PassPortNO As String = "" '身份別
    '            Dim IDNO As String = "" '身分證號碼
    '            Dim Birthday As String = "" '出生日期
    '            Dim Sex As String = "" '性別
    '            'Dim TMID As String = "" '職類代碼
    '            Dim GCODE2 As String = "" '職類代碼-GCODE2

    '            Dim IVID As String = "" '職稱代碼
    '            Dim KindEngage As String = "" '內外聘
    '            Dim KindID As String = "" '師資別
    '            Dim DegreeID As String = "" '學歷
    '            Dim GraduateStatus As String = "" '畢業狀況
    '            Dim SchoolName As String = "" '學校名稱
    '            Dim Department As String = "" '科系名稱
    '            Dim Phone As String = "" '聯絡電話
    '            Dim Mobile As String = "" '行動電話
    '            Dim Email As String = "" 'E_Mail
    '            Dim AddressZip As String = "" '戶藉地址Zip
    '            Dim AddressZIP6W As String = "" '戶藉地址Zip後2碼
    '            Dim Address As String = "" '戶藉地址
    '            Dim WorkOrg As String = "" '服務單位名稱
    '            Dim ExpYears As String = "" '服務年資
    '            Dim ServDept As String = "" '服務部門
    '            Dim WorkPhone As String = "" '服務單位電話
    '            Dim Fax As String = "" '服務單位傳真
    '            Dim WorkZip As String = "" '服務單位地址Zip
    '            Dim WorkZIP6W As String = "" '服務單位地址Zip後2碼
    '            Dim Workaddr As String = "" ' 服務單位地址
    '            Dim ExpUnit1 As String = "" '服務單位一
    '            Dim ExpUnit2 As String = "" '服務單位二
    '            Dim ExpUnit3 As String = "" '服務單位三
    '            Dim ExpYears1 As String = "" '服務年資一
    '            Dim ExpYears2 As String = "" '服務年資二
    '            Dim ExpYears3 As String = "" '服務年資三
    '            Dim EpINV1 As String = "" '服務職稱1
    '            Dim EpINV2 As String = "" '服務職稱2
    '            Dim EpINV3 As String = "" '服務職稱3
    '            Dim ExpSDate1 As String = "" '服務單位一起日
    '            Dim ExpEDate1 As String = "" '服務單位一迄日
    '            Dim ExpSDate2 As String = "" '服務單位二起日
    '            Dim ExpEDate2 As String = "" '服務單位二迄日
    '            Dim ExpSDate3 As String = "" '服務單位三起日
    '            Dim ExpEDate3 As String = "" '服務單位三迄日
    '            Dim Specialty1 As String = "" '專長一
    '            Dim Specialty2 As String = "" '專長二
    '            Dim Specialty3 As String = "" '專長三
    '            Dim Specialty4 As String = "" '專長四
    '            Dim Specialty5 As String = "" '專長五
    '            Dim Specialty1b As String = "" '專長一b
    '            Dim Specialty2b As String = "" '專長二b
    '            Dim Specialty3b As String = "" '專長三b
    '            Dim Specialty4b As String = "" '專長四b
    '            Dim Specialty5b As String = "" '專長五b

    '            Dim TransBook As String = "" '譯著

    '            'Dim ProLicense As String ="" '專業證照
    '            Dim ProLicense1 As String = "" '專業證照(政府)
    '            Dim ProLicense2 As String = "" '專業證照(其他)
    '            Dim WorkStatus As String = "" '任職狀況
    '            Dim TechType1 As String = "" '講師類別
    '            Dim TechType2 As String = "" '助教類別
    '            'RID = Right(dr("RID").ToString, 2)
    '            'If center.Value <> "" Then center.Value = Trim(center.Value)
    '            RIDLevel = TIMS.ClearSQM(center.Value)
    '            TeacherID = TIMS.ClearSQM(dr("TeacherID"))
    '            TeachCName = TIMS.ClearSQM(dr("TeachCName"))
    '            TeachEName = TIMS.ClearSQM(dr("TeachEName"))
    '            Select Case TIMS.ClearSQM(dr("PassPortNO"))
    '                Case "1", "2"
    '                    PassPortNO = dr("PassPortNO").ToString
    '                Case Else
    '                    PassPortNO = "2"
    '            End Select
    '            IDNO = TIMS.ClearSQM(dr("IDNO"))
    '            Dim flag_idno_ok As Boolean = False
    '            If IDNO <> "" AndAlso TIMS.CheckIDNO(IDNO) Then flag_idno_ok = True
    '            If flag_idno_ok Then
    '                Sex = If(IDNO.Chars(1) = "1", "M", If(IDNO.Chars(1) = "2", "F", ""))
    '            End If
    '            'Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) '
    '            If v_rblWorkMode = "1" Then IDNO = TIMS.strMask(IDNO, 1)
    '            Birthday = ""
    '            If Convert.ToString(dr("Birthday")) <> "" Then
    '                Birthday = TIMS.cdate3(dr("Birthday"))
    '                If v_rblWorkMode = "1" Then Birthday = TIMS.strMask(Birthday, 2)
    '            End If
    '            'Birthday = dr("Birthday").ToString
    '            If Convert.ToString(dr("Sex")) <> "" Then Sex = dr("Sex").ToString

    '            'TMID-OLD
    '            'Dim v_TB_career_id_txt As String = ""
    '            'If Convert.ToString(dr("TrainID")) <> "" Then
    '            '    '職類
    '            '    v_TB_career_id_txt = "[" & Convert.ToString(dr("TrainID")) & "]" & Convert.ToString(dr("TrainName"))
    '            '    TB_career_id.Text = v_TB_career_id_txt 'TMID-OLD
    '            'End If
    '            'If v_TB_career_id_txt = "" AndAlso Convert.ToString(dr("JobID")) <> "" Then
    '            '    '若取不到職類-但有業別-顯示業別
    '            '    v_TB_career_id_txt = "[" & Convert.ToString(dr("JobID")) & "]" & Convert.ToString(dr("JobName"))
    '            '    TB_career_id.Text = v_TB_career_id_txt 'TMID-OLD
    '            'End If
    '            'GCODE2-NEW
    '            'Dim v_GCODE2 As String = ""
    '            'ff3 = "TMID='" & Convert.ToString(dr("TMID")) & "'"
    '            'If dtGOVCLASSCAST3.Select(ff3).Length > 0 Then v_GCODE2 = "[" & dtGOVCLASSCAST3.Select(ff3)(0)("GCODE2") & "]" & dtGOVCLASSCAST3.Select(ff3)(0)("CNAME")
    '            'If v_GCODE2 <> "" Then TB_career_id.Text = v_GCODE2 'v_TB_career_id_txt
    '            ff3 = "TMID='" & Convert.ToString(dr("TMID")) & "'"
    '            If dtGOVCLASSCAST3.Select(ff3).Length > 0 Then GCODE2 = dtGOVCLASSCAST3.Select(ff3)(0)("GCODE2")

    '            IVID = TIMS.ClearSQM(dr("Invest"))
    '            If RIDValue.Value.Length = 1 Then IVID = TIMS.ClearSQM(dr("IVID"))
    '            KindEngage = dr("KindEngage").ToString
    '            KindID = dr("KindID").ToString
    '            DegreeID = dr("DegreeID").ToString
    '            GraduateStatus = dr("GraduateStatus").ToString
    '            SchoolName = TIMS.ClearSQM(dr("SchoolName"))
    '            Department = TIMS.ClearSQM(dr("Department"))
    '            Phone = TIMS.ClearSQM(dr("Phone"))
    '            Mobile = TIMS.ClearSQM(dr("Mobile"))
    '            Email = TIMS.ClearSQM(dr("Email"))

    '            AddressZip = TIMS.ClearSQM(dr("AddressZip"))
    '            AddressZIP6W = TIMS.ClearSQM(dr("AddressZIP6W"))
    '            Address = TIMS.ClearSQM(dr("Address"))

    '            WorkOrg = TIMS.ClearSQM(dr("WorkOrg"))
    '            ExpYears = TIMS.ClearSQM(dr("ExpYears"))
    '            ServDept = TIMS.ClearSQM(dr("ServDept"))
    '            WorkPhone = TIMS.ClearSQM(dr("WorkPhone"))
    '            Fax = TIMS.ClearSQM(dr("Fax"))
    '            WorkZip = TIMS.ClearSQM(dr("WorkZip"))
    '            WorkZIP6W = TIMS.ClearSQM(dr("WorkZIP6W"))
    '            Workaddr = TIMS.ClearSQM(dr("Workaddr"))
    '            ExpUnit1 = TIMS.ClearSQM(dr("ExpUnit1"))
    '            ExpUnit2 = TIMS.ClearSQM(dr("ExpUnit2"))
    '            ExpUnit3 = TIMS.ClearSQM(dr("ExpUnit3"))
    '            ExpYears1 = TIMS.ClearSQM(dr("ExpYears1"))
    '            ExpYears2 = TIMS.ClearSQM(dr("ExpYears2"))
    '            ExpYears3 = TIMS.ClearSQM(dr("ExpYears3"))
    '            EpINV1 = TIMS.ClearSQM(dr("INV1"))
    '            EpINV2 = TIMS.ClearSQM(dr("INV2"))
    '            EpINV3 = TIMS.ClearSQM(dr("INV3"))
    '            ExpSDate1 = If(Convert.ToString(dr("ExpSDate1")) <> "", TIMS.cdate3(dr("ExpSDate1")), "")
    '            ExpSDate2 = If(Convert.ToString(dr("ExpSDate2")) <> "", TIMS.cdate3(dr("ExpSDate2")), "")
    '            ExpSDate3 = If(Convert.ToString(dr("ExpSDate3")) <> "", TIMS.cdate3(dr("ExpSDate3")), "")
    '            ExpEDate1 = If(Convert.ToString(dr("ExpEDate1")) <> "", TIMS.cdate3(dr("ExpEDate1")), "")
    '            ExpEDate2 = If(Convert.ToString(dr("ExpEDate2")) <> "", TIMS.cdate3(dr("ExpEDate2")), "")
    '            ExpEDate3 = If(Convert.ToString(dr("ExpEDate3")) <> "", TIMS.cdate3(dr("ExpEDate3")), "")

    '            'ExpSDate1 = dr("ExpSDate1").ToString
    '            'ExpSDate2 = dr("ExpSDate2").ToString
    '            'ExpSDate3 = dr("ExpSDate3").ToString
    '            'ExpEDate1 = dr("ExpEDate1").ToString
    '            'ExpEDate2 = dr("ExpEDate2").ToString
    '            'ExpEDate3 = dr("ExpEDate3").ToString
    '            Specialty1 = TIMS.ChangeSQM(dr("Specialty1")) '專長一
    '            Specialty2 = TIMS.ChangeSQM(dr("Specialty2")) '專長二
    '            Specialty3 = TIMS.ChangeSQM(dr("Specialty3")) '專長三
    '            Specialty4 = TIMS.ChangeSQM(dr("Specialty4")) '專長四
    '            Specialty5 = TIMS.ChangeSQM(dr("Specialty5")) '專長五
    '            'Call SplitSTR250(Specialty1, Specialty1b)
    '            'Call SplitSTR250(Specialty2, Specialty2b)
    '            'Call SplitSTR250(Specialty3, Specialty3b)
    '            'Call SplitSTR250(Specialty4, Specialty4b)
    '            'Call SplitSTR250(Specialty5, Specialty5b)

    '            TransBook = TIMS.ChangeSQM(dr("TransBook"))   '譯著
    '            'ProLicense = TIMS.ChangeSQM(dr("ProLicense")) '專業證照
    '            ProLicense1 = TIMS.ChangeSQM(dr("ProLicense1")) '專業證照(政府)
    '            ProLicense2 = TIMS.ChangeSQM(dr("ProLicense2")) '專業證照(其他)
    '            WorkStatus = TIMS.ChangeSQM(dr("WorkStatus")) '任職狀況
    '            TechType1 = TIMS.ChangeSQM(dr("TechType1")) '講師類別
    '            TechType2 = TIMS.ChangeSQM(dr("TechType2")) '助教類別
    '            'ModifyAcct = dr("ModifyAcct").ToString
    '            'ModifyDate = dr("Specialty1").ToString
    '            'PassPortNO = dr("PassPortNO").ToString

    '            sql = "INSERT INTO [Sheet1$] ("
    '            sql &= "計劃階層,講師代碼,講師姓名,講師英文姓名,身份別,身分證字號,出生日期,性別,主要職類,職稱"
    '            sql &= ",內外聘,師資別,最高學歷,畢業狀況,學校名稱,科系名稱,聯絡電話,行動電話,電子郵件"
    '            sql &= ",郵遞區號前3碼,郵遞區號6碼,戶籍地址,服務單位名稱,年資,服務部門,服務單位電話,服務單位傳真"
    '            sql &= ",服務單位郵遞區號前3碼,服務單位郵遞區號6碼,服務單位地址,服務單位一,服務單位二,服務單位三"
    '            sql &= ",服務年資一,服務年資二,服務年資三"
    '            sql &= ",服務職稱一,服務職稱二,服務職稱三"
    '            sql &= ",服務期間一起日,服務期間一迄日"
    '            sql &= ",服務期間二起日,服務期間二迄日"
    '            sql &= ",服務期間三起日,服務期間三迄日"
    '            sql &= ",專長一,專長二,專長三,專長四,專長五,譯著"
    '            sql &= ",[專業證照(政府)],[專業證照(其他)],排課使用,講師類別,助教類別"
    '            sql &= ") VALUES ("
    '            sql &= "'" & RIDLevel & "','" & TeacherID & "','" & TeachCName & "','" & TeachEName & "','" & PassPortNO & "','" & IDNO & "','" & Birthday & "','" & Sex & "','" & GCODE2 & "','" & IVID & "'"
    '            sql &= ",'" & KindEngage & "','" & KindID & "','" & DegreeID & "','" & GraduateStatus & "','" & SchoolName & "','" & Department & "','" & Phone & "','" & Mobile & "','" & Email & "'"
    '            sql &= ",'" & AddressZip & "','" & AddressZIP6W & "','" & Address & "','" & WorkOrg & "','" & ExpYears & "','" & ServDept & "','" & WorkPhone & "','" & Fax & "'"
    '            sql &= ",'" & WorkZip & "','" & WorkZIP6W & "','" & Workaddr & "','" & ExpUnit1 & "','" & ExpUnit2 & "','" & ExpUnit3 & "'"
    '            sql &= ",'" & ExpYears1 & "','" & ExpYears2 & "','" & ExpYears3 & "'"
    '            sql &= ",'" & EpINV1 & "','" & EpINV2 & "','" & EpINV3 & "'"
    '            sql &= ",'" & ExpSDate1 & "','" & ExpEDate1 & "'"
    '            sql &= ",'" & ExpSDate2 & "','" & ExpEDate2 & "'"
    '            sql &= ",'" & ExpSDate3 & "','" & ExpEDate3 & "'"
    '            sql &= ",'" & Specialty1 & "','" & Specialty2 & "','" & Specialty3 & "','" & Specialty4 & "','" & Specialty5 & "','" & TransBook & "'"
    '            sql &= ",'" & ProLicense1 & "','" & ProLicense2 & "','" & WorkStatus & "','" & TechType1 & "','" & TechType2 & "'" & ")"

    '            Using ole_Cmd As New OleDb.OleDbCommand(sql, MyConn)
    '                'update TEACH_TEACHERINFO Set SPECIALTY1=concat(SPECIALTY1,'') WHERE 講師代碼=techid
    '                'cmd = New OleDb.OleDbCommand(sql, ole_conn)
    '                'Dim sParms As New Hashtable
    '                Try
    '                    If MyConn.State = ConnectionState.Closed Then MyConn.Open()
    '                    ole_Cmd.ExecuteNonQuery()
    '                    'Call UPDATE_LongText(MyConn, Specialty1b, TeacherID, "專長一", "講師代碼")
    '                    'Call UPDATE_LongText(MyConn, Specialty2b, TeacherID, "專長二", "講師代碼")
    '                    'Call UPDATE_LongText(MyConn, Specialty3b, TeacherID, "專長三", "講師代碼")
    '                    'Call UPDATE_LongText(MyConn, Specialty4b, TeacherID, "專長四", "講師代碼")
    '                    'Call UPDATE_LongText(MyConn, Specialty5b, TeacherID, "專長五", "講師代碼")

    '                    'If conn.State = ConnectionState.Open Then conn.Close()
    '                Catch ex As Exception
    '                    'Dim strErrmsg As String = ""
    '                    strErrmsg &= "/* ex.ToString: */" & vbCrLf & ex.ToString & vbCrLf
    '                    strErrmsg &= "sql:" & vbCrLf & sql & vbCrLf
    '                    strErrmsg &= "conn.ConnectionString:" & MyConn.ConnectionString & vbCrLf
    '                    strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
    '                    'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '                    Call TIMS.WriteTraceLog(strErrmsg)
    '                    If MyConn.State = ConnectionState.Open Then MyConn.Close()

    '                    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
    '                    Return 'Exit Sub
    '                    Exit For
    '                    'Throw ex
    '                End Try
    '            End Using

    '        Next
    '        If MyConn.State = ConnectionState.Open Then MyConn.Close()
    '        '根據路徑建立資料庫連線，並取出學員資料填入---------------   End
    '    End Using

    '    Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
    '    Select Case V_ExpType
    '        Case "EXCEL"
    '            ExpExccl_1(strErrmsg, MyPath)

    '        Case "ODS"
    '            Dim fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
    '            Dim br As New System.IO.BinaryReader(fr)
    '            Dim buf(fr.Length) As Byte
    '            fr.Read(buf, 0, fr.Length)
    '            fr.Close()

    '            Dim sFileName1 As String = "ExpFile" & TIMS.GetRnd6Eng()
    '            Dim parmsExp As New Hashtable
    '            parmsExp.Add("ExpType", V_ExpType) 'EXCEL/PDF/ODS
    '            parmsExp.Add("FileName", sFileName1)
    '            parmsExp.Add("xlsx_buf", buf)
    '            'parmsExp.Add("strHTML", strHTML)
    '            parmsExp.Add("ResponseNoEnd", "Y")
    '            TIMS.Utl_ExportRp1(Me, parmsExp)

    '        Case Else
    '            Dim s_log1 As String = ""
    '            s_log1 = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
    '            Common.MessageBox(Me, s_log1)
    '            Exit Sub
    '    End Select

    '    Call TIMS.MyFileDelete(MyPath)
    '    TIMS.Utl_RespWriteEnd(Me, objconn, "")
    '    Exit Sub
    '    'If strErrmsg = "" Then
    '    'End If
    '    'If strErrmsg = "" Then Response.End()
    '    'If strErrmsg <> "" Then Common.MessageBox(Me, strErrmsg)
    '    '將新建立的excel存入記憶體下載-----   End
    'End Sub
#End Region

    Sub Export1XLS()
        'copy一份sample資料-
        Dim strErrmsg As String = ""

        center.Value = TIMS.ClearSQM(center.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If center.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        'Dim ExpTitle As String = center.Value.ToString() & Format(Date.Now(), "yyyy-M-d")
        'Dim ExpTitle As String = center.Value.ToString() & Format(Date.Now(), "yyyy-M")
        Dim s_WebPath1 As String = "~\TC\01\Temp\"
        Dim ExpTitle As String = TIMS.ChangeIDNO(Replace(Replace(Replace(center.Value, ")", ""), "(", ""), "/", ""))
        Dim sFileName As String = TIMS.GetValidFileName(String.Concat(ExpTitle, TIMS.GetDateNo(), ".xlsx"))
        Dim sFileNameP2 As String = String.Concat(s_WebPath1, sFileName)
        Dim MyPath As String = Server.MapPath(sFileNameP2)
        'Dim MyFileName As String = String.Concat(ExpTitle, xlsExtNM1)
        Const cst_Sample1xls As String = "~\TC\01\Temp\Sample23.xlsx" ', xlsExtNM1
        If Not IO.File.Exists(Server.MapPath(cst_Sample1xls)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If
        Try
            TIMS.MyFileDelete(MyPath)
            IO.File.Copy(Server.MapPath(cst_Sample1xls), MyPath, True)
            '除去sample檔的唯讀屬性
            IO.File.SetAttributes(MyPath, IO.FileAttributes.Normal)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg += $" (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉){vbCrLf}{ex.Message}{vbCrLf}"
            Common.MessageBox(Me, strErrmsg)
            Return 'Exit Sub
        End Try
        'copy一份sample資料---------------------   End

        Dim sql As String = Get_sSearch1()
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode)
        Session(TIMS.gcst_rblWorkMode) = v_rblWorkMode 'rblWorkMode.SelectedValue
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "TECHID,TEACHERID,TEACHCNAME,TEACHENAME,IDNO,BIRTHDAY,PHONE,MOBILE,EMAIL")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, v_rblWorkMode, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
        Select Case V_ExpType
            Case "EXCEL"
                Dim s_ExpFile As String = TIMS.GetValidFileName(TIMS.ClearSQM(String.Concat("ExpFile", TIMS.GetRnd6Eng, ".xlsx")))
                Dim myFileName2 As String = String.Concat(s_WebPath1, s_ExpFile) '複製
                Dim sMyFileMP2 As String = Server.MapPath(myFileName2)
                TIMS.MyFileDelete(sMyFileMP2)

                Using fs As IO.FileStream = New IO.FileStream(MyPath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.ReadWrite)
                    Using ep As New OfficeOpenXml.ExcelPackage(fs)
                        Dim sheet As OfficeOpenXml.ExcelWorksheet = ep.Workbook.Worksheets(0) '取得Sheet1
                        dt.DefaultView.Sort = "TechID"
                        Call XLSX_Export_OpenXml(dt, v_rblWorkMode, sheet)
                        Using createStream As New FileStream(sMyFileMP2, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
                            ep.SaveAs(createStream) '存檔
                        End Using
                    End Using
                End Using
                TIMS.ExpExcel_1(Me, strErrmsg, sMyFileMP2, s_ExpFile)
                TIMS.MyFileDelete(sMyFileMP2)

            Case "ODS"
                Dim s_ExpFile As String = TIMS.GetValidFileName(TIMS.ClearSQM(String.Concat("ExpFile", TIMS.GetRnd6Eng, ".xlsx")))
                Dim myFileName2 As String = String.Concat(s_WebPath1, s_ExpFile) '複製
                Dim sMyFileMP2 As String = Server.MapPath(myFileName2)
                TIMS.MyFileDelete(sMyFileMP2)

                Using fs As IO.FileStream = New IO.FileStream(MyPath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.ReadWrite)
                    Using ep As New OfficeOpenXml.ExcelPackage(fs)
                        Dim sheet As OfficeOpenXml.ExcelWorksheet = ep.Workbook.Worksheets(0) '取得Sheet1
                        dt.DefaultView.Sort = "TechID"
                        Call XLSX_Export_OpenXml(dt, v_rblWorkMode, sheet)
                        Using createStream As New FileStream(sMyFileMP2, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
                            ep.SaveAs(createStream) '存檔
                        End Using
                    End Using
                End Using

                Dim fr As New System.IO.FileStream(sMyFileMP2, IO.FileMode.Open)
                Dim br As New System.IO.BinaryReader(fr)
                Dim buf(fr.Length) As Byte
                fr.Read(buf, 0, fr.Length)
                fr.Close()

                'dt.DefaultView.Sort = "TechID" 'Dim strHTML As String = ODS_Export_HTML(dt)
                Dim sFileName1 As String = "ExpFile" & TIMS.GetRnd6Eng()
                Dim parmsExp As New Hashtable
                parmsExp.Add("ExpType", V_ExpType) 'EXCEL/PDF/ODS
                parmsExp.Add("FileName", sFileName1)
                parmsExp.Add("xlsx_buf", buf)
                'parmsExp.Add("strHTML", strHTML)
                parmsExp.Add("ResponseNoEnd", "Y")
                TIMS.Utl_ExportRp1(Me, parmsExp)
                TIMS.MyFileDelete(sMyFileMP2)

            Case Else
                Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                Common.MessageBox(Me, s_log1)
                Exit Sub

        End Select

        Call TIMS.MyFileDelete(MyPath)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        Exit Sub
        'If strErrmsg = "" Then 'End If
        'If strErrmsg = "" Then Response.End()
        'If strErrmsg <> "" Then Common.MessageBox(Me, strErrmsg)
        '將新建立的excel存入記憶體下載-----   End
    End Sub

    Sub XLSX_Export_OpenXml(dt As DataTable, v_rblWorkMode As String, ByRef sheet As OfficeOpenXml.ExcelWorksheet)
        'Dim fs As IO.FileStream = New IO.FileStream(MyPath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.ReadWrite)
        ''Dim fs As FileInfo = New FileInfo(sMyFile1)
        'Dim ep As OfficeOpenXml.ExcelPackage = New OfficeOpenXml.ExcelPackage(fs)
        'Dim sheet As OfficeOpenXml.ExcelWorksheet = ep.Workbook.Worksheets(0) '取得Sheet1

        dt.DefaultView.Sort = "TechID"
        'Dim ModifyAcct As String ="" '異動者
        'Dim ModifyDate As String ="" '異動時間

        Dim i_currentRow As Integer = 1
        For Each dr As DataRow In dt.Rows
            i_currentRow += 1
            'Dim RID As String '機構
            Dim RIDLevel As String = "" '計畫階層
            Dim TeacherID As String = "" '講師代碼
            Dim TeachCName As String = "" '講師姓名
            Dim TeachEName As String = "" '講師英文姓名
            Dim PassPortNO As String = "" '身份別
            Dim IDNO As String = "" '身分證號碼
            Dim Birthday As String = "" '出生日期
            Dim Sex As String = "" '性別
            'Dim TMID As String = "" '職類代碼
            Dim GCODE2 As String = "" '職類代碼-GCODE2

            Dim IVID As String = "" '職稱代碼
            Dim KindEngage As String = "" '內外聘
            Dim KindID As String = "" '師資別
            Dim DegreeID As String = "" '學歷
            Dim GraduateStatus As String = "" '畢業狀況
            Dim SchoolName As String = "" '學校名稱
            Dim Department As String = "" '科系名稱
            Dim Phone As String = "" '聯絡電話
            Dim Mobile As String = "" '行動電話
            Dim Email As String = "" 'E_Mail
            Dim AddressZip As String = "" '戶藉地址Zip
            Dim AddressZIP6W As String = "" '戶藉地址Zip後2碼
            Dim Address As String = "" '戶藉地址
            Dim WorkOrg As String = "" '服務單位名稱
            Dim ExpYears As String = "" '服務年資
            Dim ServDept As String = "" '服務部門
            Dim WorkPhone As String = "" '服務單位電話
            Dim Fax As String = "" '服務單位傳真
            Dim WorkZip As String = "" '服務單位地址Zip
            Dim WorkZIP6W As String = "" '服務單位地址Zip後2碼
            Dim Workaddr As String = "" ' 服務單位地址
            Dim ExpUnit1 As String = "" '服務單位一
            Dim ExpUnit2 As String = "" '服務單位二
            Dim ExpUnit3 As String = "" '服務單位三
            Dim ExpYears1 As String = "" '服務年資一
            Dim ExpYears2 As String = "" '服務年資二
            Dim ExpYears3 As String = "" '服務年資三
            Dim EpINV1 As String = "" '服務職稱1
            Dim EpINV2 As String = "" '服務職稱2
            Dim EpINV3 As String = "" '服務職稱3
            Dim ExpSDate1 As String = "" '服務單位一起日
            Dim ExpEDate1 As String = "" '服務單位一迄日
            Dim ExpSDate2 As String = "" '服務單位二起日
            Dim ExpEDate2 As String = "" '服務單位二迄日
            Dim ExpSDate3 As String = "" '服務單位三起日
            Dim ExpEDate3 As String = "" '服務單位三迄日
            Dim Specialty1 As String = "" '專長一
            Dim Specialty2 As String = "" '專長二
            Dim Specialty3 As String = "" '專長三
            Dim Specialty4 As String = "" '專長四
            Dim Specialty5 As String = "" '專長五
            Dim Specialty1b As String = "" '專長一b
            Dim Specialty2b As String = "" '專長二b
            Dim Specialty3b As String = "" '專長三b
            Dim Specialty4b As String = "" '專長四b
            Dim Specialty5b As String = "" '專長五b

            Dim TransBook As String = "" '譯著

            'Dim ProLicense As String ="" '專業證照
            Dim ProLicense1 As String = "" '專業證照(政府)
            Dim ProLicense2 As String = "" '專業證照(其他)
            Dim WorkStatus As String = "" '任職狀況
            Dim TechType1 As String = "" '講師類別
            Dim TechType2 As String = "" '助教類別
            'RID = Right(dr("RID").ToString, 2)
            'If center.Value <> "" Then center.Value = Trim(center.Value)
            RIDLevel = TIMS.ClearSQM(center.Value)
            TeacherID = TIMS.ClearSQM(dr("TeacherID"))
            TeachCName = TIMS.ClearSQM(dr("TeachCName"))
            TeachEName = TIMS.ClearSQM(dr("TeachEName"))
            Select Case TIMS.ClearSQM(dr("PassPortNO"))
                Case "1", "2"
                    PassPortNO = dr("PassPortNO").ToString
                Case Else
                    PassPortNO = "2"
            End Select
            IDNO = TIMS.ClearSQM(dr("IDNO"))
            Dim flag_idno_ok As Boolean = False
            If IDNO <> "" AndAlso TIMS.CheckIDNO(IDNO) Then flag_idno_ok = True
            If flag_idno_ok Then
                Sex = If(IDNO.Chars(1) = "1", "M", If(IDNO.Chars(1) = "2", "F", ""))
            End If
            'Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) '
            If v_rblWorkMode = TIMS.cst_wmdip1 Then IDNO = TIMS.strMask(IDNO, 1)
            Birthday = ""
            If Convert.ToString(dr("Birthday")) <> "" Then
                Birthday = TIMS.Cdate3(dr("Birthday"))
                If v_rblWorkMode = TIMS.cst_wmdip1 Then Birthday = TIMS.strMask(Birthday, 2)
            End If
            'Birthday = dr("Birthday").ToString
            If Convert.ToString(dr("Sex")) <> "" Then Sex = dr("Sex").ToString

            'TMID-OLD
            'Dim v_TB_career_id_txt As String = ""
            'If Convert.ToString(dr("TrainID")) <> "" Then
            '    '職類
            '    v_TB_career_id_txt = "[" & Convert.ToString(dr("TrainID")) & "]" & Convert.ToString(dr("TrainName"))
            '    TB_career_id.Text = v_TB_career_id_txt 'TMID-OLD
            'End If
            'If v_TB_career_id_txt = "" AndAlso Convert.ToString(dr("JobID")) <> "" Then
            '    '若取不到職類-但有業別-顯示業別
            '    v_TB_career_id_txt = "[" & Convert.ToString(dr("JobID")) & "]" & Convert.ToString(dr("JobName"))
            '    TB_career_id.Text = v_TB_career_id_txt 'TMID-OLD
            'End If
            'GCODE2-NEW
            'Dim v_GCODE2 As String = ""
            'ff3 = "TMID='" & Convert.ToString(dr("TMID")) & "'"
            'If dtGOVCLASSCAST3.Select(ff3).Length > 0 Then v_GCODE2 = "[" & dtGOVCLASSCAST3.Select(ff3)(0)("GCODE2") & "]" & dtGOVCLASSCAST3.Select(ff3)(0)("CNAME")
            'If v_GCODE2 <> "" Then TB_career_id.Text = v_GCODE2 'v_TB_career_id_txt
            ff3 = "TMID='" & Convert.ToString(dr("TMID")) & "'"
            If dtGOVCLASSCAST3.Select(ff3).Length > 0 Then GCODE2 = dtGOVCLASSCAST3.Select(ff3)(0)("GCODE2")

            IVID = TIMS.ClearSQM(dr("Invest"))
            If RIDValue.Value.Length = 1 Then IVID = TIMS.ClearSQM(dr("IVID"))
            KindEngage = dr("KindEngage").ToString
            KindID = dr("KindID").ToString
            DegreeID = dr("DegreeID").ToString
            GraduateStatus = dr("GraduateStatus").ToString
            SchoolName = TIMS.ClearSQM(dr("SchoolName"))
            Department = TIMS.ClearSQM(dr("Department"))
            Phone = TIMS.ClearSQM(dr("Phone"))
            Mobile = TIMS.ClearSQM(dr("Mobile"))
            Email = TIMS.ClearSQM(dr("Email"))

            AddressZip = TIMS.ClearSQM(dr("AddressZip"))
            AddressZIP6W = TIMS.ClearSQM(dr("AddressZIP6W"))
            Address = TIMS.ClearSQM(dr("Address"))

            WorkOrg = TIMS.ClearSQM(dr("WorkOrg"))
            ExpYears = TIMS.ClearSQM(dr("ExpYears"))
            ServDept = TIMS.ClearSQM(dr("ServDept"))
            WorkPhone = TIMS.ClearSQM(dr("WorkPhone"))
            Fax = TIMS.ClearSQM(dr("Fax"))
            WorkZip = TIMS.ClearSQM(dr("WorkZip"))
            WorkZIP6W = TIMS.ClearSQM(dr("WorkZIP6W"))
            Workaddr = TIMS.ClearSQM(dr("Workaddr"))
            ExpUnit1 = TIMS.ClearSQM(dr("ExpUnit1"))
            ExpUnit2 = TIMS.ClearSQM(dr("ExpUnit2"))
            ExpUnit3 = TIMS.ClearSQM(dr("ExpUnit3"))
            ExpYears1 = TIMS.ClearSQM(dr("ExpYears1"))
            ExpYears2 = TIMS.ClearSQM(dr("ExpYears2"))
            ExpYears3 = TIMS.ClearSQM(dr("ExpYears3"))
            EpINV1 = TIMS.ClearSQM(dr("INV1"))
            EpINV2 = TIMS.ClearSQM(dr("INV2"))
            EpINV3 = TIMS.ClearSQM(dr("INV3"))
            ExpSDate1 = If(Convert.ToString(dr("ExpSDate1")) <> "", TIMS.Cdate3(dr("ExpSDate1")), "")
            ExpSDate2 = If(Convert.ToString(dr("ExpSDate2")) <> "", TIMS.Cdate3(dr("ExpSDate2")), "")
            ExpSDate3 = If(Convert.ToString(dr("ExpSDate3")) <> "", TIMS.Cdate3(dr("ExpSDate3")), "")
            ExpEDate1 = If(Convert.ToString(dr("ExpEDate1")) <> "", TIMS.Cdate3(dr("ExpEDate1")), "")
            ExpEDate2 = If(Convert.ToString(dr("ExpEDate2")) <> "", TIMS.Cdate3(dr("ExpEDate2")), "")
            ExpEDate3 = If(Convert.ToString(dr("ExpEDate3")) <> "", TIMS.Cdate3(dr("ExpEDate3")), "")

            Specialty1 = TIMS.ChangeSQM(dr("Specialty1")) '專長一
            Specialty2 = TIMS.ChangeSQM(dr("Specialty2")) '專長二
            Specialty3 = TIMS.ChangeSQM(dr("Specialty3")) '專長三
            Specialty4 = TIMS.ChangeSQM(dr("Specialty4")) '專長四
            Specialty5 = TIMS.ChangeSQM(dr("Specialty5")) '專長五

            'Call SplitSTR250(Specialty1, Specialty1b)
            'Call SplitSTR250(Specialty2, Specialty2b)
            'Call SplitSTR250(Specialty3, Specialty3b)
            'Call SplitSTR250(Specialty4, Specialty4b)
            'Call SplitSTR250(Specialty5, Specialty5b)

            TransBook = TIMS.ChangeSQM(dr("TransBook"))   '譯著
            'ProLicense = TIMS.ChangeSQM(dr("ProLicense")) '專業證照
            ProLicense1 = TIMS.ChangeSQM(dr("ProLicense1")) '專業證照(政府)
            ProLicense2 = TIMS.ChangeSQM(dr("ProLicense2")) '專業證照(其他)
            WorkStatus = TIMS.ChangeSQM(dr("WorkStatus")) '任職狀況
            TechType1 = TIMS.ChangeSQM(dr("TechType1")) '講師類別
            TechType2 = TIMS.ChangeSQM(dr("TechType2")) '助教類別
            'ModifyAcct = dr("ModifyAcct").ToString
            'ModifyDate = dr("Specialty1").ToString
            'PassPortNO = dr("PassPortNO").ToString
            sheet.Cells(i_currentRow, cst_iCol_RIDLevel).Value = RIDLevel
            sheet.Cells(i_currentRow, cst_iCol_TeacherID).Value = TeacherID
            sheet.Cells(i_currentRow, cst_iCol_TeachCName).Value = TeachCName
            sheet.Cells(i_currentRow, cst_iCol_TeachEName).Value = TeachEName
            sheet.Cells(i_currentRow, cst_iCol_PassPortNO).Value = PassPortNO
            sheet.Cells(i_currentRow, cst_iCol_IDNO).Value = IDNO
            sheet.Cells(i_currentRow, cst_iCol_Birthday).Value = Birthday
            sheet.Cells(i_currentRow, cst_iCol_Sex).Value = Sex
            sheet.Cells(i_currentRow, cst_iCol_GCODE2).Value = GCODE2
            sheet.Cells(i_currentRow, cst_iCol_IVID).Value = IVID

            sheet.Cells(i_currentRow, cst_iCol_KindEngage).Value = KindEngage
            sheet.Cells(i_currentRow, cst_iCol_KindID).Value = KindID
            sheet.Cells(i_currentRow, cst_iCol_DegreeID).Value = DegreeID
            sheet.Cells(i_currentRow, cst_iCol_GraduateStatus).Value = GraduateStatus
            sheet.Cells(i_currentRow, cst_iCol_SchoolName).Value = SchoolName
            sheet.Cells(i_currentRow, cst_iCol_Department).Value = Department
            sheet.Cells(i_currentRow, cst_iCol_Phone).Value = Phone
            sheet.Cells(i_currentRow, cst_iCol_Mobile).Value = Mobile
            sheet.Cells(i_currentRow, cst_iCol_Email).Value = Email

            sheet.Cells(i_currentRow, cst_iCol_AddressZip).Value = AddressZip
            sheet.Cells(i_currentRow, cst_iCol_AddressZIP6W).Value = AddressZIP6W
            sheet.Cells(i_currentRow, cst_iCol_Address).Value = Address
            sheet.Cells(i_currentRow, cst_iCol_WorkOrg).Value = WorkOrg
            sheet.Cells(i_currentRow, cst_iCol_ExpYears).Value = ExpYears
            sheet.Cells(i_currentRow, cst_iCol_ServDept).Value = ServDept
            sheet.Cells(i_currentRow, cst_iCol_WorkPhone).Value = WorkPhone
            sheet.Cells(i_currentRow, cst_iCol_Fax).Value = Fax

            sheet.Cells(i_currentRow, cst_iCol_WorkZip).Value = WorkZip
            sheet.Cells(i_currentRow, cst_iCol_WorkZIP6W).Value = WorkZIP6W
            sheet.Cells(i_currentRow, cst_iCol_Workaddr).Value = Workaddr
            sheet.Cells(i_currentRow, cst_iCol_ExpUnit1).Value = ExpUnit1
            sheet.Cells(i_currentRow, cst_iCol_ExpUnit2).Value = ExpUnit2
            sheet.Cells(i_currentRow, cst_iCol_ExpUnit3).Value = ExpUnit3
            sheet.Cells(i_currentRow, cst_iCol_ExpYears1).Value = ExpYears1
            sheet.Cells(i_currentRow, cst_iCol_ExpYears2).Value = ExpYears2
            sheet.Cells(i_currentRow, cst_iCol_ExpYears2).Value = ExpYears2

            sheet.Cells(i_currentRow, cst_iCol_EpINV1).Value = EpINV1
            sheet.Cells(i_currentRow, cst_iCol_EpINV2).Value = EpINV2
            sheet.Cells(i_currentRow, cst_iCol_EpINV3).Value = EpINV3
            sheet.Cells(i_currentRow, cst_iCol_ExpSDate1).Value = ExpSDate1
            sheet.Cells(i_currentRow, cst_iCol_ExpEDate1).Value = ExpEDate1
            sheet.Cells(i_currentRow, cst_iCol_ExpSDate2).Value = ExpSDate2
            sheet.Cells(i_currentRow, cst_iCol_ExpEDate2).Value = ExpEDate2
            sheet.Cells(i_currentRow, cst_iCol_ExpSDate3).Value = ExpSDate3
            sheet.Cells(i_currentRow, cst_iCol_ExpEDate3).Value = ExpEDate3

            sheet.Cells(i_currentRow, cst_iCol_Specialty1).Value = Specialty1
            sheet.Cells(i_currentRow, cst_iCol_Specialty2).Value = Specialty2
            sheet.Cells(i_currentRow, cst_iCol_Specialty3).Value = Specialty3
            sheet.Cells(i_currentRow, cst_iCol_Specialty4).Value = Specialty4
            sheet.Cells(i_currentRow, cst_iCol_Specialty5).Value = Specialty5
            sheet.Cells(i_currentRow, cst_iCol_TransBook).Value = TransBook

            sheet.Cells(i_currentRow, cst_iCol_ProLicense1).Value = ProLicense1
            sheet.Cells(i_currentRow, cst_iCol_ProLicense2).Value = ProLicense2
            sheet.Cells(i_currentRow, cst_iCol_WorkStatus).Value = WorkStatus
            sheet.Cells(i_currentRow, cst_iCol_TechType1).Value = TechType1
            sheet.Cells(i_currentRow, cst_iCol_TechType2).Value = TechType2
        Next

    End Sub


#Region "NO USE"
    'Function ODS_Export_HTML(dt As DataTable) As String
    '    Dim strHTML As String = ""
    '    strHTML &= ("<div>")
    '    strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
    '    'Common.RespWrite(Me, "<tr>")
    '    Dim sPattern As String = ""
    '    sPattern &= "計劃階層,講師代碼,講師姓名,講師英文姓名,身份別,身分證字號,出生日期,性別,主要職類,職稱"
    '    sPattern &= ",內外聘,師資別,最高學歷,畢業狀況,學校名稱,科系名稱,聯絡電話,行動電話,電子郵件"
    '    sPattern &= ",郵遞區號前3碼,郵遞區號6碼,戶籍地址,服務單位名稱,年資,服務部門,服務單位電話,服務單位傳真"
    '    sPattern &= ",服務單位郵遞區號前3碼,服務單位郵遞區號6碼,服務單位地址,服務單位一,服務單位二,服務單位三"
    '    sPattern &= ",服務年資一,服務年資二,服務年資三,服務職稱一,服務職稱二,服務職稱三"
    '    sPattern &= ",服務期間一起日,服務期間一迄日,服務期間二起日,服務期間二迄日,服務期間三起日,服務期間三迄日,專長一,專長二,專長三,專長四,專長五,譯著"
    '    sPattern &= ",專業證照(政府),專業證照(其他),排課使用,講師類別,助教類別"
    '    Dim sColumn As String = ""
    '    sColumn &= "ORGNAME,TeacherID,TeachCName,TeachEName,PassPortNO,IDNO,Birthday,Sex,GCODE2,IVID"
    '    sColumn &= ",KindEngage,KindID,DegreeID,GraduateStatus,SchoolName,Department,Phone,Mobile,Email"
    '    sColumn &= ",AddressZip,AddressZIP6W,Address,WorkOrg,ExpYears,ServDept,WorkPhone,Fax"
    '    sColumn &= ",WorkZip,WorkZIP6W,Workaddr,ExpUnit1,ExpUnit2,ExpUnit3"
    '    sColumn &= ",ExpYears1,ExpYears2,ExpYears3,INV1,INV2,INV3"
    '    sColumn &= ",ExpSDate1,ExpEDate1,ExpSDate2,ExpEDate2,ExpSDate3,ExpEDate3,Specialty1,Specialty2,Specialty3,Specialty4,Specialty5,TransBook"
    '    sColumn &= ",ProLicense1,ProLicense2,WorkStatus,TechType1,TechType2"

    '    Dim sPatternA() As String = Split(sPattern, ",")
    '    Dim sColumnA() As String = Split(sColumn, ",")

    '    '標題抬頭
    '    Dim ExportStr As String = "" '建立輸出文字
    '    ExportStr = "<tr>"
    '    For i As Integer = 0 To sPatternA.Length - 1
    '        ExportStr &= "<td>" & sPatternA(i) & "</td>" '& vbTab
    '    Next
    '    ExportStr &= "</tr>" & vbCrLf
    '    'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
    '    strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

    '    '建立資料面
    '    Dim iNum As Integer = 0
    '    For Each dr As DataRow In dt.DefaultView.Table.Rows
    '        iNum += 1
    '        ExportStr = "<tr>"
    '        For i As Integer = 0 To sColumnA.Length - 1
    '            Select Case sColumnA(i)
    '                Case "ORGNAME"
    '                    ExportStr &= String.Concat("<td>", center.Value, "</td>") '& vbTab
    '                Case "GCODE2"
    '                    ff3 = "TMID='" & Convert.ToString(dr("TMID")) & "'"
    '                    Dim V_GCODE2 As String = If(dtGOVCLASSCAST3.Select(ff3).Length > 0, dtGOVCLASSCAST3.Select(ff3)(0)("GCODE2"), "")
    '                    ExportStr &= String.Concat("<td>", V_GCODE2, "</td>") '& vbTab
    '                Case "IVID"
    '                    Dim V_IVID As String = TIMS.ClearSQM(dr("Invest"))
    '                    If RIDValue.Value.Length = 1 Then V_IVID = TIMS.ClearSQM(dr("IVID"))
    '                    ExportStr &= String.Concat("<td>", V_IVID, "</td>") '& vbTab
    '                Case "Birthday", "ExpSDate1", "ExpEDate1", "ExpSDate2", "ExpEDate2", "ExpSDate3", "ExpEDate3"
    '                    Dim V_DATATXT As String = TIMS.cdate3(dr(sColumnA(i)))
    '                    ExportStr &= String.Concat("<td>", V_DATATXT, "</td>") '& vbTab
    '                Case Else
    '                    ExportStr &= "<td>" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
    '            End Select
    '        Next
    '        ExportStr &= "</tr>" & vbCrLf
    '        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
    '    Next
    '    strHTML &= ("</div>")
    '    Return strHTML
    'End Function

    'Public Shared Sub SplitSTR250(ByRef Spec1 As String, ByRef Spec1B As String)
    '    Spec1B = ""
    '    Const i_Maxlen As Integer = 250
    '    If Spec1 <> "" AndAlso Spec1.Length > i_Maxlen Then
    '        Spec1B = Spec1.Substring(i_Maxlen, Spec1.Length - i_Maxlen)
    '        Spec1 = Spec1.Substring(1, i_Maxlen)
    '    End If
    'End Sub

    'Public Shared Sub UPDATE_LongText(ByRef MyConn As OleDb.OleDbConnection, LONGVALUE1 As String, PKVALUE1 As String, LONGCOLUMN1 As String, PKCOLUMN1 As String)
    '    'Dim LONGVALUE1 As String = TIMS.GetMyValue2(sParms, "LONGVALUE1")
    '    'Dim PKVALUE1 As String = TIMS.GetMyValue2(sParms, "PKVALUE1")
    '    'Dim LONGCOLUMN1 As String = TIMS.GetMyValue2(sParms, "LONGCOLUMN1")
    '    'Dim PKCOLUMN1 As String = TIMS.GetMyValue2(sParms, "PKCOLUMN1")
    '    If LONGVALUE1 = "" Then Return
    '    Dim u_sql As String = String.Concat("update [Sheet1$] SET ", LONGCOLUMN1, "=", LONGCOLUMN1, " +'", LONGVALUE1, "' WHERE ", PKCOLUMN1, "='" & PKVALUE1 & "'")
    '    Dim ole_uCmd As New OleDb.OleDbCommand(u_sql, MyConn)
    '    If MyConn.State = ConnectionState.Closed Then MyConn.Open()
    '    ole_uCmd.ExecuteNonQuery()
    'End Sub
#End Region

    Sub ExpExccl_1(ByRef strErrmsg As String, ByRef MyPath As String)
        '將新建立的excel存入記憶體下載-----   Start
        'Dim strErrmsg As String = ""
        'strErrmsg = ""
        Dim myFileName1 As String = String.Concat(TIMS.ClearSQM("TeacherList" & TIMS.GetRnd6Eng), ".xls") '檔名
        Try
            Dim fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
            Dim br As New System.IO.BinaryReader(fr)
            Dim buf(fr.Length) As Byte
            fr.Read(buf, 0, fr.Length)
            fr.Close()

            Response.Clear()
            Response.ClearHeaders()
            Response.Buffer = True
            Response.AppendHeader("Content-Disposition", "attachment;filename=" & HttpUtility.UrlEncode(myFileName1, System.Text.Encoding.UTF8))
            'Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8))
            Response.ContentType = "Application/vnd.ms-Excel"
            'Common.RespWrite(Me, br.ReadBytes(fr.Length))
            Response.BinaryWrite(buf)
            'Response.End()
        Catch ex As Exception
            'Dim strErrmsg As String = ""
            'strErrmsg = ""
            'strErrmsg &= "/* ex.ToString: */" & vbCrLf & ex.ToString & vbCrLf
            'strErrmsg &= "sql:" & vbCrLf & Sql & vbCrLf
            'strErrmsg &= "MyPath:" & MyPath & vbCrLf
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            'Call TIMS.WriteTraceLog(strErrmsg)

            strErrmsg = $"無法存取該檔案!!!{vbCrLf} (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) {vbCrLf}{ex.Message}"
            'strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)

            'Finally '刪除Temp中的資料 'If MyFile.Exists(MyPath) Then MyFile.Delete(MyPath)
        End Try

    End Sub

    Dim objLock_ExportXLSX As New Object

    ''' <summary>
    ''' '查詢鈕  '匯出鈕 'hidSchBtnNum.value: 1.正常查詢 2.正常匯出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub SUtl_btnSearchData1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim BtnObj As Button = CType(sender, Button)
        Const cst_button1 As String = "button1" '查詢
        Const cst_btnxlsemport As String = "btnxlsemport" '匯出'Btn_XlsEmport_Click
        Const cst_btndivPwdSubmit As String = "btndivpwdsubmit" 'hidSchBtnNum.value: 1.正常查詢 2.正常匯出
        Dim sMsg As String = ""
        eMeng.Style.Item("display") = "none"
        Datagrid2.Visible = False

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Select Case LCase(BtnObj.CommandName)
            Case cst_button1 '查詢鈕
                Call gClickSearchButton()

            Case cst_btnxlsemport '匯出鈕
                SyncLock objLock_ExportXLSX
                    Call Export1XLS() 'Btn_XlsEmport_Click
                End SyncLock

            Case cst_btndivPwdSubmit
                '正常顯示 '查詢或匯出。
                If Not TIMS.sUtl_ChkPlanPwd(sm.UserInfo.PlanID, objconn) Then
                    sMsg = "未設定計畫密碼!!"
                    labChkMsg.Text = sMsg
                    Common.MessageBox(Me, sMsg)
                    Exit Sub
                End If
                If Not TIMS.sUtl_ChkPlanPwdOK(objconn, sm.UserInfo.PlanID, txtdivPaswrd.Text) Then
                    sMsg = "個資安全密碼錯誤!!"
                    labChkMsg.Text = sMsg
                    Common.MessageBox(Me, sMsg)
                    Exit Sub
                End If
                'If rblWorkMode.SelectedValue = "2" Then flgCIShow = True '可正常顯示個資。
                txtdivPaswrd.Text = ""
                Select Case hidSchBtnNum.Value
                    Case "1"
                        Call gClickSearchButton()
                    Case "2"
                        SyncLock objLock_ExportXLSX
                            Call Export1XLS() '匯出-Btn_XlsEmport_Click
                        End SyncLock
                End Select
        End Select
    End Sub

    Protected Sub BtnImpYear_Click(sender As Object, e As EventArgs) Handles BtnImpYear.Click
        Dim url1 As String = $"TC_01_027_imp.aspx?ID={TIMS.Get_MRqID(Me)}"
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Protected Sub Btn_XlsEmport_Click(sender As Object, e As EventArgs) Handles Btn_XlsEmport.Click

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

#Region "NO USE"
    'Protected Sub btnCloseLoginDiv_Click(sender As Object, e As EventArgs) Handles btnCloseLoginDiv.Click
    '    'panelLoginDiv.Visible = False
    '    panelLoginDiv.Style.Item("display") = "none"
    '    labChkMsg.Text = ""
    'End Sub

    'Function Check_Data_Protection_State() As Boolean
    '    Dim rst As Boolean = False
    '    If rblWorkMode.SelectedValue = "1" Then
    '        rst = False
    '    Else
    '        If hidCheckPasswordState.Value = "True" Then    'Password checked
    '            'panelLoginDiv.Visible = False
    '            panelLoginDiv.Style.Item("display") = "none"
    '        Else
    '            If hidLockTime1.Value = "1" Or hidLockTime1.Value = "" Then
    '                'panelLoginDiv.Visible = True
    '                panelLoginDiv.Style.Item("display") = "inline"
    '                rst = True
    '            End If
    '        End If
    '    End If
    '    If hidLockTime1.Value = "0" Then rblWorkMode.SelectedValue = "2"
    '    hidWorkMode.Value = rblWorkMode.SelectedValue
    '    Return rst
    'End Function

    'Protected Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
    '    If Not TIMS.sUtl_ChkPlanPwd(sm.UserInfo.PlanID) Then
    '        labChkMsg.Text = "尚未設定個資安全密碼!!"
    '        'panelLoginDiv.Visible = True
    '        panelLoginDiv.Style.Item("display") = "inline"
    '        hidCheckPasswordState.Value = "False"
    '        Exit Sub
    '    End If
    '    If Not TIMS.sUtl_ChkPlanPwdOK(sm.UserInfo.PlanID, txtdivPaswrd.Text) Then
    '        labChkMsg.Text = "個資安全密碼錯誤!!"
    '        'panelLoginDiv.Visible = True
    '        panelLoginDiv.Style.Item("display") = "inline"
    '        hidCheckPasswordState.Value = "False"
    '    Else
    '        'panelLoginDiv.Visible = False
    '        panelLoginDiv.Style.Item("display") = "none"
    '        hidCheckPasswordState.Value = "True"
    '        'Button1_Click(sender, e)
    '        Call gClickSearchButton()
    '    End If
    'End Sub

    'sUtl_btnSearchData1
    'Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    'Call gClickSearchButton()
    'End Sub

    'AddHandler Btn_XlsEmport.Click, AddressOf sUtl_btnSearchData1 '匯出
    'Protected Sub Btn_XlsEmport_Click(sender As Object, e As EventArgs) Handles Btn_XlsEmport.Click
    'End Sub
#End Region

End Class
