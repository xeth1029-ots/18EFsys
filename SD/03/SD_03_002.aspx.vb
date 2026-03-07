Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Partial Class SD_03_002
    Inherits AuthBasePage

#Region "Declaration Area"
    'in_class_stud: 學員資料卡。
    'confidential information
    '~\SD\03\Sample2.xls

    'SD_03_002_classver (學員資料審核)

    'Const cst_SD03002_addaspx As String="SD_03_002_add.aspx"
    '28:產業人才投資方案
    'Const cst_SD03002_add2aspx As String="SD_03_002_add2.aspx"
    '06:在職進修訓練/'70:區域產業據點職業訓練計畫(在職)
    'FROM dbo.CLASS_STUDENTSOFCLASS a 'JOIN dbo.CLASS_CLASSINFO cc On a.OCID=cc.OCID 'JOIN dbo.ID_CLASS s3 On s3.CLSID=cc.CLSID
    'JOIN dbo.VIEW_PLAN ip On ip.PLANID= cc.PLANID
    'JOIN dbo.STUD_STUDENTINFO b On a.SID=b.SID
    'JOIN dbo.STUD_SUBDATA c On a.SID=c.SID
    'LEFT JOIN dbo.STUD_SERVICEPLACE d On a.SOCID=d.SOCID
    'LEFT JOIN dbo.STUD_TRAINBG e On a.SOCID=e.SOCID

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    Const CST_KD_STUDENTLIST As String = "StudentList" 'Session("IDNOArray")
    Dim flag_show_actno_budid As Boolean = False '保險證號/預算別代碼 false:不顯示 true:顯示
    'Dim dtBli28 As DataTable=Nothing
    Dim dtBli28e As DataTable = Nothing
    Dim CPdt As DataTable = Nothing
    Dim ff As String = ""
    Const cst_flgCIShow As String = "flgCIShow"
    Dim flgCIShow As Boolean = False '是否可正常顯示個資。false:不可以 true:可以。
    'ViewState(cst_flgCIShow) 第1次查詢記錄到 ViewState
    Dim drOCID As DataRow = Nothing '班級資料查詢。

    Dim blnTPlanUseEcfa As Boolean = False '該計畫是否使用ECFA True:使用 False:不使用
    '屆退官兵者 (依系統日期判斷)
    'Dim flagTPlanID02Plan2 As Boolean=False '判斷計畫為自辦職前。

    'OJT-21020401：在職進修訓練(自辦) - 學員資料維護：判斷學員為現役軍人時於投保保險證號顯示「在役軍人」、預算別判斷為「就安」
    Dim flagTPlanID06Plan3 As Boolean = False
    Const cst_Serviceman As String = "在役軍人" '在役軍人

    Dim MySqlStr As String
    'Dim MySOCID As String
    Dim OrgKind2 As String   '用來儲存所選班級的機構別
    Dim Key_Degree As DataTable
    Dim Key_GradState As DataTable
    Dim Key_Military As DataTable
    Dim dtIdentity As DataTable

    Dim Key_Subsidy As DataTable
    Dim Key_HandicatType As DataTable
    Dim Key_HandicatLevel As DataTable
    Dim Key_JoblessWeek As DataTable
    Dim Plan_Budget As DataTable
    'Dim PageControler1 As New PageControler
    Dim IDNOArray As New ArrayList

    Const cst_msgNoStdData As String = "查無學生資料!"
    'Const cst_msgTPlanID28NoStdData As String="請至學員參訓功能進行報到作業!"
    Const cst_msgTPlanID28NoStdData As String = "「學員參訓」作業尚未完成!!"

    Const cst_選取 As Integer = 0
    Const cst_學號 As Integer = 1
    Const cst_姓名 As Integer = 2
    Const cst_身分證號碼 As Integer = 3
    Const cst_性別 As Integer = 4
    Const cst_出生日期 As Integer = 5
    Const cst_報名路徑 As Integer = 6 'Dim EnterChannel As String="" '報名管道"/報名路徑
    Const cst_學員狀態 As Integer = 7
    'Const cst_socid As Integer=8
    Const cst_保險證號 As Integer = 8
    Const cst_預算別 As Integer = 9
    Const cst_功能 As Integer = 10

#End Region

    ' 共用設定
    'Dim fontName As String="標楷體"
    Dim fontSize12s As Single = 12.0F
    'Dim print_lock As New Object '(); //lock

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load ', Me.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        blnTPlanUseEcfa = TIMS.CheckTPlanUseEcfa(sm.UserInfo.TPlanID)

        '啟用鎖定。
        Dim work2015 As String = TIMS.Utl_GetConfigSet("work2015")
        hidLockTime2.Value = If(work2015 = "Y", "1", "2")

        AddHandler Button1.Click, AddressOf SUtl_btnSearchData1 '查詢
        'AddHandler Button6.Click, AddressOf sUtl_btnSearchData1 '匯出
        AddHandler btndivPwdSubmit.Click, AddressOf SUtl_btnSearchData1

        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值

        'flagTPlanID02Plan2=False '判斷計畫為自辦職前。'屆退官兵者 (依系統日期判斷)
        'If TIMS.Cst_TPlanID02Plan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagTPlanID02Plan2=True '判斷計畫為自辦職前。

        'OJT-21020401：在職進修訓練(自辦) - 學員資料維護：判斷學員為現役軍人時於投保保險證號顯示「在役軍人」、預算別判斷為「就安」
        flagTPlanID06Plan3 = False '判斷計畫為 在職進修訓練(自辦)。
        If TIMS.Cst_TPlanID06Plan3.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagTPlanID06Plan3 = True '在職進修訓練(自辦)。

        '啟動個資法。
        Button1.Attributes.Add("onclick", "return showLoginPwdDiv(1);")
        Button1.CommandName = "Button1"
        Button6.Attributes.Add("onclick", "return showLoginPwdDiv(2);")
        Button6.CommandName = "Button6"

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then
        msg.Text = ""
        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))
        trImport1.Visible = False '停用學員匯入功能。
        trRBListExpType.Visible = False
        trExport1.Visible = False

        Call SUtl_Create0()

        If Not IsPostBack Then
            'ImportTable.Style.Item("display")="none"
            trImport1.Style.Item("display") = "none"
            trRBListExpType.Style.Item("display") = "none"
            trExport1.Style.Item("display") = "none"

            DataGridTable.Style.Item("display") = "none"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        'If sm.UserInfo.TPlanID <> "28" Then
        '    OCIDStar.Visible=True
        '    Button1.Attributes("onclick")="javascript:return search()"
        'Else
        '    OCIDStar.Visible=False
        'End If
        'Button1.Attributes("onclick")="javascript:return search()"
        Button4.Attributes("onclick") = "CheckPrint();return false;"

        Dim iPlanKind As Integer = TIMS.Get_PlanKind(Me, objconn)
        Dim js_btn5_onclick As String = "choose_class(1);"
        If iPlanKind = 1 Then js_btn5_onclick = "choose_class(2);"
        button5.Attributes("onclick") = js_btn5_onclick

        'Button6.Attributes("onclick")="if(document.getElementById('DataGridTable').style.display=='none'){alert('無學員資料可以匯出!');return false;}"
        button7.Attributes("onclick") = "if(document.form1.File1.value==''){alert('請選擇匯入檔案的路徑');return false;}"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg.aspx?name=Stu_Maintain');SetOneOCID();"
        Const cst_javascript_openOrg_FMT2 As String = "javascript:openOrg('../../Common/LevOrg1.aspx');SetOneOCID();"
        Button8.Attributes("onclick") = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, cst_javascript_openOrg_FMT1, cst_javascript_openOrg_FMT2)

        '../../Doc/ClassStudent.zip
        hyperlink1.NavigateUrl = "../../Doc/ClassStudent_v14.zip"
        Button4.Visible = True '列印資料卡
        Button11.Visible = False '學員資料確認
        edit_but.Visible = False '學員資料審核

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            hyperlink1.NavigateUrl = "../../Doc/ClassStudentForTS_v15.zip"
            Button4.Visible = False '列印資料卡
            Button11.Visible = True '學員資料確認
            edit_but.Visible = True '學員資料審核
            Button11.ToolTip = "委訓單位進行學員資料確認"
            edit_but.ToolTip = "委訓單位進行學員資料審核"
        End If

        If Not IsPostBack Then
            CCreate1b()
        End If
        'by Milor 20080806--學員資料不再透過此功能進行新增，所以新增按鈕不再顯示，要恢復時再去除 'button2.Visible=False 'TIMS.Tooltip(button2, "停用")
    End Sub

    Sub SUtl_Create0()
        '是否為(後台)系統管理者-權限-(開放功能測試)
        If sm.UserInfo.LID = 0 AndAlso TIMS.IsSuperUser(Me, 1) Then
            tr_rblWorkMode.Visible = True
            tr_rblWorkMode.Disabled = False
            TIMS.Tooltip(tr_rblWorkMode, "系統管理者-權限-(開放功能測試)", True)
        Else
            'OJT-22080802 : 署：只能模糊顯示。(可直接隱藏 「資料顯示模式」選項)
            tr_rblWorkMode.Visible = If(sm.UserInfo.LID = 0, False, True)
            If sm.UserInfo.LID = 0 Then Common.SetListItem(rblWorkMode, TIMS.cst_wmdip1)
            tr_rblWorkMode.Disabled = If(sm.UserInfo.LID = 0, True, False)
        End If

        '70:區域產業據點職業訓練計畫(在職) 2020-12
        Dim flag_SHOW_2020x70 As Boolean = TIMS.SHOW_2020x70(sm)
        Dim flag_SHOW_2020x06 As Boolean = TIMS.SHOW_2020x06(sm)
        'Dim flag_show_actno_budid As Boolean=False '保險證號/預算別代碼 false:不顯示 true:顯示
        flag_show_actno_budid = False '保險證號/預算別代碼 false:不顯示 true:顯示
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_show_actno_budid = True
        If flag_SHOW_2020x70 Then flag_show_actno_budid = True
        If flag_SHOW_2020x06 Then flag_show_actno_budid = True

        '使用勞保勾稽,並且顯示預算別
        Hid_show_actno_budid.Value = ""
        If (flag_show_actno_budid) Then Hid_show_actno_budid.Value = "Y"
        If flag_SHOW_2020x70 Then Hid_show_actno_budid.Value = "Y"
        If flag_SHOW_2020x06 Then Hid_show_actno_budid.Value = "Y"

        '不使用補助比例，但要用勞保勾稽
        Hid_nouse_SupplyID.Value = ""
        If flag_SHOW_2020x70 Then Hid_nouse_SupplyID.Value = "Y"
        If flag_SHOW_2020x06 Then Hid_nouse_SupplyID.Value = "Y"
    End Sub

    Sub CCreate1b()
        'panelLoginDiv.Visible=False
        labChkMsg.Text = ""

        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

        Call USE_SearchStr()

        '若只有管理一個班級，自動協助帶出班級--by AMU 2009-02
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1, objconn)
    End Sub

    Sub GetSearchStr()
        Dim s_SearchStr As String = ""
        If ViewState("LastOCIDValue1") <> "" Then
            s_SearchStr = String.Concat("center=", center.Text)
            s_SearchStr += String.Concat("&RIDValue=", RIDValue.Value)
            s_SearchStr += String.Concat("&TMID1=", TMID1.Text)
            s_SearchStr += String.Concat("&TMIDValue1=", TMIDValue1.Value)
            s_SearchStr += String.Concat("&OCID1=", OCID1.Text)
            s_SearchStr += String.Concat("&OCIDValue1=", ViewState("LastOCIDValue1"))
            's_SearchStr += "OCIDValue1=" & OCIDValue1.Value & "&"
            s_SearchStr += String.Concat("&PageIndex=", DataGrid1.CurrentPageIndex + 1)
            s_SearchStr += String.Concat("&submit=", If(DataGrid1.Visible, 1, 0))
        End If
        Session("_SearchStr") = s_SearchStr
    End Sub

    Sub USE_SearchStr()
        'If Session("_SearchStr") IsNot Nothing Then 'Session("_SearchStr")=Nothing 'End If
        If Session("_SearchStr") Is Nothing Then Return
        Dim str_SearchStr As String = Convert.ToString(Session("_SearchStr"))
        Session("_SearchStr") = Nothing

        center.Text = TIMS.GetMyValue(str_SearchStr, "center")
        RIDValue.Value = TIMS.GetMyValue(str_SearchStr, "RIDValue")
        TMID1.Text = TIMS.GetMyValue(str_SearchStr, "TMID1")
        TMIDValue1.Value = TIMS.GetMyValue(str_SearchStr, "TMIDValue1")
        OCID1.Text = TIMS.GetMyValue(str_SearchStr, "OCID1")
        OCIDValue1.Value = TIMS.GetMyValue(str_SearchStr, "OCIDValue1")
        ViewState("_SearchStr") = str_SearchStr

        If ViewState("_SearchStr") IsNot Nothing Then
            If ViewState("_SearchStr").ToString.IndexOf("Load=") > -1 Then '被其他功能呼叫/SD_03_002_ver.aspx
                Me.Button12.Visible = True
                'ImportTable.Visible=False
                trImport1.Visible = False '停用學員匯入功能。
                trRBListExpType.Visible = False
                trExport1.Visible = False

                If Convert.ToString(Me.Request("OCID")) <> "" Then OCIDValue1.Value = Convert.ToString(Me.Request("OCID"))
            End If
        End If

        Dim MyValue As String = ""
        ViewState("PageIndex") = TIMS.GetMyValue(ViewState("_SearchStr"), "PageIndex")
        MyValue = TIMS.GetMyValue(ViewState("_SearchStr"), "submit")
        If MyValue = "1" Then
            'Call Button1_Click(sender, e)
            Call Search1()  '查詢按鈕 SQL

            If IsNumeric(ViewState("PageIndex")) Then
                '有資料SHOW出 跳頁
                PageControler1.PageIndex = ViewState("PageIndex")
                PageControler1.DataTableCreate(CPdt, PageControler1.Sort, PageControler1.PageIndex)
            End If
        End If
    End Sub

    '查詢按鈕 SQL
    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg8)
        '    Exit Sub
        'End If

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If

        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        'ImportTable.Style.Item("display")="" '"inline"
        trImport1.Style.Item("display") = ""
        trRBListExpType.Visible = True
        trExport1.Visible = True
        trRBListExpType.Style.Item("display") = "" '"inline"
        trExport1.Style.Item("display") = "" '"inline"

        '記載最後搜尋的班級資訊。
        ViewState("LastOCIDValue1") = ""
        If OCIDValue1.Value <> "" Then ViewState("LastOCIDValue1") = OCIDValue1.Value
        If OCIDValue1.Value = "" Then
            Common.RespWrite(Me, "<script>alert('" & "未選擇班級，請選擇" & "');</script>")
            Exit Sub
        End If
        'If Check_Data_Protection_State() Then Exit Sub '先不卡,待確認輸入密碼的邏輯再加上去

        If ViewState("LastOCIDValue1") = "" Then
            Common.RespWrite(Me, "<script>alert('" & "未選擇班級，請選擇" & "');</script>")
            Exit Sub
        End If

        '20090410(Milor)加入只能查登入年度的年度限制。
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            '停用-查無班／或不同年度
            edit_but.Visible = False
            'button2.Enabled=False '不可新增
            'TIMS.Tooltip(button2, "停用")
            button7.Enabled = False '不可匯入
            TIMS.Tooltip(button7, "停用", True)

            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If

        '若年度不合，刪除資訊
        If Convert.ToString(drCC("YEARS")) <> Right(sm.UserInfo.Years, 2) Then
            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If
        'If Convert.ToString(drC("YEARS")) <> Right(sm.UserInfo.Years, 2) Then drC=Nothing
        dateround.Text = TIMS.Cdate3(drCC("STDate")) & "~" & TIMS.Cdate3(drCC("FTDate"))
        ctname.Text = If(Convert.ToString(drCC("CTName")) <> "", Convert.ToString(drCC("CTName")), "無")
        tnum.Text = Convert.ToString(drCC("TNum")) '開班人數

        Dim flag_close_1_imp As Boolean = False '班級結訓判斷1
        If drCC("FTDate") <= Now.AddDays(-1) AndAlso Convert.ToString(drCC("IsClosed")) = "Y" Then
            flag_close_1_imp = True '班級結訓判斷1
        End If
        If flag_close_1_imp Then
            '結訓-停用
            'button2.Enabled=False
            'button2.ToolTip="結訓班級不可以新增學員"
            button7.Enabled = False
            button7.ToolTip = "結訓班級不可以匯入學員"
        End If

        If Not flag_close_1_imp Then
            '尚未結訓
            Select Case sm.UserInfo.TPlanID
                Case TIMS.Cst_TPlanID15 '學習卷計畫
                    If sm.UserInfo.RoleID = 1 Then
                        'button2.Enabled=True
                        button7.Enabled = True
                    Else
                        '停用
                        'button2.Enabled=False
                        'TIMS.Tooltip(button2, "學習券非中心使用者，停用")
                        button7.Enabled = False
                        'TIMS.Tooltip(button7, "學習券非中心使用者，停用")
                        TIMS.Tooltip(button7, "學習券非分署使用者，停用")
                    End If

                    'Case TIMS.Cst_TPlanID54AppPlan
                    'trImport1.Visible=True '啟用顯示
                    'button7.Enabled=True '可匯入
                    'TIMS.Tooltip(button7, "充電起飛計畫（在職），啟用", True)

                Case Else
                    '其它計畫-職前／產投 
                    'button2.ToolTip="按此按鈕可以進行新增學員的動作"
                    'button7.ToolTip="按此按鈕可以進行匯入學員的動作"
                    TIMS.Tooltip(button7, "按此按鈕可以進行匯入學員的動作", True)
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        '停用
                        'button2.Enabled=False '不可新增
                        'TIMS.Tooltip(button2, "產業人才投資方案，停用")
                        button7.Enabled = False '不可匯入
                        TIMS.Tooltip(button7, "產業人才投資方案，停用")
                    Else
                        button7.Enabled = True
                        TIMS.Tooltip(button7, "匯入啟用", True)
                    End If
            End Select
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            '其它計畫沒有-資料審核鈕
            edit_but.Visible = False '資料審核鈕
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim flag_STUD_CHK As Boolean = CHK_STUD_APPLIEDRESULT(OCIDValue1.Value)
            '學員資料審核鈕 - AppliedResultR
            If Convert.ToString(drCC("AppliedResultR")) = "Y" AndAlso flag_STUD_CHK Then
                edit_but.Enabled = False
                edit_but.ToolTip = "此班級學員資料已審核"
            Else
                edit_but.Enabled = True
                edit_but.ToolTip = "此班級學員資料尚未審核"
            End If
        End If

        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) ' .SelectedValue
        '1:模糊顯示 2:可正常顯示個資。
        flgCIShow = If(v_rblWorkMode = TIMS.cst_wmdip2, True, False)
        ViewState(cst_flgCIShow) = flgCIShow

        drOCID = Nothing
        If OCIDValue1.Value <> "" Then
            drOCID = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
            If drOCID Is Nothing Then
                Common.MessageBox(Me, "班級查詢有誤。")
                Exit Sub
            End If
            'If Convert.ToString(drOCID("ShowOK14"))="Y" Then flgCIShow=True '可正常顯示個資。
        End If

        Dim sql As String = ""
        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.StudentID" & vbCrLf
        sql &= " ,a.StudStatus" & vbCrLf
        sql &= " ,a.IsApprPaper" & vbCrLf
        sql &= " ,a.SubsidyID" & vbCrLf
        sql &= " ,a.levelNo" & vbCrLf
        sql &= " ,a.BudgetID" & vbCrLf '預算別判斷
        sql &= " ,a.SupplyID" & vbCrLf
        sql &= " ,(SELECT MAX(ENTERPATH) FROM V_ENTERTYPE2 WHERE OCID=a.OCID AND IDNO=b.IDNO AND Birthday=b.Birthday) ENTERPATH" & vbCrLf
        sql &= " ,a.EnterChannel" & vbCrLf '報名管道
        sql &= " ,a.MidentityID" & vbCrLf
        sql &= " ,a.IdentityID" & vbCrLf
        'sql &= " ,a.ActNo" & vbCrLf
        'sql &= " ,d.ActNo,d.ACTNAME" & vbCrLf
        sql &= " ,d.ActNo" & vbCrLf
        sql &= " ,a.BudgetID BudId" & vbCrLf
        sql &= " ,a.WorkSuppIdent" & vbCrLf
        sql &= " ,b.SID" & vbCrLf
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,UPPER(b.IDNO) IDNO" & vbCrLf
        sql &= " ,b.Sex" & vbCrLf
        sql &= " ,FORMAT(b.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " ,b.MilitaryID" & vbCrLf
        sql &= " ,b.EngName" & vbCrLf
        sql &= " ,b.DegreeID" & vbCrLf
        sql &= " ,b.IsAgree" & vbCrLf
        sql &= " ,c.school" & vbCrLf
        sql &= " ,c.Department" & vbCrLf
        sql &= " ,c.ServiceID" & vbCrLf
        sql &= " ,c.MilitaryRank" & vbCrLf
        sql &= " ,c.ServiceOrg" & vbCrLf
        sql &= " ,c.ServicePhone" & vbCrLf
        sql &= " ,c.SServiceDate" & vbCrLf
        sql &= " ,c.FServiceDate" & vbCrLf
        sql &= " ,c.PhoneD" & vbCrLf
        sql &= " ,c.PhoneN" & vbCrLf
        sql &= " ,c.CellPhone" & vbCrLf
        sql &= " ,c.address" & vbCrLf
        sql &= " ,c.EmergencyContact" & vbCrLf
        sql &= " ,c.EmergencyRelation" & vbCrLf
        sql &= " ,c.EmergencyAddress" & vbCrLf
        sql &= " ,c.Email" & vbCrLf
        sql &= " ,c.ZipCode3" & vbCrLf
        sql &= " ,c.ShowDetail" & vbCrLf
        sql &= " ,d.ServDept ,d.JobTitle ,d.Addr ,d.Tel" & vbCrLf
        sql &= " ,e.Q1 ,e.Q4 ,e.Q61 ,e.Q62 ,e.Q63 ,e.Q64" & vbCrLf
        sql &= " ,b.MaritalStatus" & vbCrLf '婚姻狀況
        sql &= " ,CONVERT(VARCHAR, a.RejectTDate1, 111) RejectTDate1" & vbCrLf '離訓日期
        sql &= " ,CONVERT(VARCHAR, a.RejectTDate2, 111) RejectTDate2" & vbCrLf '退訓日期
        sql &= " ,a.HighEduBg" & vbCrLf '加特別預算判斷
        sql &= " ,a.PWType1" & vbCrLf '受訓前任職狀況
        sql &= " ,a.RejectDayIn14" & vbCrLf
        sql &= " ,a.RejectSOCID" & vbCrLf
        sql &= " ,a.MakeSOCID" & vbCrLf '(兩週內)離退訓'遞補學員'被遞補學員
        sql &= " ,FORMAT(cc.STDATE ,'yyyy/MM/dd') STDATE" & vbCrLf

        '如果是產學訓則離退訓學員不顯示出來
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '**by Milor 20080711--產投學號直接取最後兩碼 start
            '將學號已去除固定字串後轉數值化，以利排序，為避免有不符合規則的學號存在，所以轉值前先進行數值化判斷
            '若長度為12碼，則試著取後3碼看是否為數值
            '若不是則取2碼，得學員學號，其餘部份為異常，需在使用程式修改
            sql &= " ,CASE WHEN ISNUMERIC(RIGHT(a.StudentID,3))=1 AND LEN(a.StudentID)=12 THEN CONVERT(NUMERIC,RIGHT(a.StudentID,3))" & vbCrLf
            sql &= "  WHEN ISNUMERIC(RIGHT(a.StudentID,2))=1 THEN CONVERT(NUMERIC,RIGHT(a.StudentID,2)) END StudID" & vbCrLf
        Else
            ''**by Milor 20080530 start
            ''將學號已去除固定字串後轉數值化，以利排序，為避免有不符合規則的學號存在，所以轉值前先進行數值化判斷
            'sql += " ,CASE WHEN ISNUMERIC(REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,'')) =1" & vbCrLf
            'sql += "            AND LEN(REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,'')) <= 3" & vbCrLf
            'sql += " 	    THEN CONVERT(NUMERIC, REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,''))" & vbCrLf
            'sql += " 	    ELSE CONVERT(NUMERIC, dbo.SUBSTR(a.StudentID,-3))" & vbCrLf
            'sql += "       END AS StudID" & vbCrLf
            ''**by Milor 20080530 end
            '將學號已去除固定字串後轉數值化，以利排序，為避免有不符合規則的學號存在，所以轉值前先進行數值化判斷
            sql &= " ,CASE WHEN LEN(REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,''))=3 AND ISNUMERIC(REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,'')) =1 THEN REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,'')" & vbCrLf
            sql &= "  WHEN LEN(REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,''))=2 AND ISNUMERIC(REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,'')) =1 THEN REPLACE(a.StudentID,cc.Years+'0'+s3.ClassID+cc.CyclType,'')" & vbCrLf
            sql &= "  ELSE RIGHT(a.StudentID,2) END StudID" & vbCrLf
        End If

        sql &= " FROM dbo.CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN dbo.CLASS_CLASSINFO cc ON a.OCID=cc.OCID" & vbCrLf
        sql &= " JOIN dbo.ID_CLASS s3 ON s3.CLSID=cc.CLSID" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip ON ip.PLANID= cc.PLANID" & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO b ON a.SID=b.SID" & vbCrLf
        sql &= " JOIN dbo.STUD_SUBDATA c ON a.SID=c.SID" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_SERVICEPLACE d ON a.SOCID=d.SOCID" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_TRAINBG e ON a.SOCID=e.SOCID" & vbCrLf
        'Stud_TrainBGQ2 (多筆)
        sql &= " WHERE ip.TPLANID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql &= " AND ip.YEARS='" & sm.UserInfo.Years & "'" & vbCrLf
        Select Case sm.UserInfo.LID
            Case 0
            Case 1
                sql &= " AND ip.DISTID ='" & sm.UserInfo.DistID & "'" & vbCrLf
            Case Else
                sql &= " AND ip.DISTID ='" & sm.UserInfo.DistID & "'" & vbCrLf
                sql &= " AND cc.RID ='" & sm.UserInfo.RID & "'" & vbCrLf
        End Select

        sql &= " AND a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        '被遞補學員 不為空(表示已被遞補) 維護並無實際意義
        'sql += " AND a.MakeSOCID IS NULL" & vbCrLf
        '20090410(Milor)加入只能查登入年度的年度限制。
        sql &= " AND cc.Years='" & Right(sm.UserInfo.Years, 2) & "'" & vbCrLf

        '2011改為全部顯示不管是否離退
        'If sm.UserInfo.TPlanID="28" Then sql &= " AND a.StudStatus NOT IN (2,3)" & vbCrLf '如果是產學訓則離退訓學員不顯示出來

        ViewState("SD03002_SearchSqlStr") = sql
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)

        '使用用勞保勾稽,並且顯示預算別
        If (Hid_show_actno_budid.Value = "Y") Then
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                Dim pRMS28 As New Hashtable From {{"OCID", OCIDValue1.Value}}
                Dim sql28 As String = ""
                sql28 &= " SELECT a.IDNO,a.ACTNO,cs.SOCID,cc.OCID" & vbCrLf
                sql28 &= " FROM dbo.STUD_BLIGATEDATA28 a WITH(NOLOCK)" & vbCrLf
                sql28 &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.SOCID=a.SOCID" & vbCrLf
                sql28 &= " JOIN CLASS_CLASSINFO cc on cc.OCID=cs.OCID" & vbCrLf
                sql28 &= " WHERE cc.OCID=@OCID" & vbCrLf
                '開訓日後才顯示勾稽結果
                sql28 &= " AND convert(date,GETDATE()) >= cc.STDATE" & vbCrLf
                Dim dt28 As DataTable = DbAccess.GetDataTable(sql28, objconn, pRMS28)

                If dt1.Rows.Count > 0 AndAlso dt28.Rows.Count > 0 Then
                    For Each dr1 As DataRow In dt1.Rows
                        Dim ff3 As String = String.Format("idno='{0}'", Convert.ToString(dr1("idno")))
                        If dt28.Select(ff3).Length > 0 Then
                            Dim s_ActNo As String = Convert.ToString(dt28.Select(ff3)(0)("ActNo")) 'STUD_BLIGATEDATA28
                            If (s_ActNo <> "" AndAlso Convert.ToString(dr1("ActNo")) = "") Then dr1("ActNo") = s_ActNo
                        End If
                    Next
                End If
            Else
                Dim pRMS6 As New Hashtable From {{"OCID", OCIDValue1.Value}}
                Dim sql6 As String = ""
                sql6 &= " SELECT a.IDNO,a.ACTNO,a.OCID" & vbCrLf
                sql6 &= " FROM dbo.STUD_SELRESULTBLI a WITH(NOLOCK)" & vbCrLf
                sql6 &= " JOIN dbo.CLASS_CLASSINFO cc on cc.OCID=a.OCID" & vbCrLf
                sql6 &= " WHERE a.OCID=@OCID" & vbCrLf
                '開訓日後才顯示勾稽結果
                sql6 &= " AND convert(date,GETDATE()) >= cc.STDATE" & vbCrLf
                Dim dtSBL06 As DataTable = DbAccess.GetDataTable(sql6, objconn, pRMS6)

                If dt1.Rows.Count > 0 AndAlso dtSBL06.Rows.Count > 0 Then
                    For Each dr1 As DataRow In dt1.Rows
                        Dim ff3 As String = String.Format("idno='{0}'", Convert.ToString(dr1("idno")))
                        Dim s_ActNo As String = ""
                        If dtSBL06.Select(ff3).Length > 0 Then
                            s_ActNo = Convert.ToString(dtSBL06.Select(ff3)(0)("ActNo")) 'STUD_BLIGATEDATA28
                            If (s_ActNo <> "" AndAlso Convert.ToString(dr1("ActNo")) = "") Then dr1("ActNo") = s_ActNo
                        End If
                        If flagTPlanID06Plan3 Then
                            Dim out_POSITION As String = ""
                            Dim flag_SRSOLDIERS As Boolean = False '是否為屆退官兵/受訓官兵名冊
                            'sm.UserInfo.DistID / CONVERT.TOSTRING(drC("DISTID"))
                            flag_SRSOLDIERS = TIMS.CheckRESOLDER(objconn, Convert.ToString(dr1("idno")), drCC("DISTID"), Convert.ToString(dr1("STDATE")), out_POSITION)
                            If (flag_SRSOLDIERS AndAlso Convert.ToString(dr1("ActNo")) = "") Then dr1("ActNo") = cst_Serviceman '在役軍人
                        End If
                    Next
                End If
            End If
        End If
        CPdt = dt1.Copy()

        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt1, "STUDID,NAME,IDNO,SEX,BIRTHDAY,ENTERCHANNEL,STUDSTATUS,ACTNO,BUDID")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, v_rblWorkMode, OCIDValue1.Value, "", objconn, v_INQUIRY, dt1.Rows.Count, vRESDESC)

        msg.Text = cst_msgNoStdData '"查無學生資料!"
        DataGridTable.Style.Item("display") = "none"
        'Dim v_rblWorkMode As String=TIMS.GetListValue(rblWorkMode)
        If dt1.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Style.Item("display") = ""

            stdnum.Text = Val(dt1.Rows.Count) '學員人數

            PageControler1.PageDataTable = dt1
            PageControler1.Sort = "StudID"  '**by Milor 20080530--由原本的StudentID排序，改為StudID排序
            PageControler1.PrimaryKey = "SOCID"
            PageControler1.ControlerLoad()

            Dim pms1 As New Hashtable From {{"OCID", CInt(OCIDValue1.Value)}}
            sql = ""
            sql &= " SELECT 'x' "
            sql &= " FROM CLASS_STUDENTSOFCLASS WHERE 1=1 "
            sql &= " AND STUDSTATUS NOT IN (2,3)" & vbCrLf '排除離退訓時間又到了 BY AMU 20091006
            sql &= " AND OCID=@OCID" 'sql &= " AND (IsApprPaper IS NUll OR IsApprPaper =' ') " 'C已確認'Y審核通過
            sql &= " AND ISAPPRPAPER IS NULL " 'C已確認'Y審核通過
            'NULL、空白未確認
            Dim drS As DataRow = DbAccess.GetOneRow(sql, objconn, pms1)
            '學員資料確認鈕
            Button11.ToolTip = "委訓單位進行學員資料確認" 'NULL未確認
            If drS Is Nothing Then
                TIMS.Tooltip(Button11, "資料已確認")
                Button11.Enabled = False
            Else
                TIMS.Tooltip(Button11, "尚未確認")
                Button11.Enabled = True
            End If
        End If

        If Not flag_show_actno_budid Then
            DataGrid1.Columns(cst_保險證號).Visible = False '保險證號
            DataGrid1.Columns(cst_預算別).Visible = False '預算別代碼
        End If

    End Sub

    '檢核是否有未審通過的學員-排除離退
    Function CHK_STUD_APPLIEDRESULT(ByVal OCID As String) As Boolean
        Dim rst As Boolean = True '全數-審通過的學員
        Dim parsm As New Hashtable From {{"OCID", OCID}}
        Dim sql As String = ""
        sql &= " SELECT 'X'"
        sql &= " FROM CLASS_STUDENTSOFCLASS cs"
        sql &= " JOIN STUD_STUDENTINFO ss ON ss.SID=cs.SID"
        sql &= " WHERE cs.OCID =@OCID AND cs.StudStatus NOT IN (2,3) "
        sql &= " and isnull(cs.AppliedResult,'N')!='Y'"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parsm)
        If dt.Rows.Count > 0 Then rst = False '沒有全數-審通過的學員
        Return rst
    End Function

    ''' <summary>
    ''' 列印資料卡 (in_class_stud)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
#Region "(No Use)"
        'Dim myitem As DataGridItem
        'Dim objCheckbox As CheckBox
        'Dim StudentIDstr As String
        'Dim newStudentID As String
        'For Each myitem In Me.DataGrid1.Items
        '    objCheckbox=myitem.FindControl("Checkbox1")
        '    If objCheckbox.Checked Then
        '        StudentIDstr=StudentIDstr & Convert.ToString("\'" & myitem.Cells(8).Text & "\'" & ",")
        '    End If
        'Next
        'newStudentID=Mid(StudentIDstr, 1, StudentIDstr.Length - 1)
        'Dim cGuid As String=  ReportQuery.GetGuid(Page)
        'Dim Url As String=  ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript="<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=in_class_stud&path=TIMS&StudentID=" & newStudentID & "&OCID=" & OCIDValue1.Value & "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
#End Region
    End Sub

    ''' <summary>
    ''' 匯出資料查詢
    ''' </summary>
    ''' <returns></returns>
    Function SEARCH_DATA1_dt() As DataTable
        Dim hPMS As New Hashtable From {{"OCID", Val(OCIDValue1.Value)}}
        Dim sql As String = ""
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,b.SOCID ,b.StudentID ,c.Name ,c.EngName" & vbCrLf '中文姓名/LastName/FirstName
        sql &= " ,(SELECT x.STUDID FROM V_STUDENTINFO x WHERE x.SOCID =B.SOCID) STUDID" & vbCrLf
        sql &= " ,UPPER(c.IDNO) IDNO" & vbCrLf '身分證字號
        sql &= " ,c.Sex" & vbCrLf '性別
        sql &= " ,dbo.DECODE6(c.Sex,'M','男','F','女',c.Sex) SexName" & vbCrLf '性別
        sql &= " ,c.PassPortNO" & vbCrLf '身分別
        sql &= " ,dbo.DECODE6(c.PassPortNO,1,'本國',2,'外籍','') PassPortName" & vbCrLf '身分別
        sql &= " ,c.ChinaOrNot" & vbCrLf '非本國人身份別
        sql &= " ,dbo.DECODE(c.ChinaOrNot,1,'是','否') ChinaOrNotName" & vbCrLf '非本國人身份別
        sql &= " ,c.Nationality" & vbCrLf '原屬國籍
        sql &= " ,c.PPNO" & vbCrLf '護照或工作證號
        sql &= " ,dbo.DECODE6(c.PPNO,1,'護照號碼',2,'居留證號','') PPNOName" & vbCrLf '護照或工作證號
        sql &= " ,FORMAT(c.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf '出生日期
        sql &= " ,c.DegreeID ,e.DegreeName" & vbCrLf '最高學歷
        sql &= " ,c.MaritalStatus" & vbCrLf '婚姻狀況
        sql &= " ,dbo.DECODE6(c.MARITALSTATUS,1,'已婚',2,'未婚','暫不提供') MaritalStatusName" & vbCrLf '婚姻狀況
        sql &= " ,d.School,d.Department" & vbCrLf '學校名稱'科系
        sql &= " ,f.GradID ,f.GradName" & vbCrLf '畢業狀況
        sql &= " ,g.MilitaryID ,g.MilitaryName" & vbCrLf '兵役代碼 '兵役
        sql &= " ,d.ServiceID ,d.MilitaryAppointment ,d.MilitaryRank" & vbCrLf '軍種/兵役職務/階級 
        sql &= " ,d.ServiceOrg ,d.ChiefRankName" & vbCrLf '服役單位名稱 /主管階級名稱 
        sql &= " ,d.ServicePhone ,d.ServiceAddress ,d.SServiceDate ,d.FServiceDate" & vbCrLf '通信電話/通信地址 /服役日期起 /服役日期訖

        sql &= " ,dbo.FN_GET_ZIPCODE(d.ZipCode4,d.ZipCode4_6W) ZipCode4" & vbCrLf '郵遞區號4
        sql &= " ,q.CTName4 ,p.ZipName4 ,d.ServiceAddress Address4" & vbCrLf '通信地址
        sql &= " ,c.JobState" & vbCrLf '就職狀況
        sql &= " ,CASE CONVERT(VARCHAR, c.JobState) WHEN '1' THEN '在職' WHEN '0' THEN '失業' END JobStateName" & vbCrLf '就職狀況
        sql &= " ,d.PhoneD ,d.PhoneN ,d.CellPhone" & vbCrLf '聯絡電話(日)/聯絡電話(夜)/行動電話 

        sql &= " ,dbo.FN_GET_ZIPCODE(d.ZipCode1,d.ZipCode1_6W) ZipCode1" & vbCrLf '通訊地址/郵遞區號1 
        sql &= " ,i.CTName1 ,h.ZipName1 ,d.Address Address1" & vbCrLf '通訊地址/郵遞區號1 

        sql &= " ,dbo.FN_GET_ZIPCODE(d.ZipCode2,d.ZipCode2_6W) ZipCode2" & vbCrLf '戶籍地址/郵遞區號2
        sql &= " ,k.CTName2 ,j.ZipName2 ,d.HouseholdAddress Address2" & vbCrLf '戶籍地址/郵遞區號2

        sql &= " ,d.Email ,b.IdentityID ,b.MIdentityID" & vbCrLf 'Email /參訓身份別代碼/主要參訓身份別
        sql &= " ,l.SubsidyID ,l.SubsidyName" & vbCrLf '生活津貼代碼 /
        sql &= " ,b.OpenDate ,b.CloseDate ,b.EnterDate" & vbCrLf '開訓日期/結訓日期/報到日期
        sql &= " ,d.HandTypeID ,r.HandTypeName ,d.HandLevelID ,s.HandLevelName" & vbCrLf '障礙類別/障礙等級
        sql &= " ,d.EmergencyContact ,d.EmergencyRelation ,d.EmergencyPhone" & vbCrLf '緊急聯絡人姓名/緊急聯絡人關係/緊急聯絡人電話 

        sql &= " ,dbo.FN_GET_ZIPCODE(d.ZipCode3,d.ZipCode3_6W) ZipCode3" & vbCrLf '緊急聯絡人郵遞區號3 
        sql &= " ,o.CTName3 ,n.ZipName3 ,d.EmergencyAddress Address3" & vbCrLf

        sql &= " ,b.RejectTDate1 ,b.RejectTDate2 ,m.RTReasonID ,m.Reason" & vbCrLf '離訓日期 /退訓日期 /離退訓原因代碼  /離退訓原因
        '受訓前工作單位名稱1/職稱1/任職起日1/任職迄日1
        '受訓前工作單位名稱2/職稱2/任職起日2/任職迄日2
        sql &= " ,d.PriorWorkOrg1 ,d.Title1 ,d.SOfficeYM1 ,d.FOfficeYM1 ,d.PriorWorkOrg2 ,d.Title2 ,d.SOfficeYM2 ,d.FOfficeYM2" & vbCrLf
        '受訓前薪資/失業週數/失業週數代碼/失業週數
        sql &= " ,d.PriorWorkPay ,c.RealJobless ,c.JoblessID ,t.JoblessName" & vbCrLf
        '交通方式
        sql &= " ,d.Traffic ,CASE CONVERT(VARCHAR, d.Traffic) WHEN '1' THEN '住宿' WHEN '2' THEN '通勤' END TrafficName" & vbCrLf
        '是否供求才廠商查詢/報名階段/報名管道 
        sql &= " ,d.ShowDetail ,b.LevelNo,b.EnterChannel" & vbCrLf
        '報名管道 
        sql &= " ,CASE CONVERT(VARCHAR, b.EnterChannel) WHEN '1' THEN '網路' WHEN '2' THEN '現場' WHEN '3' THEN '通訊' WHEN '4' THEN '推介' END EnterChannelName" & vbCrLf
        '推介種類/職訓卷種類
        sql &= " ,b.TRNDMode ,CASE CONVERT(VARCHAR, b.TRNDMode) WHEN '1' THEN '職訓券' WHEN '2' THEN '學習券' WHEN '3' THEN '推介券' END TRNDModeName" & vbCrLf
        '推介種類/職訓卷種類
        sql &= " ,b.TRNDType ,CASE CONVERT(VARCHAR, b.TRNDType) WHEN '1' THEN '甲式' WHEN '2' THEN '乙式' END TRNDTypeName" & vbCrLf

        sql &= " ,b.BudgetID ,u.BudName ,c.IsAgree ,b.PMode" & vbCrLf
        sql &= " ,d.ForeName ,d.ForeTitle ,d.ForeSex" & vbCrLf
        sql &= " ,CASE d.ForeSex WHEN 'M' THEN '男' WHEN 'F' THEN '女' END ForeSexName" & vbCrLf
        sql &= " ,d.ForeBirth ,UPPER(d.ForeIDNO) ForeIDNO" & vbCrLf

        sql &= " ,dbo.FN_GET_ZIPCODE(d.ForeZip,d.ForeZip6W) ForeZip" & vbCrLf
        sql &= " ,v.ZipName ForeZipName" & vbCrLf
        sql &= " ,d.ForeAddr ,kn.KNID ,kn.Name knName" & vbCrLf
        sql &= " ,ISNULL(v2.sName,' ') SupplyID" & vbCrLf
        sql &= " ,d.HandTypeID2" & vbCrLf
        sql &= " ,d.HandLevelID2" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO a" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON a.OCID=b.OCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO c ON b.SID=c.SID" & vbCrLf
        sql &= " JOIN STUD_SUBDATA d ON c.SID=d.SID" & vbCrLf
        sql &= " LEFT JOIN (SELECT DegreeID ,Name DegreeName FROM Key_Degree) e ON c.DegreeID=e.DegreeID" & vbCrLf
        sql &= " LEFT JOIN (SELECT GradID ,Name GradName FROM Key_GradState) f ON c.GraduateStatus=f.GradID" & vbCrLf
        sql &= " LEFT JOIN (SELECT MilitaryID ,Name MilitaryName FROM Key_Military) g ON c.MilitaryID=g.MilitaryID" & vbCrLf
        sql &= " LEFT JOIN (SELECT ZipCode ,ZipName ZipName1 ,CTID FROM ID_ZIP) h ON h.ZipCode=d.ZipCode1" & vbCrLf
        sql &= " LEFT JOIN (SELECT CTName CTName1 ,CTID FROM ID_City) i ON h.CTID=i.CTID" & vbCrLf
        sql &= " LEFT JOIN (SELECT ZipCode ,ZipName ZipName2 ,CTID FROM ID_ZIP) j ON j.ZipCode=d.ZipCode2" & vbCrLf
        sql &= " LEFT JOIN (SELECT CTName CTName2 ,CTID FROM ID_City) k ON j.CTID=k.CTID" & vbCrLf
        sql &= " LEFT JOIN (SELECT SubsidyID ,Name SubsidyName FROM Key_Subsidy) l ON b.SubsidyID=l.SubsidyID" & vbCrLf
        sql &= " LEFT JOIN (SELECT RTReasonID ,Reason FROM Key_RejectTReason) m ON b.RTReasonID=m.RTReasonID" & vbCrLf
        sql &= " LEFT JOIN (SELECT ZipCode ,ZipName ZipName3 ,CTID FROM ID_ZIP) n ON n.ZipCode=d.ZipCode3" & vbCrLf
        sql &= " LEFT JOIN (SELECT CTName CTName3 ,CTID FROM ID_City) o ON n.CTID=o.CTID" & vbCrLf
        sql &= " LEFT JOIN (SELECT ZipCode ,ZipName ZipName4 ,CTID FROM ID_ZIP) p ON p.ZipCode=d.ZipCode4" & vbCrLf
        sql &= " LEFT JOIN (SELECT CTName CTName4 ,CTID FROM ID_City) q ON p.CTID=q.CTID" & vbCrLf
        sql &= " LEFT JOIN (SELECT HandTypeID ,Name HandTypeName FROM Key_HandicatType) r ON d.HandTypeID=r.HandTypeID" & vbCrLf
        sql &= " LEFT JOIN (SELECT HandLevelID ,Name HandLevelName FROM Key_HandicatLevel) s ON d.HandLevelID=s.HandLevelID" & vbCrLf
        sql &= " LEFT JOIN (SELECT JoblessID ,Name JoblessName FROM Key_JoblessWeek) t ON c.JoblessID=t.JoblessID" & vbCrLf
        sql &= " LEFT JOIN (SELECT BudID,BudName FROM Key_Budget) u ON b.BudgetID=u.BudID" & vbCrLf
        sql &= " LEFT JOIN VIEW_ZIPNAME v ON d.ForeZip=v.ZipCode" & vbCrLf
        sql &= " LEFT JOIN VIEW_SUPPLYID v2 ON v2.supplyid=b.supplyid COLLATE Chinese_Taiwan_Stroke_CI_AS" & vbCrLf
        sql &= " LEFT JOIN KEY_NATIVE kn ON kn.KNID=b.Native" & vbCrLf
        sql &= " WHERE a.OCID=@OCID" & vbCrLf
        '如果是產學訓則離退訓學員不顯示出來
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then sql &= " AND b.StudStatus NOT IN (2,3)" & vbCrLf
        sql &= " ORDER BY b.StudentID" & vbCrLf
        Using dt As DataTable = DbAccess.GetDataTable(sql, objconn, hPMS)
            Return dt
        End Using
        Return Nothing
    End Function

    ''' <summary> 匯入學員資料鈕 </summary>
    Sub Import_Data1()
        Const Cst_FileSavePath As String = "~/SD/03/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        Dim dt_xls As New DataTable

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇 職類/班級!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇 職類/班級!")
            Exit Sub
        End If

        '檢查檔案格式與大小 Start
        Const cst_FileType As String = "xls"
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, cst_FileType) Then Return

        If File1.Value = "" Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        If File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        'stella modify 2007/10/30 改為必須上傳xls類型的檔案
        If LCase(MyFileType) <> cst_FileType Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為xls檔(Excel)!")
            Exit Sub
        End If

        Dim Reason As String = "" '儲存錯誤的原因
        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{Cst_FileSavePath}{MyFileName}")
        '上傳檔案
        File1.PostedFile.SaveAs(filePath1)
        'Common.MessageBox(Me, Request.BinaryRead(File1.PostedFile.ContentLength).ToString)
        dt_xls = TIMS.GetDataTable_XlsFile(filePath1, "", Reason, "身分證字號")
        '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)),IO.File.Delete(filePath1)
        TIMS.MyFileDelete(filePath1)

        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝")
            Exit Sub
        End If
        If dt_xls.Rows.Count = 0 Then
            Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝")
            Exit Sub
        End If
        'If dt_xls.Rows.Count > 0 Then
        'End If

        '將檔案讀出放入記憶體
        Dim dt As New DataTable
        'Dim Reason As String="" '儲存錯誤的原因
        Reason = ""
        Dim RowIndex As Integer = 0
        Dim colArray As Array

        '取出資料庫的所有欄位 Start
        Dim sql As String = ""
        'Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        Dim dt99 As DataTable
        'Dim p As Integer=0 '計算Stud_SubData資料表的筆數

        '**by Milor 20080509--當為產學訓時，取出機構別 start
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            OrgKind2 = TIMS.Get_OrgKind2(OCIDValue1.Value, TIMS.c_OCID, objconn)
        End If

        Dim BasicSID As String = TIMS.Get_DateNo()
        Dim SIDNum As Integer = 1
        Dim SID As String = ""
        'Dim Reason As String ="                '儲存錯誤的原因
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow

        '建立錯誤資料格式Table - Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Name"))
        dtWrong.Columns.Add(New DataColumn("StudentID"))
        dtWrong.Columns.Add(New DataColumn("IDNO"))
        dtWrong.Columns.Add(New DataColumn("Reason"))

        '取出所有鍵值當判斷 - Start
        '根據年度每年會有不同的預算別
        sql = " SELECT BudID FROM Plan_Budget WHERE TPlanID='" & sm.UserInfo.TPlanID & "' AND Syear='" & sm.UserInfo.Years & "' ORDER BY BudID"
        Plan_Budget = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Key_Degree WHERE 1=1 AND DegreeType IN ('0','1') ORDER BY DEGREEID"
        Key_Degree = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Key_GradState ORDER BY GRADID"
        Key_GradState = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Key_Military ORDER BY MILITARYID"
        Key_Military = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Key_Identity ORDER BY IDENTITYID"
        dtIdentity = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Key_Subsidy ORDER BY SUBSIDYID"
        Key_Subsidy = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Key_HandicatType ORDER BY HANDTYPEID"
        Key_HandicatType = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Key_HandicatLevel ORDER BY HANDLEVELID"
        Key_HandicatLevel = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Key_JoblessWeek WHERE joblessid IN ('01','02','03')  ORDER BY joblessid"
        If CInt(Me.sm.UserInfo.Years) >= 2010 Then sql = " SELECT * FROM Key_JoblessWeek WHERE joblessid IN ('04','05','06') ORDER BY joblessid"
        Key_JoblessWeek = DbAccess.GetDataTable(sql, objconn)

        '企訓專用'產投檢查。
        Dim flagTPlanID28a As Boolean = False '(產投 28.54)
        Dim flagTIMSNot28a As Boolean = True '(TIMS)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            flagTPlanID28a = True
            flagTIMSNot28a = False
        End If

        '建立StudentID值
        'Dim drCC As DataRow=TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        'sql="" & vbCrLf
        'sql &= " SELECT a.Years" & vbCrLf
        'sql &= " ,ISNULL(b.ClassID2,b.ClassID) ClassID" & vbCrLf
        'sql &= " ,a.CyclType" & vbCrLf
        'sql &= " ,CONVERT(VARCHAR, a.STDATE, 111) STDATE" & vbCrLf
        'sql &= " FROM Class_ClassInfo a" & vbCrLf
        'sql &= " JOIN ID_Class b ON a.CLSID=b.CLSID" & vbCrLf
        'sql &= " WHERE a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        'dr=DbAccess.GetOneRow(sql, objconn)
        'If dr Is Nothing Then
        '    Common.MessageBox(Me, "職類/班級 有誤，請重新選擇!")
        '    Exit Sub
        'End If

        Dim s_zipcode As String = ""
        Dim s_zipcodeb3 As String = ""
        Dim s_zipcode6w As String = ""
        Dim StudentID As String = ""
        'StudentIDBasic = String.Concat(drCC("Years"), "0", drCC("ClassID"), drCC("CyclType"))
        Dim StudentIDBasic As String = Convert.ToString(drCC("BASICSTDID")) '= dr("Years").ToString & "0" & dr("ClassID").ToString & dr("CyclType").ToString
        Dim STDate As String = TIMS.Cdate3(drCC("STDATE"))

        Dim StudentIDNum As String = ""
        Dim Name As String = ""
        Dim LastName As String = ""
        Dim FirstName As String = ""
        Dim IDNO As String = ""
        Dim Sex As String = ""
        Dim Birthday As String = ""
        Dim DegreeID As String = ""
        Dim School As String = ""
        Dim Department As String = ""
        Dim GraduateStatus As String = ""
        Dim JobState As String = ""
        Dim PhoneD As String = ""
        Dim PhoneN As String = ""
        Dim CellPhone As String = ""
        Dim ZipCode1 As String = ""
        Dim ZipCode2 As String = ""
        Dim ZipCode4 As String = ""
        Dim ForeZip As String = ""
        Dim Address As String = ""
        Dim Email As String = ""
        Dim IdentityID As String = ""
        Dim MIdentityID As String = ""
        Dim OpenDate As String = ""
        Dim CloseDate As String = ""
        Dim EnterDate As String = ""
        Dim HandTypeID As String = ""
        Dim HandLevelID As String = ""
        Dim EmergencyContact As String = ""
        Dim EmergencyRelation As String = ""
        Dim EmergencyPhone As String = ""
        Dim ZipCode3 As String = ""
        Dim EmergencyAddress As String = ""
        Dim EnterChannel As String = "" '報名管道"
        Dim IsAgree As String = ""
        Dim AcctMode As String = ""
        Dim PostNo As String = ""
        Dim AcctHeadNo As String = ""
        Dim AcctExNo As String = "="""
        Dim AcctNo As String = ""
        Dim BankName As String = ""
        Dim ExBankName As String = ""
        Dim FirDate As String = ""
        Dim Uname As String = ""
        Dim Intaxno As String = ""
        Dim Tel As String = ""
        Dim Fax As String = ""
        Dim Zip As String = ""
        Dim Addr As String = ""
        Dim ServDept As String = ""
        Dim JobTitle As String = ""
        Dim SDate As String = ""
        Dim SJDate As String = ""
        Dim SPDate As String = ""
        Dim Q1 As String = ""
        Dim Q2 As String = ""
        Dim Q3 As String = ""
        Dim Q3_Other As String = ""
        Dim Q4 As String = ""
        Dim Q5 As String = ""
        Dim Q61 As String = ""
        Dim Q62 As String = ""
        Dim Q63 As String = ""
        Dim Q64 As String = ""
        Dim ShowDetail As String = ""
        Dim LevelNo As String = ""
        Dim EnterChannel_none28 As String = "" '報名管道(非產投 、非企訓)
        Dim TRNDMode As String = ""
        Dim TRNDType As String = ""
        Dim BudgetID As String = ""
        Dim Native As String = ""
        Dim SubsidyID As String = ""
        Dim PassPortNO As String = ""
        Dim ChinaOrNot As String = ""
        Dim Nationality As String = ""
        Dim PPNO As String = ""
        Dim MaritalStatus As String = "" '婚姻狀況
        Dim MilitaryID As String = ""
        Dim JoblessID As String = ""
        Dim RealJobless As String = ""
        Dim ServiceID As String = ""


        'xls 方式 讀取寫入資料庫
        'Dim k As Int32
        'If dt_xls.Rows.Count > 0 Then '有資料
        'End If
        For k As Int32 = 0 To dt_xls.Rows.Count - 1
            Reason = ""
            colArray = dt_xls.Rows(k).ItemArray
            Try
                'Reason += CheckImportData(colArray)
                '企訓專用
                If flagTPlanID28a Then Reason = CheckImportData28(colArray, StudentIDBasic)

                'sm.UserInfo.TPlanID != "28" 一般計劃專用
                If flagTIMSNot28a Then Reason = CheckImportDataTIMS(colArray, StudentIDBasic)

            Catch ex As Exception
                Call TIMS.WriteTraceLog(ex.Message, ex)
                Reason += String.Concat("欄位資料有誤，請確認資料填寫完整性！<BR>", ex.Message)
                'Common.MessageBox(Me, Reason)
                'Exit Sub
            End Try

            '通過檢查，開始輸入資料- Start
            If Reason = "" Then

                StudentIDNum = (colArray(0).ToString)
                Name = (colArray(1).ToString)
                LastName = (colArray(2).ToString)
                FirstName = (colArray(3).ToString)
                IDNO = (TIMS.ChangeIDNO("" & colArray(4).ToString))
                Sex = (colArray(5).ToString)
                'ShowDetail=colArray(60).ToString

                '非產業人才計畫才有的欄位
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Birthday = colArray(6).ToString
                    DegreeID = colArray(7).ToString
                    If colArray(8).ToString = Nothing Then School = "" Else School = colArray(8).ToString
                    If colArray(9).ToString = Nothing Then Department = "" Else Department = colArray(9).ToString
                    GraduateStatus = TIMS.Get_GraduateStatusValue(colArray(10).ToString) 'colArray(10).ToString
                    PhoneD = colArray(11).ToString
                    PhoneN = colArray(12).ToString
                    CellPhone = colArray(13).ToString
                    ZipCode1 = colArray(14).ToString
                    Address = colArray(15).ToString
                    Email = colArray(16).ToString
                    IdentityID = colArray(17).ToString
                    MIdentityID = colArray(18).ToString
                    OpenDate = colArray(19).ToString
                    CloseDate = colArray(20).ToString
                    EnterDate = colArray(21).ToString
                    HandTypeID = colArray(22).ToString
                    HandLevelID = colArray(23).ToString
                    EmergencyContact = colArray(24).ToString
                    EmergencyRelation = colArray(25).ToString
                    EmergencyPhone = colArray(26).ToString
                    ZipCode3 = colArray(27).ToString
                    EmergencyAddress = colArray(28).ToString
                    EnterChannel = colArray(29).ToString '報名管道
                    IsAgree = colArray(30).ToString
                    AcctMode = colArray(31).ToString
                    PostNo = colArray(32).ToString
                    AcctHeadNo = colArray(33).ToString
                    AcctExNo = colArray(34).ToString
                    AcctNo = colArray(35).ToString
                    BankName = colArray(36).ToString
                    ExBankName = colArray(37).ToString
                    FirDate = colArray(38).ToString
                    Uname = colArray(39).ToString
                    Intaxno = colArray(40).ToString
                    Tel = colArray(41).ToString
                    Fax = colArray(42).ToString
                    Zip = colArray(43).ToString
                    Addr = colArray(44).ToString
                    ServDept = colArray(45).ToString
                    JobTitle = colArray(46).ToString
                    SDate = colArray(47).ToString
                    SJDate = colArray(48).ToString
                    SPDate = colArray(49).ToString
                    Q1 = colArray(50).ToString
                    Q2 = colArray(51).ToString
                    Q3 = colArray(52).ToString
                    Q3_Other = colArray(53).ToString
                    Q4 = colArray(54).ToString
                    Q5 = colArray(55).ToString
                    Q61 = colArray(56).ToString
                    Q62 = colArray(57).ToString
                    Q63 = colArray(58).ToString
                    Q64 = colArray(59).ToString
                    ShowDetail = colArray(60).ToString
                Else
                    PassPortNO = colArray(6).ToString
                    ChinaOrNot = colArray(7).ToString
                    Nationality = colArray(8).ToString
                    PPNO = colArray(9).ToString
                    Birthday = colArray(10).ToString
                    MaritalStatus = colArray(11).ToString '婚姻狀況
                    DegreeID = colArray(12).ToString
                    School = colArray(13).ToString
                    Department = colArray(14).ToString
                    GraduateStatus = TIMS.Get_GraduateStatusValue(colArray(15).ToString) 'colArray(15).ToString
                    MilitaryID = colArray(16).ToString
                    ServiceID = colArray(17).ToString
                    ZipCode4 = colArray(25).ToString
                    JobState = colArray(27).ToString
                    PhoneD = colArray(28).ToString
                    PhoneN = colArray(29).ToString
                    CellPhone = colArray(30).ToString
                    ZipCode1 = colArray(31).ToString
                    Address = colArray(32).ToString
                    ZipCode2 = colArray(33).ToString
                    Email = colArray(35).ToString
                    IdentityID = colArray(36).ToString
                    MIdentityID = colArray(37).ToString
                    SubsidyID = colArray(38).ToString
                    OpenDate = colArray(39).ToString
                    CloseDate = colArray(40).ToString
                    EnterDate = colArray(41).ToString
                    HandTypeID = colArray(42).ToString
                    HandLevelID = colArray(43).ToString
                    EmergencyContact = colArray(44).ToString
                    EmergencyRelation = colArray(45).ToString
                    EmergencyPhone = colArray(46).ToString
                    ZipCode3 = colArray(47).ToString
                    EmergencyAddress = colArray(48).ToString
                    JoblessID = colArray(59).ToString
                    RealJobless = colArray(60).ToString
                    ShowDetail = colArray(61).ToString
                    LevelNo = colArray(62).ToString 'Convert.ToString(colArray(61))
                    EnterChannel_none28 = colArray(63).ToString '報名管道
                    TRNDMode = colArray(64).ToString
                    TRNDType = colArray(65).ToString
                    BudgetID = colArray(66).ToString
                    IsAgree = colArray(67).ToString
                    ForeZip = colArray(74).ToString
                    Native = colArray(76).ToString
                End If

                Dim iSOCID As Integer = 0
                '建立StudentID欄位值
                StudentID = String.Concat(StudentIDBasic, If(Int(StudentIDNum) < 10, "0", ""), Int(StudentIDNum))

                '建立SID欄位值 (身分證號+生日)
                'sql=" SELECT * FROM STUD_STUDENTINFO WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "' AND Birthday='" & Birthday & "' "
                sql = "SELECT * FROM STUD_STUDENTINFO WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "'"  '2009/07/20 改成只判斷身分證字號
                Dim drSS As DataRow = DbAccess.GetOneRow(sql, objconn)
                If drSS Is Nothing Then
                    SID = String.Concat(BasicSID, If(SIDNum < 10, "0", ""), SIDNum)
                Else
                    SID = drSS("SID")
                End If
                '假如此班無個人參加紀錄

                'Call DbAccess.Open(tConn1)
                Using TransConn1 As SqlConnection = DbAccess.GetConnection()
                    Dim Trans1 As SqlTransaction = DbAccess.BeginTrans(TransConn1)
                    Try
                        sql = "" & vbCrLf
                        sql &= " INSERT INTO CLASS_STUDENTIMP (SOCID,OCID,IDNO,MODIFYACCT,MODIFYDATE)" & vbCrLf
                        sql &= " VALUES (@SOCID,@OCID,@IDNO,@MODIFYACCT,GETDATE())" & vbCrLf
                        Dim iCmd As New SqlCommand(sql, TransConn1, Trans1)

                        '2006/03/28 add conn by matt
                        'trans=DbAccess.BeginTrans(tConn1)
                        '檢查班級學員檔是否有此人
                        '有的話跳過匯入程序
                        '沒有則新增匯入
                        sql = "SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SID='" & SID & "' AND OCID='" & OCIDValue1.Value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans1)
                        If dt.Rows.Count = 0 Then
                            iSOCID = DbAccess.GetNewId(Trans1, "CLASS_STUDENTSOFCLASS_SOCID_SE,CLASS_STUDENTSOFCLASS,SOCID")
                            Dim drCS As DataRow = dt.NewRow
                            dt.Rows.Add(drCS)
                            drCS("SOCID") = iSOCID
                            drCS("OCID") = OCIDValue1.Value
                            drCS("SID") = SID
                            drCS("StudentID") = StudentID
                            drCS("StudStatus") = 1
                            drCS("IdentityID") = TIMS.Get_IdentityIDSplitVal(IdentityID, "，")
                            drCS("MIdentityID") = If(MIdentityID.Length < 2, "0" & MIdentityID, MIdentityID)
                            drCS("EnterDate") = If(EnterDate = "", Convert.DBNull, EnterDate)
                            drCS("OpenDate") = If(OpenDate = "", STDate, OpenDate)
                            drCS("CloseDate") = If(CloseDate = "", Convert.DBNull, CloseDate)
                            '企訓專用
                            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                drCS("BudgetID") = "03"
                                '1.網;2.現;3.通;4.推
                                'dr("EnterChannel")=If(EnterChannel="", Convert.DBNull, EnterChannel)
                                drCS("EnterChannel") = If(EnterChannel = "", "1", EnterChannel) '無值為 '1.網
                                drCS("SubsidyID") = "01"
                            Else '非企訓計畫
                                drCS("LevelNo") = If(LevelNo = "", Convert.DBNull, LevelNo)
                                drCS("TRNDMode") = If(TRNDMode = "", Convert.DBNull, TRNDMode)
                                drCS("TRNDType") = If(TRNDType = "", Convert.DBNull, TRNDType)
                                'dr("EnterChannel")=If(EnterChannel_none28="", Convert.DBNull, EnterChannel_none28)
                                drCS("EnterChannel") = If(EnterChannel_none28 = "", "2", EnterChannel_none28) '無值為 '2.現
                                drCS("BudgetID") = If(BudgetID IsNot Nothing AndAlso BudgetID <> "", String.Concat(If(BudgetID.Length < 2, "0", ""), BudgetID), Convert.DBNull)
                                'by Vicient 原住民別
                                drCS("Native") = If(Native <> "", String.Concat(If(Native.Length < 2, "0", ""), Native), Convert.DBNull)
                                drCS("SubsidyID") = If(SubsidyID <> "", String.Concat(If(SubsidyID.Length < 2, "0", ""), SubsidyID), Convert.DBNull)
                            End If
                            drCS("ModifyAcct") = sm.UserInfo.UserID
                            drCS("ModifyDate") = Now
                            DbAccess.UpdateDataTable(dt, da, Trans1)

                            With iCmd
                                .Parameters.Clear()
                                .Parameters.Add("SOCID", SqlDbType.VarChar).Value = iSOCID
                                .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO
                                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                                .ExecuteNonQuery()  'edit，by:20181017
                                'DbAccess.ExecuteNonQuery(iCmd.CommandText, tConn1, iCmd.Parameters)  'edit，by:20181017
                            End With

                            'SOCID=DbAccess.GetId(trans, "CLASS_STUDENTSOFCLASS_SOCID_SE")
                            '檢查學員個人基本資料是否有此人
                            'Dim f As Integer

                            'sql=" SELECT * FROM STUD_STUDENTINFO WHERE SID='" & SID & "' "
                            sql = "SELECT * FROM STUD_STUDENTINFO WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "'"
                            dt = DbAccess.GetDataTable(sql, da, Trans1)
                            'Dim g As Integer=dt.Rows.Count - 1 'Dim drSS As DataRow=Nothing
                            If dt.Rows.Count = 0 Then
                                drSS = dt.NewRow '沒有--新增 學員個人基本資料
                                dt.Rows.Add(drSS)
                                drSS("SID") = SID
                                drSS("IDNO") = TIMS.ChangeIDNO(IDNO)
                                SIDNum += 1
                                Call UPDATE_drSS(drSS, Name, LastName, FirstName, Sex, Birthday, DegreeID, GraduateStatus, PassPortNO, ChinaOrNot, Nationality, PPNO, MaritalStatus, MilitaryID, JobState, JoblessID, RealJobless, IsAgree)
                            Else
                                'STUD_STUDENTINFO
                                For f As Integer = 0 To dt.Rows.Count - 1
                                    drSS = dt.Rows(f) '有學員個人基本資料 
                                    Call UPDATE_drSS(drSS, Name, LastName, FirstName, Sex, Birthday, DegreeID, GraduateStatus, PassPortNO, ChinaOrNot, Nationality, PPNO, MaritalStatus, MilitaryID, JobState, JoblessID, RealJobless, IsAgree)
                                Next
                            End If

                            '檢查 學員資料副檔
                            sql = " SELECT * FROM STUD_STUDENTINFO WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "' "
                            dt99 = DbAccess.GetDataTable(sql, da, Trans1)
                            For p As Integer = 0 To dt99.Rows.Count - 1
                                sql = " SELECT * FROM STUD_SUBDATA WHERE SID='" & dt99.Rows(p)("SID") & "' "
                                dt = DbAccess.GetDataTable(sql, da, Trans1)
                                Dim drSB As DataRow = Nothing
                                If dt.Rows.Count = 0 Then
                                    drSB = dt.NewRow '沒有--新增 學員資料副檔
                                    dt.Rows.Add(drSB)
                                    drSB("SID") = SID
                                Else
                                    drSB = dt.Rows(0)
                                End If

                                Name = TIMS.ClearSQM(Name)
                                School = TIMS.ClearSQM(School)
                                Department = TIMS.ClearSQM(Department)
                                drSB("Name") = Name
                                If School <> "" AndAlso School.Length > 30 Then School = Left(School, 30)
                                drSB("School") = If(School <> "", School, Convert.ToString(drSB("School")))
                                drSB("Department") = If(Department <> "", Department, Convert.ToString(drSB("Department")))
                                ZipCode1 = TIMS.ClearSQM(ZipCode1)
                                s_zipcode = If(CStr(ZipCode1).Length > 3, Left(CStr(ZipCode1), 3), ZipCode1)
                                s_zipcodeb3 = If(CStr(ZipCode1).Length > 3, TIMS.GetZIPCODEB3(ZipCode1), "")
                                s_zipcode6w = TIMS.GetZIPCODE6W(s_zipcode, s_zipcodeb3)
                                drSB("ZipCode1") = If(s_zipcode <> "", s_zipcode, Convert.DBNull)
                                drSB("ZipCode1_6W") = If(s_zipcode6w <> "", s_zipcode6w, Convert.DBNull)
                                drSB("Address") = Address

                                drSB("Email") = If(Email = "", Convert.DBNull, Email)
                                drSB("PhoneD") = If(PhoneD = "", Convert.DBNull, PhoneD)
                                drSB("PhoneN") = If(PhoneN = "", Convert.DBNull, PhoneN)
                                drSB("CellPhone") = If(CellPhone = "", Convert.DBNull, CellPhone)
                                drSB("EmergencyContact") = If(EmergencyContact = "", Convert.DBNull, EmergencyContact)
                                drSB("EmergencyRelation") = If(EmergencyRelation = "", Convert.DBNull, EmergencyRelation)
                                drSB("EmergencyPhone") = If(EmergencyPhone = "", Convert.DBNull, EmergencyPhone)
                                drSB("EmergencyAddress") = If(EmergencyAddress = "", Convert.DBNull, EmergencyAddress)

                                drSB("ShowDetail") = If(ShowDetail = "Y", ShowDetail, "N")
                                If HandTypeID <> "" AndAlso HandTypeID.Length < 2 Then HandTypeID = "0" & HandTypeID
                                drSB("HandTypeID") = If(HandTypeID <> "", HandTypeID, Convert.DBNull)
                                If HandLevelID <> "" AndAlso HandLevelID.Length < 2 Then HandLevelID = "0" & HandLevelID
                                drSB("HandLevelID") = If(HandLevelID <> "", HandLevelID, Convert.DBNull)

                                '企訓專用
                                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                    drSB("ZipCode2") = Convert.DBNull
                                    drSB("HouseholdAddress") = Convert.DBNull
                                    drSB("PriorWorkOrg1") = Convert.DBNull
                                    drSB("Title1") = Convert.DBNull
                                    drSB("SOfficeYM1") = Convert.DBNull
                                    drSB("FOfficeYM1") = Convert.DBNull
                                    drSB("SOfficeYM2") = Convert.DBNull
                                    drSB("FOfficeYM2") = Convert.DBNull
                                    drSB("PriorWorkPay") = Convert.DBNull
                                    drSB("Traffic") = Convert.DBNull
                                    drSB("ServiceID") = Convert.DBNull
                                    drSB("MilitaryAppointment") = Convert.DBNull
                                    drSB("MilitaryRank") = Convert.DBNull
                                    drSB("SServiceDate") = Convert.DBNull
                                    drSB("FServiceDate") = Convert.DBNull
                                    drSB("ServiceOrg") = Convert.DBNull
                                    drSB("ChiefRankName") = Convert.DBNull
                                    drSB("ZipCode4") = Convert.DBNull
                                    drSB("ServiceAddress") = Convert.DBNull
                                    drSB("ServicePhone") = Convert.DBNull
                                Else '非企訓計畫
                                    ZipCode2 = TIMS.ClearSQM(ZipCode2)
                                    s_zipcode = If(CStr(ZipCode2).Length > 3, Left(CStr(ZipCode2), 3), ZipCode2)
                                    s_zipcodeb3 = If(CStr(ZipCode2).Length > 3, TIMS.GetZIPCODEB3(ZipCode2), "")
                                    s_zipcode6w = TIMS.GetZIPCODE6W(s_zipcode, s_zipcodeb3)
                                    drSB("ZipCode2") = If(s_zipcode <> "", s_zipcode, Convert.DBNull)
                                    drSB("ZipCode2_6W") = If(s_zipcode6w <> "", s_zipcode6w, Convert.DBNull)
                                    drSB("HouseholdAddress") = If(colArray(34).ToString = "", Convert.DBNull, colArray(34))
                                    drSB("PriorWorkOrg1") = If(colArray(49).ToString = "", Convert.DBNull, colArray(49))
                                    drSB("Title1") = If(colArray(50).ToString = "", Convert.DBNull, colArray(50))
                                    drSB("SOfficeYM1") = If(colArray(51).ToString = "", Convert.DBNull, colArray(51))
                                    drSB("FOfficeYM1") = If(colArray(52).ToString = "", Convert.DBNull, colArray(52))
                                    drSB("PriorWorkOrg2") = colArray(53).ToString
                                    drSB("Title2") = colArray(54).ToString
                                    drSB("SOfficeYM2") = If(colArray(55).ToString = "", Convert.DBNull, colArray(55))
                                    drSB("FOfficeYM2") = If(colArray(56).ToString = "", Convert.DBNull, colArray(56))
                                    drSB("PriorWorkPay") = If(colArray(57).ToString = "", Convert.DBNull, colArray(57))
                                    drSB("Traffic") = If(colArray(60).ToString = "", Convert.DBNull, colArray(60))
                                    Select Case Val(MilitaryID)
                                        Case 4
                                            drSB("ServiceID") = If(ServiceID = "", Convert.DBNull, ServiceID)
                                            drSB("MilitaryAppointment") = If(colArray(18).ToString = "", Convert.DBNull, colArray(18))
                                            drSB("MilitaryRank") = If(colArray(19).ToString = "", Convert.DBNull, colArray(19))
                                            drSB("SServiceDate") = If(colArray(23).ToString = "", Convert.DBNull, colArray(23))
                                            drSB("FServiceDate") = If(colArray(24).ToString = "", Convert.DBNull, colArray(24))
                                            drSB("ServiceOrg") = If(colArray(20).ToString = "", Convert.DBNull, colArray(20))
                                            drSB("ChiefRankName") = If(colArray(21).ToString = "", Convert.DBNull, colArray(21))
                                            'dr("ZipCode4")=If(colArray(25).ToString="", Convert.DBNull, colArray(25))
                                            ZipCode4 = TIMS.ClearSQM(ZipCode4)
                                            If ZipCode4 <> "" Then
                                                s_zipcode = If(CStr(ZipCode4).Length > 3, Left(CStr(ZipCode4), 3), ZipCode4)
                                                s_zipcodeb3 = If(CStr(ZipCode4).Length > 3, TIMS.GetZIPCODEB3(ZipCode4), "")
                                                s_zipcode6w = TIMS.GetZIPCODE6W(s_zipcode, s_zipcodeb3)
                                                drSB("ZipCode4") = If(s_zipcode <> "", s_zipcode, Convert.DBNull)
                                                drSB("ZipCode4_6W") = If(s_zipcode6w <> "", s_zipcode6w, Convert.DBNull)
                                            End If
                                            drSB("ServiceAddress") = If(colArray(26).ToString = "", Convert.DBNull, colArray(26))
                                            drSB("ServicePhone") = If(colArray(22).ToString = "", Convert.DBNull, colArray(22))
                                        Case 2
                                            drSB("ForeName") = If(colArray(69).ToString = "", Convert.DBNull, colArray(69)) 'colArray(68).ToString
                                            drSB("ForeTitle") = If(colArray(70).ToString = "", Convert.DBNull, colArray(70)) 'colArray(69).ToString
                                            drSB("ForeSex") = If(colArray(71).ToString = "", Convert.DBNull, colArray(71)) 'colArray(70).ToString
                                            'ForeBirth
                                            If colArray(72).ToString <> "" Then
                                                If IsDate(colArray(72).ToString) Then
                                                    drSB("ForeBirth") = FormatDateTime(colArray(72).ToString, 2)
                                                Else
                                                    drSB("ForeBirth") = Convert.DBNull
                                                End If
                                            End If
                                            drSB("ForeIDNO") = If(colArray(73).ToString = "", Convert.DBNull, TIMS.ChangeIDNO(colArray(73).ToString)) 'TIMS.ChangeIDNO(colArray(72).ToString)
                                            'dr("ForeZip")=If(colArray(73).ToString="", Convert.DBNull, colArray(73)) 'colArray(73).ToString
                                            ForeZip = TIMS.ClearSQM(ForeZip)
                                            If ForeZip <> "" Then
                                                s_zipcode = If(CStr(ForeZip).Length > 3, Left(CStr(ForeZip), 3), ForeZip)
                                                s_zipcodeb3 = If(CStr(ForeZip).Length > 3, TIMS.GetZIPCODEB3(ForeZip), "")
                                                s_zipcode6w = TIMS.GetZIPCODE6W(s_zipcode, s_zipcodeb3)
                                                drSB("ForeZip") = If(s_zipcode <> "", s_zipcode, Convert.DBNull)
                                                drSB("ForeZip_6W") = If(s_zipcode6w <> "", s_zipcode6w, Convert.DBNull)
                                            End If
                                            drSB("ForeAddr") = If(colArray(75).ToString = "", Convert.DBNull, colArray(75)) 'colArray(74).ToString
                                    End Select
                                End If
                                drSB("ModifyAcct") = sm.UserInfo.UserID
                                drSB("ModifyDate") = Now
                                DbAccess.UpdateDataTable(dt, da, Trans1)
                            Next

                            'statr更新Stud_EnterTemp '計算Stud_EnterTemp 的筆數
                            sql = " SELECT * FROM STUD_ENTERTEMP WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "' "
                            Dim MyTable4 As DataTable = DbAccess.GetDataTable(sql, da, Trans1)
                            If MyTable4.Rows.Count <> 0 Then
                                For x As Integer = 0 To MyTable4.Rows.Count - 1
                                    Dim mydr4 As DataRow = MyTable4.Rows(x)
                                    mydr4("Name") = Name
                                    mydr4("Sex") = Sex
                                    mydr4("Birthday") = Birthday
                                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                        mydr4("PassPortNO") = 1
                                        'mydr4("MaritalStatus")=Convert.DBNull
                                        'mydr4("MaritalStatus")=If(MaritalStatus="", Convert.DBNull, MaritalStatus)
                                        '空白
                                        If MilitaryID.ToString <> "" Then
                                            If MilitaryID.ToString.Length < 2 Then MilitaryID = "0" & MilitaryID
                                            ff = "MilitaryID='" & MilitaryID & "'"
                                            If Key_Military.Select(ff).Length = 0 Then MilitaryID = ""
                                        End If
                                        mydr4("MilitaryID") = If(MilitaryID = "", Convert.DBNull, MilitaryID)
                                        'mydr4("MilitaryID")="03" 停用預設。
                                    Else
                                        mydr4("PassPortNO") = PassPortNO
                                        'mydr4("MaritalStatus")=If(MaritalStatus="", Convert.DBNull, MaritalStatus)
                                        'mydr4("MaritalStatus")=If(MaritalStatus="", "2", MaritalStatus)
                                        '1.已;2.未;3.暫不提供Null(預設)
                                        Select Case MaritalStatus
                                            Case "1", "2"
                                                mydr4("MaritalStatus") = MaritalStatus
                                            Case Else
                                                mydr4("MaritalStatus") = Convert.DBNull
                                        End Select
                                        '空白
                                        If MilitaryID.ToString <> "" Then
                                            If MilitaryID.ToString.Length < 2 Then MilitaryID = "0" & MilitaryID
                                            ff = "MilitaryID='" & MilitaryID & "'"
                                            If Key_Military.Select(ff).Length = 0 Then MilitaryID = ""
                                        End If
                                        mydr4("MilitaryID") = If(MilitaryID = "", Convert.DBNull, MilitaryID)
                                    End If
                                    If GraduateStatus.Length < 2 Then GraduateStatus = "0" & GraduateStatus
                                    mydr4("GradID") = TIMS.Get_GraduateStatusValue(GraduateStatus) '
                                    If DegreeID.Length < 2 Then DegreeID = "0" & DegreeID '補0
                                    mydr4("DegreeID") = DegreeID
                                    mydr4("School") = School
                                    mydr4("Department") = Department

                                    s_zipcode = If(CStr(ZipCode1).Length > 3, Left(CStr(ZipCode1), 3), ZipCode1)
                                    s_zipcodeb3 = If(CStr(ZipCode1).Length > 3, TIMS.GetZIPCODEB3(ZipCode1), "")
                                    s_zipcode6w = TIMS.GetZIPCODE6W(s_zipcode, s_zipcodeb3)
                                    mydr4("zipcode") = If(s_zipcode <> "", s_zipcode, Convert.DBNull)
                                    mydr4("ZIPCODE6W") = If(s_zipcode6w <> "", s_zipcode6w, Convert.DBNull)
                                    mydr4("Address") = Address

                                    mydr4("Phone1") = If(PhoneD = "", Convert.DBNull, PhoneD)
                                    mydr4("Phone2") = If(PhoneN = "", Convert.DBNull, PhoneN)
                                    mydr4("CellPhone") = If(CellPhone = "", Convert.DBNull, CellPhone)
                                    mydr4("Email") = If(Email = "", Convert.DBNull, Email)
                                    mydr4("IsAgree") = "Y"
                                    mydr4("ModifyAcct") = sm.UserInfo.UserID
                                    mydr4("ModifyDate") = Now
                                    DbAccess.UpdateDataTable(MyTable4, da, Trans1)
                                Next
                            End If

                            'statr更新Stud_EnterTemp2 '計算Stud_EnterTemp2 的筆數
                            sql = " SELECT * FROM STUD_ENTERTEMP2 WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "' "
                            Dim MyTable5 As DataTable = DbAccess.GetDataTable(sql, da, Trans1)
                            If MyTable5.Rows.Count <> 0 Then
                                For y As Integer = 0 To MyTable5.Rows.Count - 1
                                    Dim mydr5 As DataRow = MyTable5.Rows(y)
                                    mydr5("Name") = Name 'mydr4("Name")
                                    mydr5("Sex") = Sex 'mydr4("Sex")
                                    mydr5("Birthday") = Birthday 'mydr4("Birthday")
                                    mydr5("PassPortNO") = PassPortNO 'mydr4("PassPortNO")
                                    'MaritalStatus: 1.已;2.未;3.暫不提供Null(預設)
                                    mydr5("MaritalStatus") = If(MaritalStatus <> "", MaritalStatus, Convert.DBNull)
                                    mydr5("DegreeID") = DegreeID 'mydr4("DegreeID")
                                    If GraduateStatus.Length < 2 Then GraduateStatus = "0" & GraduateStatus
                                    mydr5("GradID") = TIMS.Get_GraduateStatusValue(GraduateStatus) '
                                    mydr5("School") = School 'mydr4("School")
                                    mydr5("Department") = Department 'mydr4("Department")
                                    '空白
                                    If MilitaryID.ToString <> "" Then
                                        If MilitaryID.ToString.Length < 2 Then MilitaryID = "0" & MilitaryID
                                        ff = "MilitaryID='" & MilitaryID & "'"
                                        If Key_Military.Select(ff).Length = 0 Then MilitaryID = ""
                                    End If
                                    mydr5("MilitaryID") = If(MilitaryID = "", Convert.DBNull, MilitaryID)

                                    s_zipcode = If(CStr(ZipCode1).Length > 3, Left(CStr(ZipCode1), 3), ZipCode1)
                                    s_zipcodeb3 = If(CStr(ZipCode1).Length > 3, TIMS.GetZIPCODEB3(ZipCode1), "")
                                    s_zipcode6w = TIMS.GetZIPCODE6W(s_zipcode, s_zipcodeb3)
                                    mydr5("zipcode") = If(s_zipcode <> "", s_zipcode, Convert.DBNull)
                                    mydr5("zipCODE6W") = If(s_zipcode6w <> "", s_zipcode6w, Convert.DBNull)
                                    mydr5("Address") = Address 'mydr4("Address")

                                    mydr5("Phone1") = If(PhoneD = "", Convert.DBNull, PhoneD) 'mydr4("Phone1")
                                    mydr5("Phone2") = If(PhoneN = "", Convert.DBNull, PhoneN) ' mydr4("Phone2")
                                    mydr5("CellPhone") = If(CellPhone = "", Convert.DBNull, CellPhone) ' mydr4("CellPhone")
                                    mydr5("Email") = If(Email = "", Convert.DBNull, Email) 'mydr4("Email")
                                    mydr5("IsAgree") = "Y"
                                    mydr5("ModifyAcct") = sm.UserInfo.UserID
                                    mydr5("ModifyDate") = Now
                                    DbAccess.UpdateDataTable(MyTable5, da, Trans1)
                                Next
                            End If

                            '企訓專用
                            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                '檢查 學員服務單位(產學)
                                sql = " SELECT * FROM STUD_SERVICEPLACE WHERE SOCID='" & iSOCID & "' "
                                dt = DbAccess.GetDataTable(sql, da, Trans1)
                                Dim drSP As DataRow = Nothing
                                If dt.Rows.Count = 0 Then
                                    drSP = dt.NewRow '沒有-新增 學員服務單位(產學)
                                    dt.Rows.Add(drSP)
                                    drSP("SOCID") = iSOCID
                                    '**by Milor 20080509--加入Flag 2代表訓練單位代轉現金 start
                                    Select Case AcctMode
                                        Case "0"
                                            drSP("AcctMode") = 0
                                        Case "1"
                                            drSP("AcctMode") = 1
                                        Case "2"  '訓練單位代轉現金
                                            drSP("AcctMode") = 2
                                    End Select

                                    drSP("PostNo") = If(PostNo = "", Convert.DBNull, PostNo)
                                    drSP("AcctHeadNo") = If(AcctHeadNo = "", Convert.DBNull, AcctHeadNo)
                                    drSP("AcctExNo") = If(AcctExNo = "", Convert.DBNull, AcctExNo)
                                    drSP("AcctNo") = AcctNo
                                    drSP("BankName") = If(BankName = "", Convert.DBNull, BankName)
                                    drSP("ExBankName") = If(ExBankName = "", Convert.DBNull, ExBankName)
                                    drSP("FirDate") = If(FirDate = "", Convert.DBNull, FirDate)
                                    drSP("Uname") = If(Uname = "", Convert.DBNull, Uname)
                                    drSP("Intaxno") = If(Intaxno = "", Convert.DBNull, Intaxno)
                                    drSP("ServDept") = If(ServDept = "", Convert.DBNull, ServDept)
                                    drSP("JobTitle") = If(JobTitle = "", Convert.DBNull, JobTitle)
                                    drSP("Zip") = Zip
                                    drSP("Addr") = Addr
                                    drSP("Tel") = Tel
                                    drSP("Fax") = If(Fax = "", Convert.DBNull, Fax)
                                    drSP("SDate") = If(SDate = "", Convert.DBNull, SDate)
                                    drSP("SJDate") = If(SJDate = "", Convert.DBNull, SJDate)
                                    drSP("SPDate") = If(SPDate = "", Convert.DBNull, SPDate)
                                    drSP("ModifyAcct") = sm.UserInfo.UserID
                                    drSP("ModifyDate") = Now
                                    DbAccess.UpdateDataTable(dt, da, Trans1)
                                End If

                                '檢查 學員參訓背景(產學)
                                sql = " SELECT * FROM STUD_TRAINBG WHERE SOCID='" & iSOCID & "' "
                                dt = DbAccess.GetDataTable(sql, da, Trans1)
                                Dim drTB As DataRow = Nothing
                                If dt.Rows.Count = 0 Then
                                    drTB = dt.NewRow '沒有--新增 學員參訓背景(產學)
                                    dt.Rows.Add(drTB)
                                    drTB("SOCID") = iSOCID
                                    drTB("Q1") = If(Q1 = "Y", 1, 0)
                                    drTB("Q3") = If(Q3 = "", Convert.DBNull, Q3)
                                    drTB("Q3_Other") = If(Q3_Other = "", Convert.DBNull, Q3_Other)
                                    drTB("Q4") = If(Q4 = "", Convert.DBNull, Q4)
                                    drTB("Q5") = If(Q5 = "", Convert.DBNull, If(Q5 = "Y", 1, 0))
                                    drTB("Q61") = If(Q61 = "", Convert.DBNull, Q61)
                                    drTB("Q62") = If(Q62 = "", Convert.DBNull, Q62)
                                    drTB("Q63") = If(Q63 = "", Convert.DBNull, Q63)
                                    drTB("Q64") = If(Q64 = "", Convert.DBNull, Q64)
                                    drTB("ModifyAcct") = sm.UserInfo.UserID
                                    drTB("ModifyDate") = Now
                                    DbAccess.UpdateDataTable(dt, da, Trans1)
                                End If

                                '刪除 學員參訓背景(產學) Q2 資料 重新填入目前的答案
                                sql = " DELETE STUD_TRAINBGQ2 WHERE SOCID='" & iSOCID & "' "
                                DbAccess.ExecuteNonQuery(sql, Trans1)
                                If Q2 <> "" Then
                                    sql = " SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID='" & iSOCID & "' "
                                    dt = DbAccess.GetDataTable(sql, da, Trans1)
                                    Dim drTB2 As DataRow = Nothing
                                    If Split(Q2, "，").Length <> 0 Then
                                        For i As Integer = 0 To Split(Q2, "，").Length - 1
                                            If dt.Select("Q2='" & Split(Q2, "，")(i) & "'").Length = 0 Then
                                                drTB2 = dt.NewRow
                                                dt.Rows.Add(drTB2)
                                                drTB2("SOCID") = iSOCID
                                                drTB2("Q2") = Split(Q2, "，")(i)
                                            End If
                                        Next
                                        DbAccess.UpdateDataTable(dt, da, Trans1)
                                    End If
                                End If
                            End If
                            DbAccess.CommitTrans(Trans1)
                            DbAccess.CloseDbConn(TransConn1)
                        End If
                    Catch ex As Exception
                        Call TIMS.WriteTraceLog(ex.Message, ex)
                        DbAccess.RollbackTrans(Trans1)
                        DbAccess.CloseDbConn(TransConn1)
                        Dim sMsg As String = ""
                        sMsg = " 資料匯入失敗，若持續發生此錯誤，請連絡系統管理者處理!!" & vbCrLf
                        sMsg = ex.ToString
                        'Common.MessageBox(Me, ex.ToString)
                        Common.MessageBox(Me, sMsg)
                        Exit Sub
                    End Try
                End Using

            Else
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)
                'drWrong("Index")=RowIndex
                drWrong("Index") = k + 1
                If colArray.Length > 5 Then
                    drWrong("Name") = colArray(1)
                    drWrong("StudentID") = colArray(0)
                    drWrong("IDNO") = TIMS.ChangeIDNO("" & colArray(4))
                    drWrong("Reason") = Reason
                End If
            End If
            'End If
            'RowIndex += 1
        Next
        'Loop

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
            Common.MessageBox(Me, "資料匯入成功")
            'Exit Sub
        Else
            Session("MyWrongTable") = dtWrong
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('SD_03_002_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            Exit Sub
        End If

        Call Search1()  '查詢按鈕 SQL
    End Sub

    ''' <summary>
    ''' 匯入資料修正
    ''' </summary>
    ''' <param name="drSS"></param>
    ''' <param name="Name"></param>
    ''' <param name="LastName"></param>
    ''' <param name="FirstName"></param>
    ''' <param name="Sex"></param>
    ''' <param name="Birthday"></param>
    ''' <param name="DegreeID"></param>
    ''' <param name="GraduateStatus"></param>
    ''' <param name="PassPortNO"></param>
    ''' <param name="ChinaOrNot"></param>
    ''' <param name="Nationality"></param>
    ''' <param name="PPNO"></param>
    ''' <param name="MaritalStatus"></param>
    ''' <param name="MilitaryID"></param>
    ''' <param name="JobState"></param>
    ''' <param name="JoblessID"></param>
    ''' <param name="RealJobless"></param>
    ''' <param name="IsAgree"></param>
    Private Sub UPDATE_drSS(ByRef drSS As DataRow, Name As String, LastName As String, FirstName As String, Sex As String, Birthday As String, DegreeID As String, GraduateStatus As String,
                            PassPortNO As String, ChinaOrNot As String, Nationality As String, PPNO As String, MaritalStatus As String, MilitaryID As String, JobState As String,
                            JoblessID As String, RealJobless As String, IsAgree As String)
        drSS("Name") = Name
        drSS("EngName") = LastName & " " & FirstName
        drSS("Sex") = Sex
        drSS("Birthday") = Birthday
        drSS("DegreeID") = If(DegreeID <> "", If(DegreeID.Length = 1, "0" & DegreeID, DegreeID), Convert.DBNull)
        If GraduateStatus.Length < 2 Then GraduateStatus = "0" & GraduateStatus
        drSS("GraduateStatus") = TIMS.Get_GraduateStatusValue(GraduateStatus)
        drSS("IdentityID") = Convert.DBNull
        '企訓專用
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            drSS("PassPortNO") = 1
            drSS("MilitaryID") = "03"
            drSS("JoblessID") = "01"
            drSS("RealJobless") = Convert.DBNull
        Else '非企訓計畫
            drSS("PassPortNO") = PassPortNO
            If PassPortNO.ToString = "2" Then
                drSS("ChinaOrNot") = ChinaOrNot
                drSS("Nationality") = Nationality
                drSS("PPNO") = PPNO
            Else
                'dr("ChinaOrNot")=If(ChinaOrNot="", Convert.DBNull, ChinaOrNot)
                'dr("Nationality")=If(Nationality="", Convert.DBNull, Nationality)
                'dr("PPNO")=If(PPNO="", Convert.DBNull, PPNO)
                drSS("ChinaOrNot") = Convert.DBNull
                drSS("Nationality") = Convert.DBNull
                drSS("PPNO") = Convert.DBNull
            End If
            'dr("MaritalStatus")=If(MaritalStatus="", Convert.DBNull, MaritalStatus)
            '1.已;2.未;3.暫不提供Null(預設)
            Select Case MaritalStatus
                Case "1", "2"
                    drSS("MaritalStatus") = MaritalStatus
                Case Else
                    drSS("MaritalStatus") = Convert.DBNull
            End Select
            If MilitaryID.ToString.Length < 2 Then
                drSS("MilitaryID") = "0" & MilitaryID
            Else
                drSS("MilitaryID") = MilitaryID
            End If

            If JobState.ToString <> "" Then drSS("JobState") = JobState
            If JoblessID.ToString <> "" Then
                If JoblessID.ToString.Length < 2 Then
                    drSS("JoblessID") = "0" & JoblessID
                Else
                    drSS("JoblessID") = JoblessID
                End If
            End If

            drSS("RealJobless") = If(RealJobless = "", Convert.DBNull, RealJobless)
        End If

        Select Case IsAgree
            Case "Y"
                drSS("IsAgree") = IsAgree
            Case Else
                drSS("IsAgree") = "N"
        End Select
        drSS("ModifyAcct") = sm.UserInfo.UserID
        drSS("ModifyDate") = Now
    End Sub

    ''' <summary>
    ''' 匯入學員資料鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button7.Click
        Call Import_Data1()
    End Sub

    '企訓專用
    Function CheckImportData28(ByRef colArray As Array, StudentIDBasic As String) As String
        Dim Reason As String = ""
        Dim sql As String = ""

        If colArray.Length < 61 Then
            'Reason += "欄位數量不正確(應該為61個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        Dim StudentIDNum As String = colArray(0).ToString
        Dim Name As String = colArray(1).ToString
        Dim LastName As String = colArray(2).ToString
        Dim FirstName As String = colArray(3).ToString
        Dim IDNO As String = TIMS.ChangeIDNO(colArray(4).ToString)
        Dim Sex As String = colArray(5).ToString
        Dim Birthday As String = colArray(6).ToString
        Dim DegreeID As String = colArray(7).ToString

        '如果是產學訓且未填學校名則預設不詳
        '如果是產學訓且未填科系名則預設不詳 add by nick
        Dim school As String = ""
        Dim Department As String = ""
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            school = colArray(8).ToString()
            Department = colArray(9).ToString
        Else
            school = "" '"不詳"
            If Convert.ToString(colArray(8)) <> "" Then school = colArray(8).ToString()
            Department = "" '"不詳"
            If colArray(9).ToString <> "" Then Department = colArray(9).ToString
        End If

        Dim GraduateStatus As String = TIMS.Get_GraduateStatusValue(colArray(10).ToString) 'colArray(10).ToString
        'Dim MilitaryID As String=colArray(11).ToString
        Dim PhoneD As String = colArray(11).ToString
        Dim PhoneN As String = colArray(12).ToString
        Dim CellPhone As String = colArray(13).ToString
        Dim ZipCode1 As String = colArray(14).ToString
        Dim Address As String = colArray(15).ToString
        Dim Email As String = colArray(16).ToString
        Dim IdentityID As String = colArray(17).ToString
        Dim MIdentityID As String = colArray(18).ToString
        Dim OpenDate As String = colArray(19).ToString
        Dim CloseDate As String = colArray(20).ToString
        Dim EnterDate As String = colArray(21).ToString
        Dim HandTypeID As String = colArray(22).ToString
        Dim HandLevelID As String = colArray(23).ToString
        Dim EmergencyContact As String = colArray(24).ToString
        Dim EmergencyRelation As String = colArray(25).ToString
        Dim EmergencyPhone As String = colArray(26).ToString
        Dim ZipCode3 As String = colArray(27).ToString
        Dim EmergencyAddress As String = colArray(28).ToString
        Dim EnterChannel As String = colArray(29).ToString '報名管道
        Dim IsAgree As String = colArray(30).ToString
        Dim AcctMode As String = colArray(31).ToString
        Dim PostNo As String = colArray(32).ToString
        Dim AcctHeadNo As String = colArray(33).ToString
        Dim AcctExNo As String = colArray(34).ToString
        Dim AcctNo As String = colArray(35).ToString
        Dim BankName As String = colArray(36).ToString
        Dim ExBankName As String = colArray(37).ToString
        Dim FirDate As String = colArray(38).ToString
        Dim Uname As String = colArray(39).ToString
        Dim Intaxno As String = colArray(40).ToString
        Dim Tel As String = colArray(41).ToString
        Dim Fax As String = colArray(42).ToString
        Dim Zip As String = colArray(43).ToString
        Dim Addr As String = colArray(44).ToString
        Dim ServDept As String = colArray(45).ToString
        Dim JobTitle As String = colArray(46).ToString
        Dim SDate As String = colArray(47).ToString
        Dim SJDate As String = colArray(48).ToString
        Dim SPDate As String = colArray(49).ToString
        Dim Q1 As String = colArray(50).ToString
        Dim Q2 As String = colArray(51).ToString
        Dim Q3 As String = colArray(52).ToString
        Dim Q3_Other As String = colArray(53).ToString
        Dim Q4 As String = colArray(54).ToString
        Dim Q5 As String = colArray(55).ToString
        Dim Q61 As String = colArray(56).ToString
        Dim Q62 As String = colArray(57).ToString
        Dim Q63 As String = colArray(58).ToString
        Dim Q64 As String = colArray(59).ToString
        Dim ShowDetail As String = colArray(60).ToString

        If StudentIDNum = "" Then
            Reason += "必須填寫學號<Br>"
        Else
            If IsNumeric(StudentIDNum) Then
                Dim MyKey As String = Int(StudentIDNum)
                If Int(MyKey) > 1000 Then
                    Reason += "學號必須為在(1~999)範圍內<BR>"
                Else
                    If Int(MyKey) < 10 Then MyKey = "0" & Int(MyKey)
                    sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCIDValue1.Value & "' AND StudentID='" & StudentIDBasic & MyKey & "' "
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                    If dr IsNot Nothing Then Reason += "學號重複<BR>"
                End If
            Else
                Reason += "學號必須為數字(1~999)<BR>"
            End If
        End If

        If Name = "" Then Reason += "必須填寫中文姓名<BR>"

#Region "(No Use)"
        '產學訓不擋以下判斷 ===== add by nick
        'If sm.UserInfo.TPlanID <> "28" Then
        '    If LastName="" Then
        '        Reason += "必須填寫英文姓名(LastName)<BR>"
        '    Else
        '        For i As Integer=0 To LastName.Length - 1
        '            If SearchEngStr.IndexOf(LastName.ToUpper.Chars(i))=-1 Then
        '                Reason += "英文姓名必須只有英文字(LastName)<BR>"
        '            End If
        '        Next
        '    End If
        '    If FirstName="" Then
        '        Reason += "必須填寫英文姓名(FirstName)<BR>"
        '    Else
        '        For i As Integer=0 To FirstName.Length - 1
        '            If SearchEngStr.IndexOf(FirstName.ToUpper.Chars(i))=-1 Then
        '                Reason += "英文姓名必須只有英文字(FirstName)<BR>"
        '            End If
        '        Next
        '    End If
        '    If EmergencyContact="" Then
        '        Reason += "必須填寫緊急通知人姓名<BR>"
        '    End If
        '    If EmergencyRelation="" Then
        '        Reason += "必須填寫緊急通知人關係<BR>"
        '    End If
        '    If EmergencyPhone="" Then
        '        Reason += "必須填寫緊急通知人電話<BR>"
        '    End If
        '    If ZipCode3="" Then
        '        Reason += "必須填寫緊急通知人地址郵遞區號<BR>"
        '    Else
        '        If IsNumeric(ZipCode3)=False Then
        '            Reason += "郵遞區號必須為數字<BR>"
        '        End If
        '    End If
        '    If EmergencyAddress="" Then
        '        Reason += "必須填寫緊急通知人地址<BR>"
        '    End If
        'End If
        ' end 快樂的產學訓結束
#End Region

        If IDNO = "" Then
            Reason += "必須填寫身分證號碼<BR>"
        Else
            If sm.UserInfo.RoleID = 1 Then
                Dim IDNOFlag As Boolean = True
                Dim EngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                If IDNO.Length <> 10 Then
                    IDNOFlag = False
                ElseIf IDNO.Chars(1) <> "1" And IDNO.Chars(1) <> "2" Then
                    IDNOFlag = False
                ElseIf EngStr.IndexOf(IDNO.ToUpper.Chars(0)) = -1 Then
                    IDNOFlag = False
                ElseIf IDNO = "A123456789" Then
                    IDNOFlag = False
                End If

                If IDNOFlag Then
                    sql = "SELECT * FROM STUD_STUDENTINFO WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "' AND SID IN (SELECT SID FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCIDValue1.Value & "') "
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                    If dr IsNot Nothing Then
                        Reason += "此班已經有相同的身分證號碼<BR>"
                    Else
                        Dim Flag As Boolean = True
                        For i As Integer = 0 To IDNOArray.Count - 1
                            If TIMS.ChangeIDNO(IDNOArray(i)) = TIMS.ChangeIDNO(IDNO) Then
                                Reason += "檔案中有相同的身分證號碼<BR>"
                                Flag = False
                            End If
                        Next
                        If Flag Then IDNOArray.Add(TIMS.ChangeIDNO(IDNO))
                    End If

                    'stella add 2007/11/02判斷是否已有報名資料
                    IDNO = TIMS.ChangeIDNO(IDNO)
                    Dim SingUp As Boolean = TIMS.CheckIfSingUp(IDNO, OCIDValue1.Value, 1, objconn)
                    If Not SingUp Then
                        Reason += "無報名資料，請先輸入此學員之報名資料！<BR>"
                        Reason += "可用「首頁>>學員動態管理>>招生作業>>報名登錄 」匯入功能<BR>"
                    End If
                Else
                    Reason += "身分證號碼錯誤!<BR>"
                End If
            Else
                If TIMS.CheckIDNO(IDNO) Then
                    sql = " SELECT * FROM STUD_STUDENTINFO WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "' AND SID IN (SELECT SID FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCIDValue1.Value & "') "
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                    If dr IsNot Nothing Then
                        Reason += "此班已經有相同的身分證號碼<BR>"
                    Else
                        Dim Flag As Boolean = True
                        For i As Integer = 0 To IDNOArray.Count - 1
                            If TIMS.ChangeIDNO(IDNOArray(i)) = TIMS.ChangeIDNO(IDNO) Then
                                Reason += "檔案中有相同的身分證號碼<BR>"
                                Flag = False
                            End If
                        Next
                        If Flag Then IDNOArray.Add(TIMS.ChangeIDNO(IDNO))
                    End If

                    'stella add 2007/11/02判斷是否已有報名資料
                    IDNO = TIMS.ChangeIDNO(IDNO)
                    Dim SingUp As Boolean = TIMS.CheckIfSingUp(IDNO, OCIDValue1.Value, 1, objconn)
                    If Not SingUp Then Reason += "無報名資料，請先輸入此學員之報名資料！<BR>"
                Else
                    Reason += "身分證號碼錯誤!請聯絡系統管理員<BR>"
                End If
            End If
        End If

        If Sex = "" Then
            Reason += "必須填寫性別<BR>"
        Else
            Select Case Sex
                Case "M", "F"
                Case Else
                    Reason += "性別代號只能是M或者是F<BR>"
            End Select
        End If

        If Reason = "" Then
            '(限定本國)
            If Not TIMS.CheckMemberSex(IDNO, Sex) Then Reason += "依身分證號判斷 性別選項 不正確<BR>"
        End If

        If Birthday = "" Then
            Reason += "必須填寫出生日期<BR>"
        Else
            If IsDate(Birthday) = False Then
                Reason += "出生日期必須是西元年格式(yyyy/mm/dd)<BR>"
            Else
                If CDate(Birthday) < "1900/1/1" Or CDate(Birthday) > "2100/1/1" Then Reason += "出生日期範圍有誤<BR>"
            End If
        End If

        If DegreeID = "" Then
            Reason += "必須填寫最高學歷<BR>"
        Else
            Dim MyKey As String = TIMS.AddZero(DegreeID, 2)
            ff = "DegreeID='" & MyKey & "'"
            If Key_Degree.Select("DegreeID='" & MyKey & "'").Length = 0 Then Reason += "學歷值有錯，不符合鍵詞<BR>"
        End If

        If school = "" Then
            Reason += "必須填寫學校名稱<BR>"
        Else
            If Len(school) > 20 Then Reason += "學校名稱 超過 系統長度範圍(20)<BR>"
        End If
        If Department = "" Then
            Reason += "必須填寫科系<BR>"
        Else
            If Len(Department) > 20 Then Reason += "科系 超過 系統長度範圍(20)<BR>"
        End If

        If GraduateStatus <> "" Then GraduateStatus = Trim(GraduateStatus)
        'Reason += "必須填寫畢業狀況<BR>"
        If GraduateStatus <> "" Then
            Dim MyKey As String = TIMS.AddZero(GraduateStatus, 2)
            If Key_GradState.Select("GradID='" & MyKey & "'").Length = 0 Then Reason += "畢業狀況有錯，不符合鍵詞<BR>"
        End If
        'If MilitaryID="" Then
        '    Reason += "必須填寫兵役狀況<BR>"
        'Else
        '    Dim MyKey As String=MilitaryID
        '    If MilitaryID.Length < 2 Then MyKey="0" & MilitaryID
        '    If Key_Military.Select("MilitaryID='" & MyKey & "'").Length=0 Then Reason += "兵役狀況有錯，不符合鍵詞<BR>"
        'End If

        If PhoneD = "" Then Reason += "必須填寫聯絡電話_日<BR>"
        If ZipCode1 = "" Then
            Reason += "必須填寫通訊地址郵遞區號<BR>"
        Else
            If IsNumeric(ZipCode1) = False Then Reason += "通訊地址郵遞區號必須要是數字<BR>"
        End If
        If Address = "" Then Reason += "必須填寫通訊地址<BR>"
        If Email = "" Then Reason += "必須填寫電子郵件帳號(如沒有請填寫""無"")<BR>"
        If IdentityID = "" Then
            Reason += "必須填寫參訓身分別<BR>"
        Else
            If IdentityID.ToString.IndexOf("，") = -1 Then
                Dim MyKey As String = TIMS.AddZero(IdentityID, 2)
                If dtIdentity.Select("IdentityID='" & MyKey & "'").Length = 0 Then Reason += "參訓身分別不符合鍵詞<BR>"
            Else
                For i As Integer = 0 To Split(IdentityID, "，").Length - 1
                    Dim MyKey As String = TIMS.AddZero(Split(IdentityID, "，")(i), 2)
                    If dtIdentity.Select("IdentityID='" & MyKey & "'").Length = 0 Then Reason += "參訓身分別不符合鍵詞<BR>"
                Next
                If Split(IdentityID, "，").Length > 5 Then Reason += "參訓身分別只能選擇5種<BR>"
            End If
        End If
        If MIdentityID = "" Then
            Reason += "必須填寫主要參訓身分別<BR>"
        Else
            Dim MyKey As String = TIMS.AddZero(MIdentityID, 2)
            If dtIdentity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                Reason += "主要參訓身分別不符合鍵詞<BR>"
            Else
                Dim flag As Boolean = False
                Dim MyArray As Array = Split(MIdentityID, "，")
                For i As Integer = 0 To MyArray.Length - 1
                    If Int(MyKey) = Int(MyArray(i)) Then flag = True
                Next
                If flag = False Then Reason += "主要參訓身分別必須在參訓身分別的身分中<BR>"
            End If
        End If
        If OpenDate <> "" Then
            If Not IsDate(OpenDate) Then Reason += "開訓日期必須為正確的日期格式<BR>"
        End If

        If Reason = "" Then
            '檢測此學員是否 可參訓 產業人才投資方案 (大於15歲者)
            If Not TIMS.Check_YearsOld15(Birthday, OpenDate) Then Reason += "學員資格 年齡不滿15歲 不符合可參訓條件<BR>"
        End If

        If CloseDate <> "" Then
            If Not IsDate(OpenDate) Then Reason += "結訓日期必須為正確的日期格式<BR>"
        End If
        If EnterDate <> "" Then
            If Not IsDate(EnterDate) Then Reason += "報到日期必須為正確的日期格式<BR>"
        End If

        If IsAgree = "" Then
            Reason += "必須填寫願意是否提供個人資料給 勞動部勞動力發展署 暨所屬機關運用(Y/N)<BR>"
        Else
            Select Case IsAgree
                Case "Y", "N"
                Case Else
                    Reason += "願意是否提供個人資料給 勞動部勞動力發展署 暨所屬機關運用必須為Y或N值<BR>"
            End Select
        End If
        If AcctMode = "" Then
            Reason += "請輸入撥款方式(0郵政,1金融,2訓練單位代轉現金)<BR>"
        Else
            Select Case AcctMode
                Case "0"
                    If PostNo = "" Then Reason += "請輸入郵政_局號<BR>"
                    If AcctNo = "" Then Reason += "請輸入帳號<BR>"
                Case "1"
                    If AcctHeadNo = "" Then Reason += "請輸入金融_總代號<BR>"
                    'mark by nick 取消金融機構分支 20060414
                    ' If AcctExNo="" Then Reason += "請輸入金融_分支代號<BR>"
                    'If ExBankName="" Then Reason += "請輸入分行名稱<BR>"
                    If AcctNo = "" Then Reason += "請輸入帳號<BR>"
                    If BankName = "" Then Reason += "請輸入銀行名稱<BR>"
                        '**by Milor 20080509--如果非產學訓或產學訓但機構別不為勞工團體時，不能匯入AcctMode=2 start
                Case "2"
                    If OrgKind2 <> "W" Then Reason += "機構別不為勞工團體時，不能填入撥款方式-2訓練單位代轉現金<BR>"
                    'If orgKind <> "10" Then Reason += "機構別不為勞工團體時，不能填入撥款方式-2訓練單位代轉現金<BR>"
                Case Else
                    Reason += "撥款方式超過參數範圍(0郵政,1金融)<BR>"
            End Select
        End If
        If FirDate <> "" Then
            If IsDate(FirDate) = False Then Reason += "第一次投保勞保日必須為正確的日期格式(YYYY/MM/DD)<BR>"
        End If
        If Tel = "" Then Reason += "請輸入公司電話<BR>"
        If Zip = "" Then
            Reason += "必須填寫公司地址郵遞區號<BR>"
        Else
            If IsNumeric(Zip) = False Then Reason += "公司地址郵遞區號必須為數字<BR>"
        End If
        If Addr = "" Then Reason += "必須填寫公司地址<BR>"
        If SDate <> "" Then
            If IsDate(SDate) = False Then Reason += "個人到任目前任職公司起日必須為正確的日期格式(YYYY/MM/DD)<BR>"
        End If
        If SJDate <> "" Then
            If IsDate(SJDate) = False Then Reason += "個人到任目前職務起日必須為正確的日期格式(YYYY/MM/DD)<BR>"
        End If
        If SPDate <> "" Then
            If IsDate(SPDate) = False Then Reason += "最近升遷日期必須為正確的日期格式(YYYY/MM/DD)<BR>"
        End If
        If Q1 = "" Then
            Reason += "是否由公司推薦參訓(Y/N值)<BR>"
        Else
            Select Case Q1
                Case "Y", "N"
                Case Else
                    Reason += "是否由公司推薦參訓必須為Y/N值<BR>"
            End Select
        End If
        If Q2 = "" Then
            Reason += "必須填寫參訓動機(1~4)<BR>"
        Else
            If Q2.IndexOf("，") = -1 Then
                If Not IsNumeric(Q2) Then
                    Reason += "參訓動機必須為數字(1~4)<BR>"
                Else
                    If Int(Q2) > 4 Or Int(Q2) < 1 Then Reason += "參訓動機範圍1~4<BR>"
                End If
            Else
                For i As Integer = 0 To Split(Q2, "，").Length - 1
                    If Not IsNumeric(Split(Q2, "，")(i)) Then
                        Reason += "參訓動機必須為數字(1~4)<BR>"
                    Else
                        If Int(Split(Q2, "，")(i)) > 4 Or Int(Split(Q2, "，")(i)) < 1 Then Reason += "參訓動機範圍1~4<BR>"
                    End If
                Next
            End If
        End If
        If Q3 <> "" Then
            If Not IsNumeric(Q3) Then
                Reason += "訓後動向必須為數字(1~3)"
            Else
                If Int(Q3) > 3 Or Int(Q3) < 1 Then Reason += "訓後動向範圍1~3<BR>"
            End If
        End If
        If Q4 = "" Then
            Reason += "必須填寫服務單位行業別<BR>"
        Else
            If Not IsNumeric(Q4) Then
                Reason += "服務單位行業別必須為數字(01~31)"
            Else
                If Int(Q4) > 31 Or Int(Q4) < 1 Then Reason += "服務單位行業別範圍01~31<BR>"
            End If
        End If
        If Q5 <> "" Then
            Select Case Q5
                Case "Y", "N"
                Case Else
                    Reason += "服務單位是否屬於中小企業只能輸入Y或N<BR>"
            End Select
        End If

        Dim tmp_Reason2 As String = ""
        Q61 = TIMS.CHECK_Q61TXTVAL("個人工作年資", Q61, tmp_Reason2)
        Q62 = TIMS.CHECK_Q61TXTVAL("在這家公司的年資", Q62, tmp_Reason2)
        Q63 = TIMS.CHECK_Q61TXTVAL("在這職位的年資", Q63, tmp_Reason2)
        Q64 = TIMS.CHECK_Q61TXTVAL("最近升遷離本職幾年", Q64, tmp_Reason2)
        If tmp_Reason2 <> "" Then Reason &= Replace(tmp_Reason2, vbCrLf, "<BR>")

        If ShowDetail = "" Then
            Reason += "必須填寫是否提供基本資料查詢<BR>"
        Else
            Select Case ShowDetail
                Case "Y", "N"
                Case Else
                    Reason += "是否提供基本資料查詢必須為Y或N值<BR>"
            End Select
        End If

        Return Reason
    End Function

    'sm.UserInfo.TPlanID != "28" 一般計劃專用
    Function CheckImportDataTIMS(ByRef colArray As Array, StudentIDBasic As String) As String
        Dim Reason As String = ""
        Dim SearchEngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
        Dim sql As String = ""

        'sm.UserInfo.TPlanID != "28" 一般計劃專用
        If colArray.Length < 76 Then
            'Reason += "欄位數量不正確(應該為76個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        Dim drOCID As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)

        If colArray(0).ToString = "" Then
            Reason += "必須填寫學號<Br>"
        Else
            If IsNumeric(colArray(0)) Then
                Dim MyKey As String = Int(colArray(0))
                If Int(MyKey) > 1000 OrElse Int(MyKey) < 1 Then
                    Reason += "學號必須為在(1~999)範圍內<BR>"
                Else
                    If Int(MyKey) < 10 Then MyKey = "0" & Int(MyKey)
                    sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCIDValue1.Value & "' AND StudentID='" & StudentIDBasic & MyKey & "' "
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                    If dr IsNot Nothing Then Reason += "學號重複<BR>"
                End If
            Else
                Reason += "學號必須為數字(1~999)<BR>"
            End If
        End If

        If colArray(1).ToString = "" Then Reason += "必須填寫中文姓名<BR>"
        If colArray(2).ToString = "" Then
            Reason += "必須填寫英文姓名(LastName)<BR>"
        Else
            For i As Integer = 0 To colArray(2).ToString.Length - 1
                If SearchEngStr.IndexOf(colArray(2).ToString.ToUpper.Chars(i)) = -1 Then Reason += "英文姓名必須只有英文字(LastName)<BR>"
            Next
        End If
        If colArray(3).ToString = "" Then
            Reason += "必須填寫英文姓名(FirstName)<BR>"
        Else
            For i As Integer = 0 To colArray(3).ToString.Length - 1
                If SearchEngStr.IndexOf(colArray(3).ToString.ToUpper.Chars(i)) = -1 Then Reason += "英文姓名必須只有英文字(FirstName)<BR>"
            Next
        End If
        If colArray(4).ToString = "" Then
            Reason += "必須填寫身分證號碼<BR>"
        Else
            Dim IDNO As String = TIMS.ChangeIDNO(colArray(4).ToString)
            If colArray(6).ToString = "1" Then
                If sm.UserInfo.RoleID <> 5 Then
                    Dim IDNOFlag As Boolean = True
                    Dim EngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                    If IDNO.Length <> 10 Then
                        IDNOFlag = False
                    ElseIf IDNO.Chars(1) <> "1" And IDNO.Chars(1) <> "2" Then
                        IDNOFlag = False
                    ElseIf EngStr.IndexOf(IDNO.ToUpper.Chars(0)) = -1 Then
                        IDNOFlag = False
                    ElseIf IDNO = "A123456789" Then
                        IDNOFlag = False
                    End If
                    If IDNOFlag = False Then Reason += "身分證號碼錯誤!<BR>"
                Else
                    If TIMS.CheckIDNO(IDNO) = False Then Reason += "身分證號碼錯誤!請聯絡系統管理員<BR>"
                End If
            End If

            sql = "" & vbCrLf
            sql &= " WITH WS1 AS (SELECT SID FROM STUD_STUDENTINFO WHERE UPPER(IDNO)=@IDNO)" & vbCrLf
            sql &= " SELECT 'X' FROM CLASS_STUDENTSOFCLASS CS" & vbCrLf
            sql &= " JOIN WS1 ON WS1.SID=CS.SID" & vbCrLf
            sql &= " WHERE 1=1 AND CS.OCID=@OCID" & vbCrLf
            Dim dtS As New DataTable
            Dim sCmd As New SqlCommand(sql, objconn)
            TIMS.OpenDbConn(objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = TIMS.ChangeIDNO(IDNO)
                .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                dtS.Load(.ExecuteReader())
            End With
            If dtS.Rows.Count > 0 Then Reason += "此班已經有相同的身分證號碼<BR>"
            If dtS.Rows.Count = 0 Then
                Dim Flag As Boolean = True
                For i As Integer = 0 To IDNOArray.Count - 1
                    If TIMS.ChangeIDNO(IDNOArray(i)) = TIMS.ChangeIDNO(IDNO) Then
                        Reason += "檔案中有相同的身分證號碼<BR>"
                        Flag = False
                    End If
                Next
                If Flag Then IDNOArray.Add(TIMS.ChangeIDNO(IDNO))
            End If

#Region "(No Use)"
            'sql="SELECT * FROM STUD_STUDENTINFO WHERE UPPER(IDNO)='" & TIMS.ChangeIDNO(IDNO) & "' and SID IN (SELECT SID FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCIDValue1.Value & "')"
            'Dim dr As DataRow=DbAccess.GetOneRow(sql, objconn)
            'If dr IsNot Nothing Then
            '    Reason += "此班已經有相同的身分證號碼<BR>"
            'Else
            '    Dim Flag As Boolean=True
            '    For i As Integer=0 To IDNOArray.Count - 1
            '        If TIMS.ChangeIDNO(IDNOArray(i))=TIMS.ChangeIDNO(IDNO) Then
            '            Reason += "檔案中有相同的身分證號碼<BR>"
            '            Flag=False
            '        End If
            '    Next
            '    If Flag Then IDNOArray.Add(TIMS.ChangeIDNO(IDNO))
            'End If
#End Region

            'stella add 2007/11/02判斷是否已有報名資料
            IDNO = TIMS.ChangeIDNO(IDNO)
            Dim SingUp As Boolean = TIMS.CheckIfSingUp(IDNO, OCIDValue1.Value, 1, objconn)
            If Not SingUp Then Reason += "無報名資料，請先輸入此學員之報名資料！<BR>"
        End If
        If colArray(5).ToString = "" Then
            Reason += "必須填寫性別<BR>"
        Else
            Select Case colArray(5).ToString
                Case "M", "F"
                Case Else
                    Reason += "性別代號只能是M或者是F<BR>"
            End Select
        End If

        If colArray(6).ToString = "" Then
            Reason += "必須填寫身分別(1~2)<BR>"
        Else
            Select Case colArray(6).ToString
                Case "1"
                Case "2"
                    If colArray(7).ToString = "" Then
                        Reason += "請輸入非本國人身分別"
                    Else
                        Select Case colArray(7).ToString
                            Case "1"
                            Case "2"
                            Case Else
                                Reason += "非本國人身分別只能輸入1(大陸人士)或2(非大陸人士)"
                        End Select
                    End If
                    If colArray(8).ToString = "" Then Reason += "請輸入原屬國籍"
                    If colArray(9).ToString = "" Then
                        Reason += "請輸入護照或工作證號"
                    Else
                        Select Case colArray(9).ToString
                            Case "1"
                            Case "2"
                            Case Else
                                Reason += "非本國人身分別只能輸入1(護照號碼)或2(工作證號)"
                        End Select
                    End If
                Case Else
                    Reason += "身分別只能輸入1(本國)或2(外籍)<BR>"
            End Select
        End If

        If Reason = "" Then
            Select Case colArray(6).ToString
                Case "1" '1(本國)
                    If Not TIMS.CheckMemberSex(colArray(4).ToString, colArray(5).ToString) Then Reason += "依身分證號判斷 性別選項 不正確<BR>"
            End Select
        End If

        If colArray(10).ToString = "" Then
            Reason += "必須填寫出生日期<BR>"
        Else
            If IsDate(colArray(10)) = False Then
                Reason += "出生日期必須是西元年格式(yyyy/mm/dd)<BR>"
            Else
                If CDate(colArray(10)) < "1900/1/1" Or CDate(colArray(10)) > "2100/1/1" Then Reason += "出生日期範圍有誤<BR>"
            End If
        End If
        If colArray(11).ToString <> "" Then
            Select Case colArray(11).ToString
                Case "1", "2"
                Case Else
                    Reason += "婚姻狀況必須是1(已婚)或2(未婚)<BR>"
            End Select
        End If
        If colArray(12).ToString = "" Then
            Reason += "必須填寫最高學歷<BR>"
        Else
            Dim MyKey As String = TIMS.AddZero(colArray(12), 2)
            If Key_Degree.Select("DegreeID='" & MyKey & "'").Length = 0 Then Reason += "學歷值有錯，不符合鍵詞<BR>"
        End If

        If Convert.ToString(colArray(13)) = "" Then
            Reason += "必須填寫學校名稱<BR>"
        Else
            If Len(Convert.ToString(colArray(13))) > 20 Then Reason += "學校名稱 超過 系統長度範圍(20)<BR>"
        End If
        If Convert.ToString(colArray(14)) = "" Then
            Reason += "必須填寫科系<BR>"
        Else
            If Len(Convert.ToString(colArray(14))) > 20 Then Reason += "科系 超過 系統長度範圍(20)<BR>"
        End If

        If colArray(15).ToString = "" Then
            Reason += "必須填寫畢業狀況<BR>"
        Else
            Dim MyKey As String = TIMS.AddZero(colArray(15), 2)
            If Key_GradState.Select("GradID='" & MyKey & "'").Length = 0 Then Reason += "畢業狀況有錯，不符合鍵詞<BR>"
        End If
        If colArray(16).ToString <> "" Then
            'Reason += "必須填寫兵役狀況<BR>"
            Dim MyKey As String = colArray(16)
            If colArray(16).ToString.Length < 2 Then MyKey = "0" & colArray(16)
            ff = "MilitaryID='" & MyKey & "'"
            If Key_Military.Select(ff).Length = 0 Then
                Reason += "兵役狀況有錯，不符合鍵詞<BR>"
            Else
                If Int(colArray(16)) = "4" Then
                    If colArray(17).ToString = "" Then Reason += "必須填寫軍種<BR>"
                    If colArray(19).ToString = "" Then Reason += "必須填寫階級<BR>"
                    If colArray(20).ToString = "" Then Reason += "必須填寫服務單位名稱<BR>"
                    If colArray(22).ToString = "" Then Reason += "必須填寫單位電話<BR>"
                    If colArray(23).ToString = "" Then
                        Reason += "必須填寫服役起日期<BR>"
                    Else
                        If IsDate(colArray(23)) = False Then
                            Reason += "服役起日期不是正確的日期格式<BR>"
                        Else
                            If CDate(colArray(23)) < "1900/1/1" Or CDate(colArray(23)) > "2100/1/1" Then Reason += "服役起日期範圍有誤<BR>"
                        End If
                    End If
                    If colArray(24).ToString = "" Then
                        Reason += "必須填寫服役迄日期<BR>"
                    Else
                        If IsDate(colArray(24)) = False Then
                            Reason += "服役迄日期不是正確的日期格式<BR>"
                        Else
                            If CDate(colArray(24)) < "1900/1/1" Or CDate(colArray(24)) > "2100/1/1" Then Reason += "服役迄日期範圍有誤<BR>"
                        End If
                    End If
                    If colArray(25).ToString <> "" Then
                        If IsNumeric(colArray(25)) = False Then
                            Reason += "服役單位地址郵遞區號必須為數字<BR>"
                        Else
                            If colArray(25).ToString.Length <> 5 Then Reason += "服役單位地址郵遞區號必須為5碼<BR>"
                        End If
                    End If
                End If
            End If
        End If

        If colArray(27).ToString = "" Then
            Reason += "必須填寫就職狀況<BR>"
        Else
            Select Case colArray(27).ToString
                Case "0", "1"
                Case Else
                    Reason += "就職狀況必須為0或1，0.表失業，1.表在職<BR>"
            End Select
        End If

        If colArray(28).ToString = "" Then Reason += "必須填寫聯絡電話_日<BR>"

        If colArray(31).ToString = "" Then
            Reason += "必須填寫通訊地址郵遞區號<BR>"
        Else
            If IsNumeric(colArray(31)) = False Then
                Reason += "通訊地址郵遞區號必須要是數字<BR>"
            Else
                If colArray(31).ToString.Length <> 5 Then Reason += "通訊地址郵遞區號必須為5碼<BR>"
            End If
        End If
        If colArray(32).ToString = "" Then Reason += "必須填寫通訊地址<BR>"

        If colArray(33).ToString <> "" Then
            If IsNumeric(colArray(33)) = False Then
                Reason += "戶籍地址郵遞區號必須要是數字<BR>"
            Else
                If colArray(33).ToString.Length <> 5 Then Reason += "戶籍地址郵遞區號必須為5碼<BR>"
            End If
        Else
            Reason += "必須填寫戶籍地址郵遞區號<BR>"
        End If
        If colArray(34).ToString = "" Then Reason += "必須填寫戶籍地址<BR>"
        If colArray(35).ToString = "" Then Reason += "必須填寫電子郵件帳號(如沒有請填寫""無"")<BR>"

        Dim all_Identity2 As String = ""

        If colArray(36).ToString = "" Then
            Reason += "必須填寫參訓身分別<BR>"
        Else
            If colArray(36).ToString.IndexOf("，") = -1 Then
                Dim MyKey As String = TIMS.AddZero(colArray(36), 2)
                If dtIdentity.Select("IdentityID='" & MyKey & "'").Length = 0 Then Reason += "參訓身分別不符合鍵詞<BR>"
            Else
                For i As Integer = 0 To Split(colArray(36), "，").Length - 1
                    Dim MyKey As String = TIMS.AddZero(Split(colArray(36), "，")(i), 2)
                    If dtIdentity.Select("IdentityID='" & MyKey & "'").Length = 0 Then Reason += "參訓身分別不符合鍵詞<BR>"
                    If dtIdentity.Select("IdentityID='" & MyKey & "'").Length > 0 Then
                        If all_Identity2 <> "" Then all_Identity2 &= ","
                        all_Identity2 &= MyKey
                    End If
                Next
                If Split(colArray(36), "，").Length > 5 Then Reason += "參訓身分別只能選擇5種<BR>"
            End If
        End If

        Dim STDATE1 As String = TIMS.Cdate3(drOCID("STDATE"))
        If Reason = "" Then
            'If flagTPlanID02Plan2 Then
            '    '屆退官兵者 (依開訓日期(系統日期)判斷)
            '    If TIMS.CheckRESOLDER(objconn, colArray(4).ToString, sm.UserInfo.DistID, STDATE1) Then
            '        If all_Identity2="" Then
            '            Reason += "此訓練學員為屆退官兵，參訓身分別不符合鍵詞<BR>"
            '        End If
            '        If all_Identity2 <> "" Then
            '            If all_Identity2.IndexOf("12")=-1 Then
            '                Reason += "此訓練學員為屆退官兵，參訓身分別不符合鍵詞<BR>"
            '                'Reason += "此訓練學員為屆退官兵，請於參訓身分別勾選！" & vbCrLf
            '            End If
            '        End If
            '    Else
            '        If all_Identity2 <> "" Then
            '            If all_Identity2.IndexOf("12") > -1 Then Reason += "此訓練學員不為屆退官兵，參訓身分別不符合鍵詞<BR>" & vbCrLf
            '        End If
            '    End If
            'End If
        End If

        If colArray(37).ToString = "" Then
            Reason += "必須填寫主要參訓身分別<BR>"
        Else
            Dim MyKey As String = TIMS.AddZero(colArray(37), 2)
            If dtIdentity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                Reason += "主要參訓身分別不符合鍵詞<BR>"
            Else
                Dim flag As Boolean = False
                Dim MyArray As Array = Split(colArray(36), "，")
                For i As Integer = 0 To MyArray.Length - 1
                    If Int(MyKey) = Int(MyArray(i)) Then flag = True
                Next
                If flag = False Then Reason += "主要參訓身分別必須在參訓身分別的身分中<BR>"
            End If
            'by Vicient
            Dim aIDNO As String = colArray(4).ToString
            Select Case MyKey
                Case "12" '依主要參訓身分別
                    'If flagTPlanID02Plan2 Then
                    '    '屆退官兵者 (依開訓日期(系統日期)判斷)
                    '    If Not TIMS.CheckRESOLDER(objconn, aIDNO, sm.UserInfo.DistID, STDATE1) Then Reason += "主要參訓身分別，學員資格不符合屆退官兵身分<BR>" & vbCrLf
                    'End If
                Case "05", "5" '依主要參訓身分別
                    If colArray(76).ToString = "" Then
                        Reason += "必須填寫原住民別<BR>"
                    Else
                        Dim MyKey1 As String = colArray(76)
                        Dim Key_Native As DataTable
                        If IsNumeric(MyKey1) Then
                            MyKey1 = CStr(CInt(MyKey1))
                            If Len(MyKey1) < 2 Then MyKey1 = "0" & MyKey1
                            sql = "SELECT * FROM KEY_NATIVE WHERE KNID='" & MyKey1 & "'"
                            Key_Native = DbAccess.GetDataTable(sql, objconn)
                            If Key_Native.Rows.Count = 0 Then Reason += "民族別有錯，不符合鍵詞<BR>"
                        Else
                            Reason += "民族別有錯，不符合鍵詞<BR>"
                        End If
                    End If
            End Select
        End If

        If colArray(38).ToString = "" Then
            Reason += "必須填寫生活津貼代碼<BR>"
        Else
            Dim MyKey As String = TIMS.AddZero(colArray(38), 2)
            If Key_Subsidy.Select("SubsidyID='" & MyKey & "'").Length = 0 Then Reason += "生活津貼不符合鍵詞<BR>"
        End If
        If colArray(39).ToString <> "" Then
            If IsDate(colArray(39)) = False Then
                Reason += "開訓日期不符合日期格式<BR>"
            Else
                If CDate(colArray(39)) < "1900/1/1" Or CDate(colArray(39)) > "2100/1/1" Then Reason += "開訓日期範圍有誤<BR>"
            End If
        End If
        If colArray(40).ToString <> "" Then
            If IsDate(colArray(40)) = False Then
                Reason += "結訓日期不符合日期格式<BR>"
            Else
                If CDate(colArray(40)) < "1900/1/1" Or CDate(colArray(40)) > "2100/1/1" Then Reason += "結訓日期範圍有誤<BR>"
            End If
        End If
        If colArray(41).ToString <> "" Then
            If IsDate(colArray(41)) = False Then
                Reason += "報到日期不符合日期格式<BR>"
            Else
                If CDate(colArray(41)) < "1900/1/1" Or CDate(colArray(41)) > "2100/1/1" Then Reason += "報到日期範圍有誤<BR>"
            End If
        End If
        If colArray(42).ToString <> "" Then
            Dim MyKey As String = TIMS.AddZero(colArray(42), 2)
            If Key_HandicatType.Select("HandTypeID='" & MyKey & "'").Length = 0 Then Reason += "障礙類別有錯，不符合鍵詞<BR>"
        End If
        If colArray(43).ToString <> "" Then
            Dim MyKey As String = TIMS.AddZero(colArray(43), 2)
            If Key_HandicatLevel.Select("HandLevelID='" & MyKey & "'").Length = 0 Then Reason += "障礙等級有錯，不符合鍵詞<BR>"
        End If
        If colArray(44).ToString = "" Then Reason += "必須填寫緊急通知人姓名<BR>"
        If colArray(45).ToString = "" Then Reason += "必須填寫緊急通知人關係<BR>"
        If colArray(46).ToString = "" Then Reason += "必須填寫緊急通知人電話<BR>"
        If colArray(47).ToString = "" Then
            Reason += "必須填寫緊急通知人地址郵遞區號<BR>"
        Else
            If IsNumeric(colArray(47)) = False Then
                Reason += "緊急通知人地址郵遞區號必須為數字<BR>"
            Else
                If colArray(47).ToString.Length <> 5 Then Reason += "緊急通知人地址郵遞區號必須為5碼<BR>"
            End If
        End If
        If colArray(48).ToString = "" Then Reason += "必須填寫緊急通知人地址<BR>"
        If colArray(51).ToString <> "" Then
            If IsDate(colArray(51)) = False Then
                Reason += "受訓前服務單位1任職起日不符合日期格式<BR>"
            Else
                If CDate(colArray(51)) < "1900/1/1" Or CDate(colArray(51)) > "2100/1/1" Then Reason += "受訓前服務單位1任職起日範圍有誤<BR>"
            End If
        End If
        If colArray(52).ToString <> "" Then
            If IsDate(colArray(52)) = False Then
                Reason += "受訓前服務單位1任職迄日不符合日期格式<BR>"
            Else
                If CDate(colArray(52)) < "1900/1/1" Or CDate(colArray(52)) > "2100/1/1" Then Reason += "受訓前服務單位1任職迄日範圍有誤<BR>"
            End If
        End If
        If colArray(55).ToString <> "" Then
            If IsDate(colArray(55)) = False Then
                Reason += "受訓前服務單位2任職起日不符合日期格式<BR>"
            Else
                If CDate(colArray(55)) < "1900/1/1" Or CDate(colArray(55)) > "2100/1/1" Then Reason += "受訓前服務單位2任職起日範圍有誤<BR>"
            End If
        End If
        If colArray(56).ToString <> "" Then
            If IsDate(colArray(56)) = False Then
                Reason += "受訓前服務單位2任職迄日不符合日期格式<BR>"
            Else
                If CDate(colArray(56)) < "1900/1/1" Or CDate(colArray(56)) > "2100/1/1" Then Reason += "受訓前服務單位2任職迄日範圍有誤<BR>"
            End If
        End If
        If colArray(57).ToString <> "" Then
            If IsNumeric(colArray(57)) = False Then Reason += "受訓前薪資必須為數字<BR>"
        End If
        If colArray(58).ToString <> "" Then
            If IsNumeric(colArray(58)) = False Then Reason += "受訓前真正失業週數必須為數字<BR>"
        End If
        If colArray(59).ToString = "" Then
            Reason += "必須填寫失業週數代碼<BR>"
        Else
            Dim MyKey As String = TIMS.AddZero(colArray(59), 2)
            If Key_JoblessWeek.Select("JoblessID='" & MyKey & "'").Length = 0 Then Reason += "失業週數代碼有錯，不符合鍵詞<BR>"
        End If
        If colArray(60).ToString <> "" Then
            Select Case colArray(60).ToString
                Case "1", "2"
                Case Else
                    Reason += "交通方式必須為1(住宿)或2(通勤)<BR>"
            End Select
        End If
        If colArray(61).ToString = "" Then
            Reason += "必須填寫是否提供基本資料查詢<BR>"
        Else
            Select Case colArray(61).ToString
                Case "Y", "y", "N", "n"
                Case Else
                    Reason += "是否提供基本資料查詢必須為Y或N值<BR>"
            End Select
        End If
        If colArray(63).ToString <> "" Then
            Select Case colArray(63).ToString
                Case "1", "2", "3"
                Case "4"
                    If colArray(64).ToString = "" Then
                        Reason += "報名管道為推介時，必須選擇卷別<BR>"
                    Else
                        Select Case colArray(64).ToString
                            Case "1", "3"
                                If colArray(65).ToString = "" Then
                                    Reason += "券別種類必須填入甲乙式<BR>"
                                Else
                                    Select Case colArray(65).ToString
                                        Case "1", "2"
                                        Case Else
                                            Reason += "券別種類只有1(甲式)2(乙式)<BR>"
                                    End Select
                                End If
                            Case "2"
                                If colArray(65).ToString <> "" Then Reason += "學習券不區分甲乙式<BR>"
                            Case Else
                                Reason += "推介種類只有1(職訓券)2(學習券)3(推介券)<BR>"
                        End Select
                    End If
                Case Else
                    Reason += "報名管道只有1(網路)2(現場)3(通訊)4(推介)<BR>"
            End Select
        End If
        If colArray(66).ToString = "" And Plan_Budget.Rows.Count <> 0 Then
            Reason += "必須填寫預算別<BR>"
        Else
            Dim MyKey As String = TIMS.AddZero(colArray(66), 2)
            If Plan_Budget.Select("BudID='" & MyKey & "'").Length = 0 Then Reason += "預算別不符合此訓練計畫<BR>"
        End If
        If colArray(67).ToString = "" Then
            Reason += "必須填寫願意是否提供個人資料給 勞動部勞動力發展署 暨所屬機關運用(Y/N)<BR>"
        Else
            Select Case colArray(67).ToString
                Case "Y", "N"
                Case Else
                    Reason += "願意是否提供個人資料給 勞動部勞動力發展署 暨所屬機關運用必須為Y或N值<BR>"
            End Select
        End If
        If colArray(68).ToString = "" Then
            If sm.UserInfo.TPlanID = "12" Then Reason += "必須填寫公費(1)或自費(2)<BR>"
        ElseIf colArray(68).ToString <> "" Then
            Select Case colArray(68).ToString
                Case "1", "2"
                Case Else
                    Reason += "公費(1)/自費(2)值超出範圍<BR>"
            End Select
        End If
        If colArray(71).ToString <> "" Then
            Select Case colArray(71).ToString
                Case "M", "F"
                Case Else
                    Reason += "國內親屬資料_性別只能輸入M(男性)F(女性)<BR>"
            End Select
        End If
        If colArray(72).ToString <> "" Then
            If Not IsDate(colArray(72).ToString) Then Reason += "國內親屬資料_生日必須為正確的日期格式<BR>"
        End If
        If colArray(73).ToString <> "" Then
            If Not TIMS.CheckIDNO(TIMS.ChangeIDNO(colArray(73).ToString)) Then Reason += "國內親屬資料_身分證號碼不是正確的身分證號碼<BR>"
        End If
        If colArray(74).ToString <> "" Then
            If Not IsNumeric(colArray(74)) Then
                Reason += "國內親屬資料_郵遞區號必須為數字<BR>"
            Else
                If colArray(74).ToString.Length <> 5 Then Reason += "國內親屬資料_郵遞區號必須為5碼<BR>"
            End If
        End If

        Return Reason
    End Function

    Sub CHK_Bli28_DataTable()
        'Dim dtBli28 As DataTable=TIMS.GET_BLIGATEDATA28dt(objconn, OCIDValue1.Value)
        'Dim dtBli28e As DataTable=TIMS.GET_BLIGATEDATA28Edt(objconn, OCIDValue1.Value)
        If dtBli28e Is Nothing Then dtBli28e = TIMS.GET_BLIGATEDATA28Edt(objconn, OCIDValue1.Value)

        'If dtBli28 Is Nothing Then
        '    dtBli28=TIMS.GET_BLIGATEDATA28dt(objconn, OCIDValue1.Value)
        'End If
        '1.在【學員資料維護】，若該學員於於「公法救助」是屬於「M：多元就業計畫進用人員不適用就保」，系統預算別要預設帶「就安」! (圖4)。
        'Hid_BIEF.Value=TIMS.GET_BLIGATEDATA28(dr("SOCID"), dr("IDNO"), objconn, "BIEF")
        'If Hid_BIEF.Value="M" Then
        '    v_def_BudgetID="02"
        'End If
        ''1.在【學員資料維護】，若該學員於於「公法救助」是屬於「M：多元就業計畫進用人員不適用就保」，系統預算別要預設帶「就安」! (圖4)。
        'Hid_BIEF.Value=TIMS.GET_BLIGATEDATA28E(dr("IDNO"), dr("OCID"), objconn, "BIEF")
        'If Hid_BIEF.Value="M" Then
        '    v_def_BudgetID="02"
        'End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim Checkbox3 As HtmlInputCheckBox = e.Item.FindControl("Checkbox3")
                Checkbox3.Attributes("onclick") = "ChangeAll(this);"

            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                'objControl=e.Item.FindControl("Checkbox1")
                Dim btn1 As LinkButton = e.Item.FindControl("Button3") '修改
                'Dim btn10Delete As LinkButton=e.Item.FindControl("btn10Delete") '刪除。
                Dim Checkbox3 As HtmlInputCheckBox = e.Item.FindControl("Checkbox3")
                Dim Checkbox2 As HtmlInputCheckBox = e.Item.FindControl("Checkbox2")
                Dim star2 As Label = e.Item.FindControl("star2") '#表示為該學員資料尚未確認或為退件修正狀態
                Dim BudID As DropDownList = e.Item.FindControl("BudID")
                Dim Hid_idno As HiddenField = e.Item.FindControl("Hid_idno")
                Hid_idno.Value = Convert.ToString(drv("idno"))
                Hid_idno.Value = TIMS.EncryptAes(Hid_idno.Value)

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    annotate.Visible = True
                    'False:不顯示 /True:顯示
                    star2.Visible = If(Convert.ToString(drv("IsApprPaper")) = "Y", False, True)
                    If Convert.ToString(drv("IsApprPaper")) <> "Y" Then
                        TIMS.Tooltip(star2, "學員資料尚未確認")
                    End If
                End If

                '保險證號/預算別代碼 false:不顯示 true:顯示
                If (Hid_show_actno_budid.Value = "Y") Then
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then BudID = TIMS.Get_Budget(BudID, 2)
                    If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then BudID = TIMS.Get_Budget(BudID, 29)
                    If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then BudID = TIMS.Get_Budget(BudID, 29)
                    Dim v_BudgetID As String = Convert.ToString(drv("BudgetID"))
                    If v_BudgetID = "" Then
                        'BudID (預算別) by AMU 20080602
                        '根據參訓學員於e網所填列之保險證號前2碼判讀, 前2碼為
                        '01、04、05、15、08 其補助經費來源歸屬為 03:就保基金
                        '02、03、06、07 其經費來源歸屬為 02:就安基金
                        '09與無法辨視者為 99:不予補助對象
                        '2.開頭數字為075、175（裁減續保）、076、176（職災續保）、09（訓）皆為不予補助對象，並設定阻擋。
                        Select Case Left(Convert.ToString(drv("ActNo")), 2) 'CLASS_STUDENTSOFCLASS
                            Case "01", "04", "05", "15", "08"
                                v_BudgetID = "03"
                            Case "02", "03", "06", "07"
                                v_BudgetID = "02"
                            Case "09"
                                v_BudgetID = "99"
                            Case Else
                                v_BudgetID = "99"
                        End Select
                        Select Case Left(Convert.ToString(drv("ActNo")), 3) 'CLASS_STUDENTSOFCLASS
                            Case "075", "175", "076", "176"
                                v_BudgetID = "99"
                        End Select
                        '在役軍人'02:就安基金
                        If Convert.ToString(drv("ActNo")).Equals(cst_Serviceman) Then v_BudgetID = "02" '02:就安基金

                        Call CHK_Bli28_DataTable()
                        '1.在【學員資料維護】，若該學員於於「公法救助」是屬於「M：多元就業計畫進用人員不適用就保」，系統預算別要預設帶「就安」! (圖4)。
                        'Dim v_BIEF_28 As String=TIMS.GET_BLIGATEDATA28(Convert.ToString(drv("SOCID")), Convert.ToString(drv("IDNO")), dtBli28, "BIEF")
                        'If v_BIEF_28="M" Then v_BudgetID="02"
                        '1.在【學員資料維護】，若該學員於於「公法救助」是屬於「M：多元就業計畫進用人員不適用就保」，系統預算別要預設帶「就安」! (圖4)。
                        Dim v_BIEF_28E As String = TIMS.GET_BLIGATEDATA28E(Convert.ToString(drv("IDNO")), Convert.ToString(drv("OCID")), dtBli28e, "BIEF")
                        If v_BIEF_28E = "M" Then v_BudgetID = "02"
                    End If
                    'CLASS_STUDENTSOFCLASS
                    If BudID IsNot Nothing Then
                        Common.SetListItem(BudID, v_BudgetID)
                    End If

                    '20090123 andy  edit 產投 2009年 身分別為「就業保險被保險人非自願失業者」時
                    '1.預算來源設定為 02:就保基金 ； 2.補助比例為100% 'start
                    'If CInt(Me.sm.UserInfo.Years) > 2008 Then
                    '    For i As Integer=0 To Split(Convert.ToString(drv("IdentityID")), ",").Length - 1
                    '        If Split(drv("IdentityID").ToString, ",")(i)="02" Then
                    '            BudID.ClearSelection()
                    '            Common.SetListItem(BudID, "02")
                    '        End If
                    '    Next
                    'End If
                    '20080604  Andy 修改供承辦人可正確比對保險證號與預算別 
                    '2010/05/24 預算別改成 不能修改
                    BudID.Enabled = False
                End If

                '加入判斷是否有資料未填 by nick
                Dim message As String = ""
                '*表示為該學員有必填資料未填
                Dim star1 As Label = e.Item.FindControl("star1") '* 必填項目確認
                star1.Visible = TIMS.CheckDataComplete(Me, drv, message) '--Chk_ClassStdApprAll /FN_STDNOMUSTDATA
                If message <> "" Then TIMS.Tooltip(star1, message, True)
                'end 加入判斷是否有資料未填

                Checkbox2.Value = Convert.ToString(drv("StudentID"))
                Checkbox2.Attributes("onclick") = "InsertValue(this.checked,this.value)"

                If PrintValue.Value.IndexOf(drv("StudentID")) <> -1 Then Checkbox2.Checked = True

                e.Item.Cells(cst_學號).Text = Right(String.Concat("00", drv("StudID")), 3)
                '**by Milor 20080530 start
                ''學號顯示的方式，改為去除學號前固定字串
                'If sm.UserInfo.TPlanID="28" Then
                '    e.Item.Cells(1).Text=Right(drv("StudentID").ToString, 2)
                'Else
                '    'Dim FWStudentID As String=DbAccess.ExecuteScalar("select a.Years+'0'+b.ClassID+a.CyclType from Class_ClassInfo a,ID_CLass b where a.CLSID=b.CLSID and a.OCID='" & ViewState("LastOCIDValue1") & "'")
                '    Dim FWStudentID As String=DbAccess.ExecuteScalar("select a.Years+'0'+b.ClassID+a.CyclType from Class_ClassInfo a,ID_CLass b where a.CLSID=b.CLSID and a.OCID='" & Me.OCIDValue1.Value & "'")
                '    e.Item.Cells(1).Text=Replace(e.Item.Cells(1).Text, FWStudentID, "")
                'End If
                'If Len(e.Item.Cells(1).Text)=12 Then
                '    e.Item.Cells(1).Text=Right(e.Item.Cells(1).Text, 3)
                'Else
                '    e.Item.Cells(1).Text=Right(e.Item.Cells(1).Text, 2)
                'End If
                e.Item.Cells(cst_性別).Text = ""
                Select Case Convert.ToString(drv("Sex"))
                    Case "M"
                        e.Item.Cells(cst_性別).Text = "男"
                    Case "F"
                        e.Item.Cells(cst_性別).Text = "女"
                End Select

                Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(drv("StudStatus"))
                e.Item.Cells(cst_學員狀態).Text = STUDSTATUS_N '"在訓"
                Select Case Convert.ToString(drv("StudStatus"))
                    Case "2"
                        e.Item.Cells(cst_學員狀態).Text = String.Concat(STUDSTATUS_N, "<br>(", drv("RejectTDate1"), ")")
                    Case "3"
                        e.Item.Cells(cst_學員狀態).Text = String.Concat(STUDSTATUS_N, "<br>(", drv("RejectTDate2"), ")")
                End Select

                Dim tmpEC3 As String = "" 'Convert.ToString(drv("EnterChannel"))
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If Convert.ToString(drv("ENTERPATH")) = "O" Then
                        tmpEC3 = "外網" '"網路"
                    ElseIf Convert.ToString(drv("ENTERPATH")) = "o" Then
                        tmpEC3 = "內網" '"現場"
                    End If
                    'TIMS.Tooltip(e.Item.Cells(cst_報名路徑), Convert.ToString(drv("ENTERPATH")), True)
                End If
                If tmpEC3 = "" Then
                    Select Case Convert.ToString(drv("EnterChannel"))
                        Case "1"
                            tmpEC3 = "外網"'"網路"
                        Case "2"
                            tmpEC3 = "內網"'"現場"
                        Case "3"
                            tmpEC3 = "通訊"
                        Case "4"
                            tmpEC3 = "推介"
                        Case Else
                            tmpEC3 = Convert.ToString(drv("EnterChannel"))
                    End Select
                End If
                e.Item.Cells(cst_報名路徑).Text = tmpEC3

                'btn.CommandArgument="SID=" & drv("SID") & "&SOCID=" & drv("SOCID") & ""
                Dim cmdArg As String = String.Concat("&SOCID=", drv("SOCID"), "&tmpName=", drv("Name"))
                btn1.CommandArgument = cmdArg

#Region "(No Use)"
                'btn1.Enabled=False '修改鈕
                'If blnCanMod Then btn1.Enabled=True

                'btn2.Visible=False '刪除動作
                'Select Case Convert.ToString(sm.UserInfo.RoleID)
                '    Case "1", "5"
                '        btn10Delete.Visible=True
                '        btn10Delete.Attributes("onclick")="return confirm('這樣會刪除此學員的相關班級資料,\n但不會刪除此學員的個人基本資料,\n確定要繼續刪除?');"
                'End Select

                'btn10Delete.CommandArgument=cmdArg
                'btn10Delete.Visible=False '刪除。
                'If Convert.ToString(sm.UserInfo.RoleID)="0" AndAlso Convert.ToString(sm.UserInfo.LID)="0" Then
                '    btn10Delete.Visible=True
                '    btn10Delete.Attributes("onclick")="return confirm('這樣會刪除此學員的相關班級資料,\n但不會刪除此學員的個人基本資料,\n確定要繼續刪除?');"
                'End If
#End Region

                If Not ViewState("_SearchStr") Is Nothing Then
                    If Convert.ToString(ViewState("_SearchStr")).IndexOf("Load=") > -1 Then
                        btn1.CommandName = "view"
                        btn1.Text = "檢視"
                        'btn2.Visible=False
                    End If
                End If

#Region "(No Use)"
                ''被遞補學員 為正式學員
                'If Convert.ToString(drv("MakeSOCID")) <> "" Then
                '    btn2.Enabled=False
                '    TIMS.Tooltip(btn2, " 已有被遞補學員：" & TIMS.GetSOCIDName(Convert.ToString(drv("MakeSOCID"))))
                'End If

                'If rblWorkMode.SelectedValue="1" Then
                '    e.Item.Cells(cst_身分證號碼).Text=TIMS.strMask(e.Item.Cells(cst_身分證號碼).Text, 1)
                '    e.Item.Cells(cst_出生日期).Text=TIMS.strMask(e.Item.Cells(cst_出生日期).Text, 2)
                'End If
#End Region

                If Not ViewState(cst_flgCIShow) Then
                    '不可顯示個資。
                    e.Item.Cells(cst_身分證號碼).Text = TIMS.strMask(e.Item.Cells(cst_身分證號碼).Text, 1)
                    e.Item.Cells(cst_出生日期).Text = TIMS.strMask(e.Item.Cells(cst_出生日期).Text, 2)
                End If
        End Select

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'Const cst_SOCID=8
        Dim Url_1 As String = ""
        Const cst_SD03002_addaspx As String = "SD_03_002_add.aspx" '28:產業人才投資方案  '有補助比例 (產投)
        Const cst_SD03002_add2aspx As String = "SD_03_002_add2.aspx" '06:在職進修訓練 '70:區域產業據點職業訓練計畫(在職)  '沒有補助比例(自辦在職／區域) 
        Dim str_SDADDASPX As String = cst_SD03002_addaspx
        If (Hid_nouse_SupplyID.Value = "Y") Then str_SDADDASPX = cst_SD03002_add2aspx

        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) ' .SelectedValue
        Select Case e.CommandName
            Case "edit", "view"
                Session(TIMS.gcst_rblWorkMode) = v_rblWorkMode 'rblWorkMode.SelectedValue 'Dim SOCID_value As String=""
                Dim SOCID_value As String = TIMS.GetMyValue(e.CommandArgument, "SOCID") 'Dim tmpName As String=TIMS.GetMyValue(e.CommandArgument, "tmpName")
                Session("SearchSOCID") = SOCID_value 'Session("SearchSOCID")=e.Item.Cells(cst_SOCID).Text
                If ViewState("_SearchStr") IsNot Nothing Then
                    If $"{ViewState("_SearchStr")}" <> "" AndAlso ViewState("_SearchStr").ToString.IndexOf("Load=") > -1 Then
                        Session("_SearchStr") = ViewState("_SearchStr")
                        ViewState("_SearchStr") = Nothing
                        Url_1 = str_SDADDASPX & "?ID=" & Request("ID") & "&todo=2&SD_03_002=VIEW&OCID=" & ViewState("LastOCIDValue1") & "&SOCID=" & SOCID_value
                        Call TIMS.Utl_Redirect(Me, objconn, Url_1)
                    Else
                        Call GetSearchStr()
                        Url_1 = str_SDADDASPX & "?ID=" & Request("ID") & "&OCID=" & ViewState("LastOCIDValue1") & "&SOCID=" & SOCID_value
                        Call TIMS.Utl_Redirect(Me, objconn, Url_1)
                    End If
                Else
                    Call GetSearchStr()
                    Url_1 = str_SDADDASPX & "?ID=" & Request("ID") & "&OCID=" & ViewState("LastOCIDValue1") & "&SOCID=" & SOCID_value
                    Call TIMS.Utl_Redirect(Me, objconn, Url_1)
                End If
            Case "del"
                '停用刪除動作
                Call sUtl_DeleteStud(e)
        End Select
    End Sub

    '停用刪除動作
    Sub sUtl_DeleteStud(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        Dim SOCID_VAL As String = TIMS.GetMyValue(e.CommandArgument, "SOCID")
        Dim tmpName As String = TIMS.GetMyValue(e.CommandArgument, "tmpName")

        SOCID_VAL = TIMS.ClearSQM(SOCID_VAL)
        If SOCID_VAL = "" Then
            Common.MessageBox(Me, "傳入參數有誤，停止刪除!!")
            Exit Sub
        End If

        Dim dr As DataRow = Nothing
        Dim sql As String = ""
        Dim MsgBox As String = ""

        '津貼
        sql = $" SELECT 'x' FROM Stud_SubsidyResult WHERE SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有津貼資料，不能刪除" & vbCrLf

        '職訓生活津貼
        sql = $" SELECT 'x' FROM Sub_SubSidyApply WHERE SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有職訓生活津貼資料，不能刪除" & vbCrLf

        '技能檢定
        sql = $" SELECT 'x' FROM Stud_TechExam WHERE SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有技能檢定資料，不能刪除" & vbCrLf

        '結訓成績 (分數大於0)
        sql = $" SELECT 'x' FROM Stud_TrainingResults WHERE Results >0 AND SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有結訓成績資料，不能刪除" & vbCrLf

        '操行
        sql = $" SELECT 'x' FROM Stud_Conduct WHERE SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有操行成績資料，不能刪除" & vbCrLf

        '轉班
        sql = $" SELECT 'x' FROM Stud_TranClassRecord WHERE SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有轉班資料，不能刪除" & vbCrLf

        '出缺勤
        sql = $" SELECT 'x' FROM Stud_Turnout WHERE SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有出缺勤資料，不能刪除" & vbCrLf

        '獎懲
        sql = $" SELECT * FROM Stud_Sanction WHERE SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有獎懲資料，不能刪除" & vbCrLf

        '結訓學員資料卡
        sql = $" SELECT 'x' FROM Stud_ResultStudData WHERE SOCID={SOCID_VAL}"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then MsgBox += "此學員現在有填寫結訓學員資料卡，不能刪除" & vbCrLf

        If MsgBox <> "" Then
            Common.MessageBox(Me, MsgBox)
            Exit Sub
        End If

        Dim sMemo As String = ""
        sMemo &= "&動作=刪除"
        sMemo &= "&NAME=" & tmpName
        'Session(TIMS.gcst_rblWorkMode)=rblWorkMode.SelectedValue
        '寫入Log查詢(SubInsAccountLog1(Auth_Accountlog))
        Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm刪除, TIMS.GetListValue(rblWorkMode), Me.OCIDValue1.Value, "")
        Page.RegisterStartupScript("del", "<script>wopen('SD_03_002_del.aspx?ID=" & Request("ID") & "&SOCID=" & SOCID_VAL & "','del',350,250,0)</script>")
    End Sub

    '查詢參訓歷史
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button9.Click
        Dim sql As String = ViewState("SD03002_SearchSqlStr")
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        'Dim IDNOArray As New ArrayList
        'For Each dr As DataRow In dt.Rows
        '    IDNOArray.Add(TIMS.ChangeIDNO(dr("IDNO").ToString))
        'Next
        'Session("IDNOArray")=IDNOArray
        'Page.RegisterStartupScript("History", "<script>window.open('../../SD/01/SD_01_001_old.aspx','history','width=700,height=500,scrollbars=1')</script>")

        Dim stmp1 As String = ""
        Dim IDNOArray As New ArrayList
        Dim iRow2 As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim checkbox2 As HtmlInputCheckBox = eItem.FindControl("checkbox2")
            Dim Hid_idno As HiddenField = eItem.FindControl("Hid_idno")
            If checkbox2.Checked Then
                stmp1 = TIMS.DecryptAes(Hid_idno.Value)
                If stmp1 <> "" AndAlso IDNOArray.IndexOf(stmp1) = -1 Then
                    iRow2 += 1
                    IDNOArray.Add(stmp1)
                End If
            End If
        Next
        If iRow2 = 0 Then
            For Each dr As DataRow In dt.Rows
                stmp1 = TIMS.ChangeIDNO(dr("IDNO").ToString)
                If stmp1 <> "" AndAlso IDNOArray.IndexOf(stmp1) = -1 Then IDNOArray.Add(stmp1)
            Next
        End If

        '排序方式
        Session("IDNOArray") = IDNOArray
        Dim rqID As String = TIMS.Get_MRqID(Me)
        Dim Script2 As String = $"<script>window.open('../05/SD_05_010_pop.aspx?ID={rqID}&SD_01_004_Type={CST_KD_STUDENTLIST}' ,'history','width=1400,height=820,scrollbars=1')</script>"
        Page.RegisterStartupScript("History2", Script2)
    End Sub

    'BudgetID 都有值, SupplyID 都有值 則為真，其餘為否 true:ok false:異常ng
    Function Check_CLASS_STUDENTSOFCLASSBS(ByVal OCIDValue As String) As Boolean
        Dim Rst As Boolean = True 'true:ok false:異常ng
        If OCIDValue1.Value = "" Then Return False 'true:ok false:異常ng
        'BudgetID 都有值, SupplyID 都有值 則為真，其餘為否
        Dim pms1 As New Hashtable From {{"OCID", CInt(OCIDValue)}}
        Dim sql As String = " SELECT COUNT(1) CNT FROM CLASS_STUDENTSOFCLASS WHERE OCID =@OCID AND (ISNULL(BudgetID,' ')=' ' OR ISNULL(SupplyID,' ')=' ')"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)
        If dt.Rows.Count = 0 Then Return False 'true:ok false:異常ng '(查無資料)
        If Convert.ToString(dt.Rows(0).Item("CNT")) = "" Then Return False 'true:ok false:異常ng '(有資料但為空)
        If Convert.ToString(dt.Rows(0).Item("CNT")) <> "0" Then Return False 'true:ok false:異常ng '(有資料不為0)
        'If Convert.ToString(dt.Rows(0).Item("CNT"))="0" Then Return Rst
        '(有資料且為0) TRUE:OK
        Return Rst
    End Function

    ''' <summary>
    ''' 學員資料確認
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If
        If ViewState("LastOCIDValue1") = "" Then
            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If
        If ViewState("LastOCIDValue1") <> OCIDValue1.Value Then
            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If
        'true:ok false:異常ng
        If Not Check_CLASS_STUDENTSOFCLASSBS(OCIDValue1.Value) Then
            Common.MessageBox(Me, String.Format("送出前學員預算別、補助比例尚未輸入，請確認(有星號的資料)"))
            Exit Sub
        End If

        Try
            '含離退
            Dim parms As New Hashtable From {{"OCID", OCIDValue1.Value}}
            Dim sql As String = " UPDATE CLASS_STUDENTSOFCLASS SET IsApprPaper='Y' WHERE OCID=@OCID "
            Dim iCnt As Integer = DbAccess.ExecuteNonQuery(sql, objconn, parms)

            'AppliedResultR
            'C 全班學員資料確認'Y 通過'R 全班學員資料確認，遭退件'NULL未確認過狀態
            sql = " UPDATE CLASS_CLASSINFO SET AppliedResultR='C' WHERE OCID=@OCID "
            iCnt = DbAccess.ExecuteNonQuery(sql, objconn, parms)
            If iCnt > 0 Then
                Button11.Enabled = False
                TIMS.Tooltip(Button11, "全班學員資料確認")
            End If

            'If Button11.Enabled=False Then Button1_Click(sender, e)
            If Not Button11.Enabled Then Call Search1() '查詢按鈕 SQL
        Catch ex As Exception
            Call TIMS.WriteTraceLog(ex.Message, ex)
            Dim slogMsg1 As String = ex.ToString
            TIMS.WriteLog(Me, slogMsg1)
            Common.MessageBox(Me, "!!儲存失敗!!")
            'Common.MessageBox(Me, ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' 學員資料審核
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub edit_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edit_but.Click
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If
        If ViewState("LastOCIDValue1") = "" Then
            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If
        If ViewState("LastOCIDValue1") <> OCIDValue1.Value Then
            Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
            Exit Sub
        End If
        Call GetSearchStr()
        Call TIMS.Utl_Redirect(Me, objconn, "SD_03_002_classver.aspx?OCID=" & OCIDValue1.Value & "&ID=" & Request("ID") & "" & "")
    End Sub

    '回上頁
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        'Session("_SearchStr")=ViewState("_SearchStr") 'Dim Review As String ''970505  Andy  回上頁 'Session("Review")="yes"
        Session("_SearchStr") = ViewState("_SearchStr")
        Dim url As String = TIMS.GetFunIDUrl(Request("ID"), 0, objconn)
        Call TIMS.Utl_Redirect(Me, objconn, url & "?ID=" & Request("ID") & "")
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        'ImportTable.Style.Item("display")="none"
        trImport1.Style.Item("display") = "none"
        DataGridTable.Style.Item("display") = "none"
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        'ImportTable.Style.Item("display")="none"
        trImport1.Style.Item("display") = "none"
        DataGridTable.Style.Item("display") = "none"

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    ''' <summary>'查詢鈕  '匯出鈕 'hidSchBtnNum.value: 1.正常查詢 2.正常匯出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Sub SUtl_btnSearchData1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim BtnObj As Button = CType(sender, Button)
        Const cst_button1 As String = "button1" '查詢
        'Const cst_button6 As String="button6" '匯出
        Const cst_btndivPwdSubmit As String = "btndivpwdsubmit" ' hidSchBtnNum.value: 1.正常查詢 2.正常匯出
        Dim sMsg As String = ""

        Select Case LCase(BtnObj.CommandName)
            Case cst_button1 '查詢鈕
                Call Search1()
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If msg.Text = cst_msgNoStdData Then Common.MessageBox(Me, cst_msgTPlanID28NoStdData)
                End If
            'Case cst_button6 '匯出鈕 Button6_Click ' Call Export1()

            Case cst_btndivPwdSubmit
                '正常顯示 '查詢或匯出。
                If Not TIMS.sUtl_ChkPlanPwd(sm.UserInfo.PlanID, objconn) Then
                    sMsg = "未設定計畫密碼!!"
                    labChkMsg.Text = sMsg
                    Common.MessageBox(Me, sMsg)
                    Exit Sub
                End If
                If Not TIMS.sUtl_ChkPlanPwdOK(objconn, sm.UserInfo.PlanID, txtdivPxssward.Text) Then
                    sMsg = "個資安全密碼錯誤!!"
                    labChkMsg.Text = sMsg
                    Common.MessageBox(Me, sMsg)
                    Exit Sub
                End If
                'If rblWorkMode.SelectedValue="2" Then flgCIShow=True '可正常顯示個資。
                txtdivPxssward.Text = ""
                Select Case hidSchBtnNum.Value
                    Case "1"
                        Call Search1()
                        'Case "2" '匯出鈕 Button6_Click  Call Export1()
                End Select
        End Select
    End Sub

    ''' <summary>
    ''' 匯出名冊
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim v_SHIFTSORT As String = TIMS.GetListValue(shiftsort)
        Dim htPP As New Hashtable From {{"v_SHIFTSORT", v_SHIFTSORT}}

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用(產投) '~\SD\03\Sample2.xls  / '~\SD\03\Sample28.xls
            '學號	中文姓名	LastName	FirstName '身分證字號	性別	
            '最高學歷    畢業狀況 '聯絡電話_日	聯絡電話_夜	行動電話	主要參訓身分別	開訓日期	結訓日期	報到日期
            ExportXlsStd28(htPP)
        Else
            'TIMS '~\SD\03\Sample.xls  / '~\SD\03\Sample06.xls
            '學號	中文姓名	LastName	FirstName '身分證字號	性別 
            '身分別	非本國人身分別 '原屬國籍	護照或工作證號	
            '最高學歷 '畢業狀況	'聯絡電話_日	聯絡電話_夜	手機	主要參訓身分別	開訓日期	結訓日期	報到日期
            ExportXlsStd06(htPP)
        End If

    End Sub

    ''' <summary>
    ''' 匯出名冊 產投
    ''' </summary>
    ''' <param name="hPP"></param>
    Sub ExportXlsStd28(hPP As Hashtable)
        Const Cst_FileSavePath As String = "~/SD/03/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim Sql_iden As String = " SELECT * FROM Key_Identity ORDER BY IDENTITYID"
        dtIdentity = DbAccess.GetDataTable(Sql_iden, objconn)

        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) ' .SelectedValue
        '1:模糊顯示 2:可正常顯示個資。
        flgCIShow = If(v_rblWorkMode = TIMS.cst_wmdip2, True, False)
        ViewState(cst_flgCIShow) = flgCIShow

        drOCID = Nothing
        If OCIDValue1.Value <> "" Then
            drOCID = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
            If drOCID Is Nothing Then
                Common.MessageBox(Me, "班級查詢有誤。")
                Exit Sub
            End If
            'If Convert.ToString(drOCID("ShowOK14"))="Y" Then flgCIShow=True '可正常顯示個資。
        End If

        Dim dtXls1 As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        Dim v_SHIFTSORT As String = TIMS.GetMyValue2(hPP, "v_SHIFTSORT")

        Dim END_COL_NM As String = "O"
        'Dim cellsCOLSPNumF As String=String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim cellsCOLSPNumF2 As String = String.Concat("A1:", END_COL_NM, "{0}") '(畫格子使用)
        Dim strErrmsg As String = ""

        '113年度下半年提升勞工自主學習計畫核定課程明細表(北基宜花金馬分署)									
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim V_SHEETNM1 As String = TIMS.GetDateNo() ' String.Concat(s_ROCYEAR1, s_APPSTAGE_NM2, V_SHTNM1, "-", V_DISTNAME3)

        Dim ClsNM1 As String = TIMS.ChangeIDNO(Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", ""))
        Dim s_FILENAME1 As String = TIMS.GetValidFileName(TIMS.ClearSQM(String.Concat(ClsNM1, V_SHEETNM1)))

        'SyncLock print_lock
        'ExcelPackage.LicenseContext=LicenseContext.Commercial 'ExcelPackage.LicenseContext=LicenseContext.NonCommercial

        'Dim file1 As New FileInfo(filePath1)
        Dim ndt As DateTime = Now
        Dim ep As New ExcelPackage()

        Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
        'Dim ws As ExcelWorksheet=ep.Workbook.Worksheets(0)

        Dim idxStr As Integer = 1
        TIMS.SetCellValue(ws, "A" & idxStr, "學號")
        TIMS.SetCellValue(ws, "B" & idxStr, "中文姓名")
        TIMS.SetCellValue(ws, "C" & idxStr, "LastName")
        TIMS.SetCellValue(ws, "D" & idxStr, "FirstName")
        TIMS.SetCellValue(ws, "E" & idxStr, "身分證字號")
        TIMS.SetCellValue(ws, "F" & idxStr, "性別")
        TIMS.SetCellValue(ws, "G" & idxStr, "最高學歷")
        TIMS.SetCellValue(ws, "H" & idxStr, "畢業狀況")
        TIMS.SetCellValue(ws, "I" & idxStr, "聯絡電話_日")
        TIMS.SetCellValue(ws, "J" & idxStr, "聯絡電話_夜")
        TIMS.SetCellValue(ws, "K" & idxStr, "行動電話")
        TIMS.SetCellValue(ws, "L" & idxStr, "主要參訓身分別")
        TIMS.SetCellValue(ws, "M" & idxStr, "開訓日期")
        TIMS.SetCellValue(ws, "N" & idxStr, "結訓日期")
        TIMS.SetCellValue(ws, "O" & idxStr, "報到日期")
        ws.Cells("A1:O1").Style.Font.Bold = True

        idxStr = 2
        For Each dr As DataRow In dtXls1.Rows
            Dim STUDID As String = ""
            Dim CNAME As String = ""
            Dim LastName As String = ""
            Dim FirstName As String = ""
            Dim IDNO As String = ""
            Dim Sex As String = ""
            Dim DegreeID As String = ""
            Dim GradID As String = ""
            Dim PhoneD As String = ""
            Dim PhoneN As String = ""
            Dim CellPhone As String = ""
            Dim MIdentityID As String = ""
            Dim OpenDate As String = ""
            Dim CloseDate As String = ""
            Dim EnterDate As String = ""

            STUDID = dr("STUDID").ToString().Replace("'", "")
            CNAME = dr("Name").ToString().Replace("'", "")
            LastName = Convert.ToString(dr("EngName")).Replace("'", "")
            FirstName = ""
            If dr("EngName").ToString <> "" Then
                If dr("EngName").ToString.IndexOf(" ") <> -1 Then
                    LastName = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
                    FirstName = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - dr("EngName").ToString.IndexOf(" ") - 1)).Replace("'", "")
                End If
            End If
            IDNO = TIMS.ChangeIDNO(dr("IDNO").ToString())
            If Not flgCIShow Then IDNO = TIMS.strMask(IDNO, 1) '不顯示個資。
            'If rblWorkMode.SelectedValue="1" Then IDNO=TIMS.strMask(IDNO, 1)
            Sex = Convert.ToString(If(v_SHIFTSORT = "1", dr("SEX"), dr("SEXName")))
            '最高學歷
            DegreeID = Convert.ToString(If(v_SHIFTSORT = "1", dr("DegreeID"), dr("DegreeName")))
            '畢業狀況
            GradID = Convert.ToString(If(v_SHIFTSORT = "1", dr("GradID"), dr("GradName"))).Replace("'", "")
            'flgCIShow : 可顯示個資 / 不顯示個資。
            PhoneD = If(flgCIShow, dr("PhoneD").ToString().Replace("'", ""), "****")
            PhoneN = If(flgCIShow, dr("PhoneN").ToString().Replace("'", ""), "****")
            CellPhone = If(flgCIShow, dr("CellPhone").ToString().Replace("'", ""), "****")
            OpenDate = TIMS.Cdate3(dr("OpenDate"))
            CloseDate = TIMS.Cdate3(dr("CloseDate"))
            EnterDate = TIMS.Cdate3(dr("EnterDate"))
            MIdentityID = dr("MIdentityID").ToString
            If Not (v_SHIFTSORT = "1") Then MIdentityID = TIMS.Get_IdentityName(dr("IdentityID").ToString, dtIdentity, "，")

            TIMS.SetCellValue(ws, "A" & idxStr, STUDID)
            TIMS.SetCellValue(ws, "B" & idxStr, CNAME)
            TIMS.SetCellValue(ws, "C" & idxStr, LastName)
            TIMS.SetCellValue(ws, "D" & idxStr, FirstName)
            TIMS.SetCellValue(ws, "E" & idxStr, IDNO)
            TIMS.SetCellValue(ws, "F" & idxStr, Sex)
            TIMS.SetCellValue(ws, "G" & idxStr, DegreeID)
            TIMS.SetCellValue(ws, "H" & idxStr, GradID)
            TIMS.SetCellValue(ws, "I" & idxStr, PhoneD)
            TIMS.SetCellValue(ws, "J" & idxStr, PhoneN)
            TIMS.SetCellValue(ws, "K" & idxStr, CellPhone)
            TIMS.SetCellValue(ws, "L" & idxStr, MIdentityID)
            TIMS.SetCellValue(ws, "M" & idxStr, OpenDate)
            TIMS.SetCellValue(ws, "N" & idxStr, CloseDate)
            TIMS.SetCellValue(ws, "O" & idxStr, EnterDate)
            idxStr += 1
        Next

        idxStr -= 1 '(畫線)
        Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
            With exlRow3X
                '.Style.Font.Name=fontName
                .Style.Font.Size = fontSize12s 'FontSize
                .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                '.AutoFitColumns(10.0, 250.0)
            End With
            TIMS.SetCellBorder(exlRow3X)
        End Using

        ' 設定貨幣格式，小數位數為 0
        'ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format="$#,##0" ' 美元符號，您可以根據需要更改
        'ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width=33

        ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom=70 無法運行 修正為 ws.View.ZoomScale=70 才可運行
        ws.View.ZoomScale = 90

        Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
        Select Case V_ExpType
            Case "EXCEL"
                TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                TIMS.Utl_RespWriteEnd(Me, objconn, "")
            Case "ODS"
                TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                TIMS.Utl_RespWriteEnd(Me, objconn, "")
            Case Else
                Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                Common.MessageBox(Me, s_log1)
                Return ' Exit Sub
        End Select
        'End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If

    End Sub

    ''' <summary>
    ''' 匯出名冊 非產投
    ''' </summary>
    ''' <param name="hPP"></param>
    Sub ExportXlsStd06(hPP As Hashtable)
        Const Cst_FileSavePath As String = "~/SD/03/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim Sql_iden As String = " SELECT * FROM Key_Identity ORDER BY IDENTITYID"
        dtIdentity = DbAccess.GetDataTable(Sql_iden, objconn)

        Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) ' .SelectedValue
        '1:模糊顯示 2:可正常顯示個資。
        flgCIShow = If(v_rblWorkMode = TIMS.cst_wmdip2, True, False)
        ViewState(cst_flgCIShow) = flgCIShow

        drOCID = Nothing
        If OCIDValue1.Value <> "" Then
            drOCID = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
            If drOCID Is Nothing Then
                Common.MessageBox(Me, "班級查詢有誤。")
                Exit Sub
            End If
            'If Convert.ToString(drOCID("ShowOK14"))="Y" Then flgCIShow=True '可正常顯示個資。
        End If

        Dim dtXls1 As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        Dim v_SHIFTSORT As String = TIMS.GetMyValue2(hPP, "v_SHIFTSORT")

        Dim END_COL_NM As String = "S"
        'Dim cellsCOLSPNumF As String=String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim cellsCOLSPNumF2 As String = String.Concat("A1:", END_COL_NM, "{0}") '(畫格子使用)
        Dim strErrmsg As String = ""

        '113年度下半年提升勞工自主學習計畫核定課程明細表(北基宜花金馬分署)									
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim V_SHEETNM1 As String = TIMS.GetDateNo() ' String.Concat(s_ROCYEAR1, s_APPSTAGE_NM2, V_SHTNM1, "-", V_DISTNAME3)

        Dim ClsNM1 As String = TIMS.ChangeIDNO(Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", ""))
        Dim s_FILENAME1 As String = TIMS.GetValidFileName(TIMS.ClearSQM(String.Concat(ClsNM1, V_SHEETNM1)))

        'SyncLock print_lock
        'ExcelPackage.LicenseContext=LicenseContext.Commercial'ExcelPackage.LicenseContext=LicenseContext.NonCommercial

        'Dim file1 As New FileInfo(filePath1)
        Dim ndt As DateTime = Now
        Dim ep As New ExcelPackage()

        Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
        'Dim ws As ExcelWorksheet=ep.Workbook.Worksheets(0)

        Dim idxStr As Integer = 1
        TIMS.SetCellValue(ws, "A" & idxStr, "學號")
        TIMS.SetCellValue(ws, "B" & idxStr, "中文姓名")
        TIMS.SetCellValue(ws, "C" & idxStr, "LastName")
        TIMS.SetCellValue(ws, "D" & idxStr, "FirstName")
        TIMS.SetCellValue(ws, "E" & idxStr, "身分證字號")
        TIMS.SetCellValue(ws, "F" & idxStr, "性別")
        TIMS.SetCellValue(ws, "G" & idxStr, "身分別")
        TIMS.SetCellValue(ws, "H" & idxStr, "非本國人身分別")
        TIMS.SetCellValue(ws, "I" & idxStr, "原屬國籍")
        TIMS.SetCellValue(ws, "J" & idxStr, "護照或工作證號")

        TIMS.SetCellValue(ws, "K" & idxStr, "最高學歷")
        TIMS.SetCellValue(ws, "L" & idxStr, "畢業狀況")
        TIMS.SetCellValue(ws, "M" & idxStr, "聯絡電話_日")
        TIMS.SetCellValue(ws, "N" & idxStr, "聯絡電話_夜")
        TIMS.SetCellValue(ws, "O" & idxStr, "手機")
        TIMS.SetCellValue(ws, "P" & idxStr, "主要參訓身分別")
        TIMS.SetCellValue(ws, "Q" & idxStr, "開訓日期")
        TIMS.SetCellValue(ws, "R" & idxStr, "結訓日期")
        TIMS.SetCellValue(ws, "S" & idxStr, "報到日期")
        ws.Cells("A1:S1").Style.Font.Bold = True

        idxStr = 2
        For Each dr As DataRow In dtXls1.Rows
            Dim STUDID As String = ""
            Dim CNAME As String = ""
            Dim LastName As String = ""
            Dim FirstName As String = ""
            Dim IDNO As String = ""
            Dim Sex As String = ""
            Dim PassPortNO As String = ""
            Dim ChinaOrNot As String = ""
            Dim Nationality As String = ""
            Dim PPNO As String = ""
            Dim DegreeID As String = ""
            Dim GradID As String = ""
            Dim PhoneD As String = ""
            Dim PhoneN As String = ""
            Dim CellPhone As String = ""
            Dim MIdentityID As String = ""
            Dim OpenDate As String = ""
            Dim CloseDate As String = ""
            Dim EnterDate As String = ""

            STUDID = dr("STUDID").ToString().Replace("'", "")
            CNAME = dr("Name").ToString().Replace("'", "")
            LastName = Convert.ToString(dr("EngName")).Replace("'", "")
            FirstName = ""
            If dr("EngName").ToString <> "" Then
                If dr("EngName").ToString.IndexOf(" ") <> -1 Then
                    LastName = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
                    FirstName = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - dr("EngName").ToString.IndexOf(" ") - 1)).Replace("'", "")
                End If
            End If
            IDNO = TIMS.ChangeIDNO(dr("IDNO").ToString())
            If Not flgCIShow Then IDNO = TIMS.strMask(IDNO, 1) '不顯示個資。
            'If rblWorkMode.SelectedValue="1" Then IDNO=TIMS.strMask(IDNO, 1)
            Sex = Convert.ToString(If(v_SHIFTSORT = "1", dr("SEX"), dr("SEXName")))

            PassPortNO = Convert.ToString(If(v_SHIFTSORT = "1", dr("PassPortNO"), dr("PassPortName"))).Replace("'", "")
            ChinaOrNot = Convert.ToString(If(v_SHIFTSORT = "1", dr("ChinaOrNot"), dr("ChinaOrNotName"))).Replace("'", "")
            Nationality = Convert.ToString(dr("Nationality")).Replace("'", "")
            PPNO = Convert.ToString(If(v_SHIFTSORT = "1", dr("PPNO"), dr("PPNOName"))).Replace("'", "")
            '最高學歷
            DegreeID = Convert.ToString(If(v_SHIFTSORT = "1", dr("DegreeID"), dr("DegreeName")))
            '畢業狀況
            GradID = Convert.ToString(If(v_SHIFTSORT = "1", dr("GradID"), dr("GradName"))).Replace("'", "")
            'flgCIShow : 可顯示個資 / 不顯示個資。
            PhoneD = If(flgCIShow, dr("PhoneD").ToString().Replace("'", ""), "****")
            PhoneN = If(flgCIShow, dr("PhoneN").ToString().Replace("'", ""), "****")
            CellPhone = If(flgCIShow, dr("CellPhone").ToString().Replace("'", ""), "****")
            OpenDate = TIMS.Cdate3(dr("OpenDate"))
            CloseDate = TIMS.Cdate3(dr("CloseDate"))
            EnterDate = TIMS.Cdate3(dr("EnterDate"))
            MIdentityID = dr("MIdentityID").ToString
            If Not (v_SHIFTSORT = "1") Then MIdentityID = TIMS.Get_IdentityName(dr("IdentityID").ToString, dtIdentity, "，")

            TIMS.SetCellValue(ws, "A" & idxStr, STUDID)
            TIMS.SetCellValue(ws, "B" & idxStr, CNAME)
            TIMS.SetCellValue(ws, "C" & idxStr, LastName)
            TIMS.SetCellValue(ws, "D" & idxStr, FirstName)
            TIMS.SetCellValue(ws, "E" & idxStr, IDNO)
            TIMS.SetCellValue(ws, "F" & idxStr, Sex)
            TIMS.SetCellValue(ws, "G" & idxStr, PassPortNO) '"身分別")
            TIMS.SetCellValue(ws, "H" & idxStr, ChinaOrNot) '"非本國人身分別")
            TIMS.SetCellValue(ws, "I" & idxStr, Nationality) '"原屬國籍")
            TIMS.SetCellValue(ws, "J" & idxStr, PPNO) '"護照或工作證號")

            TIMS.SetCellValue(ws, "K" & idxStr, DegreeID)
            TIMS.SetCellValue(ws, "L" & idxStr, GradID)
            TIMS.SetCellValue(ws, "M" & idxStr, PhoneD)
            TIMS.SetCellValue(ws, "N" & idxStr, PhoneN)
            TIMS.SetCellValue(ws, "O" & idxStr, CellPhone)
            TIMS.SetCellValue(ws, "P" & idxStr, MIdentityID)
            TIMS.SetCellValue(ws, "Q" & idxStr, OpenDate)
            TIMS.SetCellValue(ws, "R" & idxStr, CloseDate)
            TIMS.SetCellValue(ws, "S" & idxStr, EnterDate)
            idxStr += 1
        Next

        idxStr -= 1 '(畫線)
        Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
            With exlRow3X
                '.Style.Font.Name=fontName
                .Style.Font.Size = fontSize12s 'FontSize
                .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                '.AutoFitColumns(10.0, 250.0)
            End With
            TIMS.SetCellBorder(exlRow3X)
        End Using

        ' 設定貨幣格式，小數位數為 0
        'ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format="$#,##0" ' 美元符號，您可以根據需要更改
        'ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width=33

        ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom=70 無法運行 修正為 ws.View.ZoomScale=70 才可運行
        ws.View.ZoomScale = 90

        Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
        Select Case V_ExpType
            Case "EXCEL"
                TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                TIMS.Utl_RespWriteEnd(Me, objconn, "")
            Case "ODS"
                TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                TIMS.Utl_RespWriteEnd(Me, objconn, "")
            Case Else
                Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                Common.MessageBox(Me, s_log1)
                Return ' Exit Sub
        End Select
        'End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If

    End Sub

#Region "NO USE"
    'Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click  OUTPUT 匯出名冊 (匯出學員資料)
    'Public Shared Sub ExcelOutColumn(ByRef MyPage As Page, ByRef sm As SessionModel, ByRef oConn As SqlConnection,
    '                                 ByRef MyConn As OleDb.OleDbConnection, ByRef dtIdentity As DataTable,
    '                                 ByRef flgCIShow As Boolean, ByRef hPP As Hashtable, ByRef dr As DataRow, ByRef t_sql As String)
    '    Dim v_SHIFTSORT As String = TIMS.GetMyValue2(hPP, "v_SHIFTSORT")

    '    Dim StudentID As String = ""
    '    Dim Name As String = ""
    '    Dim LastName As String = ""
    '    Dim FirstName As String = ""
    '    Dim IDNO As String = ""
    '    Dim IDNO_MK As String = ""
    '    Dim SEX As String = ""
    '    Dim PassPortNO As String = ""
    '    Dim ChinaOrNot As String = ""
    '    Dim Nationality As String = ""
    '    Dim PPNO As String = ""
    '    Dim Birthday As String = ""
    '    Dim MaritalStatus As String = ""
    '    Dim DegreeID As String = ""
    '    Dim School As String = ""
    '    Dim Department As String = ""
    '    Dim GradID As String = ""
    '    Dim MilitaryID As String = ""
    '    Dim ServiceID As String = ""
    '    Dim MilitaryAppointment As String = ""
    '    Dim MilitaryRank As String = ""
    '    Dim ServiceOrg As String = ""
    '    Dim ChiefRankName As String = ""
    '    Dim ServicePhone As String = ""
    '    Dim SServiceDate As String = ""
    '    Dim FServiceDate As String = ""
    '    Dim ZipCode4 As String = ""
    '    Dim Address4 As String = ""
    '    Dim JobState As String = ""
    '    Dim PhoneD As String = ""
    '    Dim PhoneN As String = ""
    '    Dim CellPhone As String = ""
    '    Dim ZipCode1 As String = ""
    '    Dim Address1 As String = ""
    '    Dim ZipCode2 As String = ""
    '    Dim Address2 As String = ""
    '    Dim Email As String = ""
    '    Dim IdentityID As String = ""
    '    Dim MIdentityID As String = ""
    '    Dim SubsidyID As String = ""
    '    Dim OpenDate As String = ""
    '    Dim CloseDate As String = ""
    '    Dim EnterDate As String = ""
    '    Dim HandTypeID As String = ""
    '    Dim HandLevelID As String = ""
    '    Dim EmergencyContact As String = ""
    '    Dim EmergencyRelation As String = ""
    '    Dim EmergencyPhone As String = ""
    '    Dim ZipCode3 As String = ""
    '    Dim Address3 As String = ""
    '    Dim PriorWorkOrg1 As String = ""
    '    Dim Title1 As String = ""
    '    Dim SOfficeYM1 As String = ""
    '    Dim FOfficeYM1 As String = ""
    '    Dim PriorWorkOrg2 As String = ""
    '    Dim Title2 As String = ""
    '    Dim SOfficeYM2 As String = ""
    '    Dim FOfficeYM2 As String = ""
    '    Dim PriorWorkPay As String = ""
    '    Dim Traffic As String = ""
    '    Dim RealJobless As String = ""
    '    Dim JoblessID As String = ""
    '    Dim ShowDetail As String = ""
    '    Dim LevelNo As String = ""
    '    Dim EnterChannel As String = ""
    '    Dim TRNDMode As String = ""
    '    Dim TRNDType As String = ""
    '    Dim BudgetID As String = ""
    '    Dim SupplyID As String = ""
    '    Dim IsAgree As String = ""
    '    Dim PMode As String = ""
    '    Dim ForeName As String = ""
    '    Dim ForeTitle As String = ""
    '    Dim ForeSex As String = ""
    '    Dim ForeBirth As String = ""
    '    Dim ForeIDNO As String = ""
    '    Dim ForeZip As String = ""
    '    Dim ForeAddr As String = ""
    '    Dim KNID As String = ""

    '    Dim AcctMode As String = ""
    '    Dim PostNo As String = ""
    '    Dim AcctHeadNo As String = ""
    '    Dim AcctExNo As String = ""
    '    Dim AcctNo As String = ""
    '    Dim BankName As String = ""
    '    Dim ExBankName As String = ""
    '    Dim FirDate As String = ""
    '    Dim Uname As String = ""
    '    Dim Intaxno As String = ""
    '    Dim Tel As String = ""
    '    Dim Fax As String = ""
    '    Dim Zip As String = ""
    '    Dim Addr As String = ""
    '    Dim ServDept As String = ""
    '    Dim JobTitle As String = ""
    '    Dim SDate As String = ""
    '    Dim SJDate As String = ""
    '    Dim SPDate As String = ""
    '    Dim Q1 As String = ""
    '    Dim Q2 As String = ""
    '    Dim Q3 As String = ""
    '    Dim Q3_Other As String = ""
    '    Dim Q4 As String = ""
    '    Dim Q5 As String = ""
    '    Dim Q61 As String = ""
    '    Dim Q62 As String = ""
    '    Dim Q63 As String = ""
    '    Dim Q64 As String = ""
    '    '出生日期、學校名稱、科系、通訊地址、電子郵件帳號、參訓身分別、障礙類別、障礙等級、緊急通知人姓名、緊急通知人關係、緊急通知人電話、
    '    '緊急通知人地址郵遞區號、緊急通知人地址、報名管道、個資法意願、撥款方式、郵政_局號、金融_總代號、金融_分支代號、帳號、銀行名稱、分行名稱、
    '    '第一次投保勞保日、公司名稱、統編、公司電話、公司傳真、公司地址郵遞區號、公司地址、目前任職部門、職稱、個人到任目前任職公司起日、
    '    '個人到任目前職務起日、最近升遷日期、是否由公司推薦參訓、參訓動機、訓後動向、訓後動向其他說明、服務單位行業別、服務單位是否屬於中小企業、
    '    '個人工作年資、在這家公司的年資、在這職位的年資、最近升遷離本職幾年、是否提供基本資料查詢、戶籍地址郵遞區號、戶籍地址、預算別(經費來源)、補助比例。

    '    'StudentID=Right(dr("StudentID").ToString, 2).Replace("'", "")
    '    StudentID = dr("STUDID").ToString().Replace("'", "")
    '    Name = dr("Name").ToString().Replace("'", "")
    '    LastName = Convert.ToString(dr("EngName")).Replace("'", "")
    '    FirstName = ""
    '    If dr("EngName").ToString <> "" Then
    '        If dr("EngName").ToString.IndexOf(" ") <> -1 Then
    '            LastName = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
    '            FirstName = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - dr("EngName").ToString.IndexOf(" ") - 1)).Replace("'", "")
    '        End If
    '    End If
    '    IDNO = TIMS.ChangeIDNO(dr("IDNO").ToString())
    '    If Not flgCIShow Then IDNO = TIMS.strMask(IDNO, 1) '不顯示個資。

    '    'If rblWorkMode.SelectedValue="1" Then IDNO=TIMS.strMask(IDNO, 1)
    '    SEX = Convert.ToString(If(v_SHIFTSORT = "1", dr("SEX"), dr("SEXName")))
    '    'PassPortNO=dr("PassPortNO").ToString

    '    PassPortNO = Convert.ToString(If(v_SHIFTSORT = "1", dr("PassPortNO"), dr("PassPortName"))).Replace("'", "")
    '    ChinaOrNot = Convert.ToString(If(v_SHIFTSORT = "1", dr("ChinaOrNot"), dr("ChinaOrNotName"))).Replace("'", "")
    '    Nationality = Convert.ToString(dr("Nationality")).Replace("'", "")
    '    PPNO = Convert.ToString(If(v_SHIFTSORT = "1", dr("PPNO"), dr("PPNOName"))).Replace("'", "")

    '    Birthday = ""
    '    If Convert.ToString(dr("Birthday")) <> "" Then
    '        Birthday = TIMS.Cdate3(dr("Birthday")) 'DateFormat.ShortDate
    '        If Not flgCIShow Then Birthday = TIMS.strMask(Birthday, 2) '不顯示個資。
    '    End If
    '    MaritalStatus = Convert.ToString(If(v_SHIFTSORT = "1", dr("MaritalStatus"), dr("MaritalStatusName")))
    '    '最高學歷
    '    DegreeID = Convert.ToString(If(v_SHIFTSORT = "1", dr("DegreeID"), dr("DegreeName")))
    '    School = dr("School").ToString()

    '    '11~20
    '    Department = dr("Department").ToString()
    '    '畢業狀況
    '    GradID = Convert.ToString(If(v_SHIFTSORT = "1", dr("GradID"), dr("GradName"))).Replace("'", "")
    '    MilitaryID = Convert.ToString(If(v_SHIFTSORT = "1", dr("MilitaryID"), dr("MilitaryName"))).Replace("'", "")

    '    ServiceID = dr("ServiceID").ToString
    '    MilitaryAppointment = dr("MilitaryAppointment").ToString
    '    MilitaryRank = dr("MilitaryRank").ToString
    '    ServiceOrg = dr("ServiceOrg").ToString
    '    ChiefRankName = dr("ChiefRankName").ToString
    '    ServicePhone = dr("ServicePhone").ToString
    '    SServiceDate = dr("SServiceDate").ToString

    '    '21~30
    '    FServiceDate = dr("FServiceDate").ToString
    '    ZipCode4 = Convert.ToString(dr("ZipCode4"))
    '    If dr("ZipCode4").ToString <> "" Then
    '        'ZipCode4=dr("ZipCode4").ToString
    '        If v_SHIFTSORT = "2" Then
    '            If dr("CTName4").ToString <> dr("ZipName4").ToString Then
    '                ZipCode4 = "(" & dr("ZipCode4").ToString & ")" & dr("CTName4").ToString & dr("ZipName4").ToString
    '            Else
    '                ZipCode4 = "(" & dr("ZipCode4").ToString & ")" & dr("CTName4").ToString
    '            End If
    '        End If
    '    End If
    '    Address4 = dr("Address4").ToString

    '    JobState = Convert.ToString(If(v_SHIFTSORT = "1", dr("JobState"), dr("JobStateName")))

    '    'flgCIShow : 可顯示個資 / 不顯示個資。
    '    PhoneD = If(flgCIShow, dr("PhoneD").ToString().Replace("'", ""), "****")
    '    PhoneN = If(flgCIShow, dr("PhoneN").ToString().Replace("'", ""), "****")
    '    CellPhone = If(flgCIShow, dr("CellPhone").ToString().Replace("'", ""), "****")

    '    ZipCode1 = Convert.ToString(dr("ZipCode1"))
    '    If dr("ZipCode1").ToString <> "" Then
    '        'ZipCode1=dr("ZipCode1").ToString()
    '        If v_SHIFTSORT = "2" Then
    '            If dr("CTName1").ToString <> dr("ZipName1").ToString Then
    '                ZipCode1 = "(" & dr("ZipCode1").ToString & ")" & dr("CTName1").ToString & dr("ZipName1").ToString
    '            Else
    '                ZipCode1 = "(" & dr("ZipCode1").ToString & ")" & dr("CTName1").ToString
    '            End If
    '        End If
    '    End If

    '    Address1 = dr("Address1").ToString()
    '    'If rblWorkMode.SelectedValue="1" Then Address1=TIMS.strMask(Address1, 3)
    '    ZipCode2 = Convert.ToString(dr("ZipCode2"))
    '    If dr("ZipCode2").ToString <> "" Then 'ZipCode2=dr("ZipCode2").ToString()
    '        If v_SHIFTSORT = "2" Then
    '            If dr("CTName2").ToString <> dr("ZipName2").ToString Then
    '                ZipCode2 = "(" & dr("ZipCode2").ToString & ")" & dr("CTName2").ToString & dr("ZipName2").ToString
    '            Else
    '                ZipCode2 = "(" & dr("ZipCode2").ToString & ")" & dr("CTName2").ToString
    '            End If
    '        End If
    '    End If
    '    Address2 = dr("Address2").ToString()
    '    'If rblWorkMode.SelectedValue="1" Then Address2=TIMS.strMask(Address2, 3)

    '    '31~41
    '    Email = dr("Email").ToString()
    '    IdentityID = dr("IdentityID").ToString
    '    MIdentityID = dr("MIdentityID").ToString
    '    If v_SHIFTSORT = "1" Then
    '        IdentityID = Replace(dr("IdentityID").ToString, ",", "，")
    '    Else
    '        IdentityID = TIMS.Get_IdentityName(dr("IdentityID").ToString, dtIdentity, "，")
    '        MIdentityID = TIMS.Get_IdentityName(dr("IdentityID").ToString, dtIdentity, "，")
    '    End If
    '    'by Vicient
    '    KNID = Convert.ToString(If(v_SHIFTSORT = "1", dr("KNID"), dr("knName")))

    '    SubsidyID = Convert.ToString(If(v_SHIFTSORT = "1", dr("SubsidyID"), dr("SubsidyName")))

    '    OpenDate = TIMS.Cdate3(dr("OpenDate"))
    '    CloseDate = TIMS.Cdate3(dr("CloseDate"))
    '    EnterDate = TIMS.Cdate3(dr("EnterDate"))

    '    HandTypeID = TIMS.ClearSQM(If(v_SHIFTSORT = "1", dr("HandTypeID"), dr("HandTypeName")))
    '    HandLevelID = TIMS.ClearSQM(If(v_SHIFTSORT = "1", dr("HandLevelID"), dr("HandLevelName")))
    '    'If HandTypeID <> "" Then HandTypeID=Trim(HandTypeID)
    '    'If HandLevelID <> "" Then HandLevelID=Trim(HandLevelID)
    '    '若是都空白看看有沒有 HandTypeID2, HandLevelID2
    '    If HandTypeID = "" AndAlso HandLevelID = "" Then
    '        HandTypeID = Convert.ToString(dr("HandTypeID2"))
    '        HandTypeID = Convert.ToString(If(v_SHIFTSORT = "1", dr("HandTypeID2"), TIMS.Get_HandTypeName2(HandTypeID)))
    '        HandLevelID = Convert.ToString(dr("HandLevelID2"))
    '        HandLevelID = TIMS.ClearSQM(If(v_SHIFTSORT = "1", dr("HandLevelID2"), TIMS.Get_HandLevelName2(HandLevelID)))
    '    End If

    '    EmergencyContact = dr("EmergencyContact").ToString()
    '    EmergencyRelation = dr("EmergencyRelation").ToString()

    '    '42~50
    '    EmergencyPhone = dr("EmergencyPhone").ToString()
    '    ZipCode3 = Convert.ToString(dr("ZipCode3"))
    '    If dr("ZipCode3").ToString <> "" Then
    '        'ZipCode3=dr("ZipCode3").ToString()
    '        If v_SHIFTSORT = "2" Then
    '            If dr("CTName3").ToString <> dr("ZipName3").ToString Then
    '                ZipCode3 = "(" & dr("ZipCode3").ToString & ")" & dr("CTName3").ToString & dr("ZipName3").ToString
    '            Else
    '                ZipCode3 = "(" & dr("ZipCode3").ToString & ")" & dr("CTName3").ToString
    '            End If
    '        End If
    '    End If
    '    Address3 = dr("Address3").ToString()
    '    'If rblWorkMode.SelectedValue="1" Then Address3=TIMS.strMask(Address3, 3)

    '    PriorWorkOrg1 = dr("PriorWorkOrg1").ToString()
    '    Title1 = dr("Title1").ToString
    '    SOfficeYM1 = TIMS.Cdate3(dr("SOfficeYM1"))
    '    FOfficeYM1 = TIMS.Cdate3(dr("FOfficeYM1"))

    '    PriorWorkOrg2 = dr("PriorWorkOrg2").ToString()
    '    Title2 = dr("Title2").ToString

    '    SOfficeYM2 = TIMS.Cdate3(dr("SOfficeYM2"))
    '    '51~56
    '    FOfficeYM2 = TIMS.Cdate3(dr("FOfficeYM2"))

    '    PriorWorkPay = dr("PriorWorkPay").ToString

    '    Traffic = Convert.ToString(If(v_SHIFTSORT = "1", dr("Traffic"), dr("TrafficName")))
    '    RealJobless = dr("RealJobless").ToString()
    '    JoblessID = Convert.ToString(If(v_SHIFTSORT = "1", dr("JoblessID"), dr("JoblessName")))
    '    ShowDetail = If(v_SHIFTSORT = "1", Convert.ToString(dr("ShowDetail")), If(Convert.ToString(dr("ShowDetail")) = "Y", "是", "否"))
    '    LevelNo = dr("LevelNo").ToString
    '    EnterChannel = Convert.ToString(If(v_SHIFTSORT = "1", dr("EnterChannel"), dr("EnterChannelName")))
    '    TRNDMode = Convert.ToString(If(v_SHIFTSORT = "1", dr("TRNDMode"), dr("TRNDModeName")))
    '    TRNDType = Convert.ToString(If(v_SHIFTSORT = "1", dr("TRNDType"), dr("TRNDTypeName")))
    '    BudgetID = Convert.ToString(If(v_SHIFTSORT = "1", dr("BudgetID"), dr("BudName")))
    '    SupplyID = dr("SupplyID").ToString
    '    IsAgree = dr("IsAgree").ToString
    '    PMode = dr("PMode").ToString
    '    ForeName = dr("ForeName").ToString
    '    ForeTitle = dr("ForeTitle").ToString
    '    ForeSex = Convert.ToString(If(v_SHIFTSORT = "1", dr("ForeSex"), dr("ForeSexName")))
    '    ForeBirth = TIMS.Cdate3(dr("ForeBirth"))
    '    ForeIDNO = TIMS.ChangeIDNO(dr("ForeIDNO").ToString)
    '    'Dim v_shiftsort As String=TIMS.GetListValue(shiftsort)
    '    ForeZip = Convert.ToString(If(v_SHIFTSORT = "1", dr("ForeZip"), dr("ForeZipName")))
    '    ForeAddr = dr("ForeAddr").ToString

    '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        '企訓專用
    '        Dim iSOCID As Integer = Val(dr("SOCID"))
    '        Dim dr1 As DataRow = TIMS.GET_STUD_SERVICEPLACE(iSOCID, oConn)

    '        AcctMode = ""
    '        AcctHeadNo = ""
    '        AcctExNo = ""
    '        PostNo = ""
    '        AcctNo = ""
    '        BankName = ""
    '        ExBankName = ""
    '        Uname = ""
    '        Intaxno = ""
    '        ServDept = ""
    '        JobTitle = ""
    '        Zip = ""
    '        Addr = ""
    '        Tel = ""
    '        Fax = ""
    '        SDate = ""
    '        SJDate = ""
    '        SPDate = ""

    '        If dr1 IsNot Nothing Then
    '            If Not IsDBNull(dr1("AcctMode")) Then
    '                '**by Milor AMU 20080509--撥款方式多加入2-訓練單位代轉現金 start
    '                Select Case Convert.ToString(dr1("AcctMode"))
    '                    Case "0"
    '                        AcctMode = If(v_SHIFTSORT = "1", Convert.ToString(dr1("AcctMode")), "郵政")
    '                        PostNo = dr1("PostNo").ToString
    '                        AcctNo = dr1("AcctNo").ToString
    '                    Case "1"
    '                        AcctMode = If(v_SHIFTSORT = "1", Convert.ToString(dr1("AcctMode")), "金融")
    '                        AcctHeadNo = dr1("AcctHeadNo").ToString
    '                        AcctExNo = dr1("AcctExNo").ToString
    '                        AcctNo = dr1("AcctNo").ToString
    '                        BankName = dr1("BankName").ToString
    '                        ExBankName = dr1("ExBankName").ToString
    '                    Case "2"
    '                        AcctMode = If(v_SHIFTSORT = "1", Convert.ToString(dr1("AcctMode")), "訓練單位代轉現金")
    '                End Select
    '                '**by Milor AMU 20080509 end
    '            End If

    '            FirDate = If(Convert.ToString(dr1("FirDate")) <> "", TIMS.Cdate3(dr1("FirDate")), "")
    '            Uname = dr1("Uname").ToString
    '            Intaxno = dr1("Intaxno").ToString
    '            ServDept = dr1("ServDept").ToString
    '            JobTitle = dr1("JobTitle").ToString
    '            'Zip="(" & dr1("Zip").ToString & ")" & dr1("CTName").ToString & dr1("ZipName").ToString
    '            Zip = "(" & dr1("Zip2").ToString & ")" & dr1("CTName").ToString & dr1("ZipName").ToString '原郵遞區號3碼+2碼
    '            Addr = dr1("Addr").ToString
    '            Tel = dr1("Tel").ToString
    '            Fax = dr1("Fax").ToString
    '            SDate = If(Convert.ToString(dr1("SDate")) <> "", TIMS.Cdate3(dr1("SDate")), "")
    '            SJDate = If(Convert.ToString(dr1("SJDate")) <> "", TIMS.Cdate3(dr1("SJDate")), "")
    '            SPDate = If(Convert.ToString(dr1("SPDate")) <> "", TIMS.Cdate3(dr1("SPDate")), "")
    '        End If

    '        Q1 = ""
    '        Q3 = ""
    '        Q3_Other = ""
    '        Q4 = ""
    '        Q5 = ""
    '        Q61 = ""
    '        Q62 = ""
    '        Q63 = ""
    '        Q64 = ""

    '        Call TIMS.OpenDbConn(oConn)
    '        Dim sql2 As String = "SELECT * FROM STUD_TRAINBG WITH(NOLOCK) WHERE SOCID=@SOCID"
    '        Dim sCmd2 As New SqlCommand(sql2, oConn)
    '        With sCmd2
    '            .Parameters.Clear()
    '            .Parameters.Add("SOCID", SqlDbType.BigInt).Value = iSOCID
    '        End With
    '        Dim dr2 As DataRow = TIMS.GetOneRow(sCmd2, oConn)
    '        If dr2 IsNot Nothing Then
    '            Q1 = If(dr2("Q1") = 1, "Y", "N")
    '            Q3 = dr2("Q3").ToString
    '            Q3_Other = dr2("Q3_Other").ToString
    '            Q4 = dr2("Q4").ToString
    '            Q5 = If(Not IsDBNull(dr2("Q5")), If(dr2("Q5") = 1, "Y", "N"), "")
    '            Q61 = dr2("Q61").ToString
    '            Q62 = dr2("Q62").ToString
    '            Q63 = dr2("Q63").ToString
    '            Q64 = dr2("Q64").ToString
    '        End If

    '        Dim sql3 As String = "SELECT * FROM STUD_TRAINBGQ2 WITH(NOLOCK) WHERE SOCID=@SOCID"
    '        Dim sCmd3 As New SqlCommand(sql3, oConn)
    '        Dim dt3 As New DataTable
    '        With sCmd3
    '            .Parameters.Clear()
    '            .Parameters.Add("SOCID", SqlDbType.BigInt).Value = iSOCID
    '            dt3.Load(.ExecuteReader())
    '        End With
    '        Q2 = ""
    '        For Each dr3 As DataRow In dt3.Rows
    '            Q2 &= String.Concat(If(Q2 <> "", "，", ""), dr3("Q2"))
    '        Next
    '    End If

    '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        '企訓專用(產投) '~\SD\03\Sample2.xls  / '~\SD\03\Sample28.xls
    '        '學號	中文姓名	LastName	FirstName	身分證字號	性別	最高學歷	畢業狀況
    '        '聯絡電話_日	聯絡電話_夜	行動電話	主要參訓身分別	開訓日期	結訓日期	報到日期
    '        t_sql = ""
    '        t_sql &= String.Concat("INSERT INTO [Sheet1$] (", "學號,中文姓名,LastName,FirstName", ",身分證字號,性別,最高學歷,畢業狀況")
    '        t_sql &= String.Concat(",聯絡電話_日,聯絡電話_夜,行動電話,主要參訓身分別", ",開訓日期,結訓日期,報到日期)")
    '        t_sql &= "VALUES ("
    '        t_sql &= String.Format("'{0}','{1}','{2}','{3}'", StudentID, Name, LastName, FirstName)
    '        t_sql &= String.Format(",'{0}','{1}','{2}','{3}'", IDNO, SEX, DegreeID, GradID)
    '        t_sql &= String.Format(",'{0}','{1}','{2}','{3}'", PhoneD, PhoneN, CellPhone, MIdentityID)
    '        t_sql &= String.Format(",'{0}','{1}','{2}'", OpenDate, CloseDate, EnterDate)
    '        t_sql &= ")"
    '    Else
    '        'TIMS '~\SD\03\Sample.xls  / '~\SD\03\Sample06.xls
    '        '學號	中文姓名	LastName	FirstName	
    '        '身分證字號	性別 '身分別	非本國人身分別 '原屬國籍	護照或工作證號	最高學歷 '畢業狀況	
    '        '聯絡電話_日	聯絡電話_夜	手機	主要參訓身分別	開訓日期	結訓日期	報到日期
    '        t_sql = ""
    '        t_sql &= String.Concat("INSERT INTO [Sheet1$] (", "學號,中文姓名,LastName,FirstName", ",身分證字號,性別,身分別,非本國人身分別", ",原屬國籍,護照或工作證號,最高學歷,畢業狀況")
    '        t_sql &= String.Concat(",聯絡電話_日,聯絡電話_夜,手機,主要參訓身分別", ",開訓日期,結訓日期,報到日期)")
    '        t_sql &= "VALUES ("
    '        t_sql &= String.Format("'{0}','{1}','{2}','{3}'", StudentID, Name, LastName, FirstName)
    '        t_sql &= String.Format(",'{0}','{1}','{2}','{3}'", IDNO, SEX, PassPortNO, ChinaOrNot)
    '        t_sql &= String.Format(",'{0}','{1}','{2}','{3}'", Nationality, PPNO, DegreeID, GradID)
    '        t_sql &= String.Format(",'{0}','{1}','{2}','{3}'", PhoneD, PhoneN, CellPhone, MIdentityID)
    '        t_sql &= String.Format(",'{0}','{1}','{2}'", OpenDate, CloseDate, EnterDate)
    '        t_sql &= ")"
    '    End If

    '    Dim strErrmsg As String = ""
    '    Using OleCmd1 As New OleDb.OleDbCommand(t_sql, MyConn)
    '        Try
    '            If MyConn.State = ConnectionState.Closed Then MyConn.Open()
    '            OleCmd1.ExecuteNonQuery()
    '            'If conn.State=ConnectionState.Open Then conn.Close()
    '        Catch ex As Exception
    '            strErrmsg = String.Concat("程式錯誤!!!", vbCrLf, "t_sql:", t_sql, vbCrLf, "ex.Message:", ex.Message, vbCrLf)
    '            strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入 'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '            Call TIMS.WriteTraceLog(strErrmsg, ex)

    '            If MyConn IsNot Nothing Then
    '                If MyConn.State = ConnectionState.Open Then MyConn.Close()
    '                MyConn = Nothing
    '            End If
    '            Throw ex
    '        End Try
    '    End Using
    'End Sub

    'Public Shared Sub ExcelOutColumnBk1(ByRef MyPage As Page, ByRef sm As SessionModel, ByRef oConn As SqlConnection,
    '                                 ByRef MyConn As OleDb.OleDbConnection, ByRef dtIdentity As DataTable,
    '                                 ByRef flgCIShow As Boolean, ByRef hPP As Hashtable, ByRef dr As DataRow, ByRef t_sql As String)
    '    Dim v_shiftsort As String=TIMS.GetMyValue2(hPP, "v_shiftsort")

    '    Dim StudentID As String=""
    '    Dim Name As String=""
    '    Dim LastName As String=""
    '    Dim FirstName As String=""
    '    Dim IDNO As String=""
    '    Dim IDNO_MK As String=""
    '    Dim SEX As String=""
    '    Dim PassPortNO As String=""
    '    Dim ChinaOrNot As String=""
    '    Dim Nationality As String=""
    '    Dim PPNO As String=""
    '    Dim Birthday As String=""
    '    Dim MaritalStatus As String=""
    '    Dim DegreeID As String=""
    '    Dim School As String=""
    '    Dim Department As String=""
    '    Dim GradID As String=""
    '    Dim MilitaryID As String=""
    '    Dim ServiceID As String=""
    '    Dim MilitaryAppointment As String=""
    '    Dim MilitaryRank As String=""
    '    Dim ServiceOrg As String=""
    '    Dim ChiefRankName As String=""
    '    Dim ServicePhone As String=""
    '    Dim SServiceDate As String=""
    '    Dim FServiceDate As String=""
    '    Dim ZipCode4 As String=""
    '    Dim Address4 As String=""
    '    Dim JobState As String=""
    '    Dim PhoneD As String=""
    '    Dim PhoneN As String=""
    '    Dim CellPhone As String=""
    '    Dim ZipCode1 As String=""
    '    Dim Address1 As String=""
    '    Dim ZipCode2 As String=""
    '    Dim Address2 As String=""
    '    Dim Email As String=""
    '    Dim IdentityID As String=""
    '    Dim MIdentityID As String=""
    '    Dim SubsidyID As String=""
    '    Dim OpenDate As String=""
    '    Dim CloseDate As String=""
    '    Dim EnterDate As String=""
    '    Dim HandTypeID As String=""
    '    Dim HandLevelID As String=""
    '    Dim EmergencyContact As String=""
    '    Dim EmergencyRelation As String=""
    '    Dim EmergencyPhone As String=""
    '    Dim ZipCode3 As String=""
    '    Dim Address3 As String=""
    '    Dim PriorWorkOrg1 As String=""
    '    Dim Title1 As String=""
    '    Dim SOfficeYM1 As String=""
    '    Dim FOfficeYM1 As String=""
    '    Dim PriorWorkOrg2 As String=""
    '    Dim Title2 As String=""
    '    Dim SOfficeYM2 As String=""
    '    Dim FOfficeYM2 As String=""
    '    Dim PriorWorkPay As String=""
    '    Dim Traffic As String=""
    '    Dim RealJobless As String=""
    '    Dim JoblessID As String=""
    '    Dim ShowDetail As String=""
    '    Dim LevelNo As String=""
    '    Dim EnterChannel As String=""
    '    Dim TRNDMode As String=""
    '    Dim TRNDType As String=""
    '    Dim BudgetID As String=""
    '    Dim SupplyID As String=""
    '    Dim IsAgree As String=""
    '    Dim PMode As String=""
    '    Dim ForeName As String=""
    '    Dim ForeTitle As String=""
    '    Dim ForeSex As String=""
    '    Dim ForeBirth As String=""
    '    Dim ForeIDNO As String=""
    '    Dim ForeZip As String=""
    '    Dim ForeAddr As String=""
    '    Dim KNID As String=""

    '    Dim AcctMode As String=""
    '    Dim PostNo As String=""
    '    Dim AcctHeadNo As String=""
    '    Dim AcctExNo As String=""
    '    Dim AcctNo As String=""
    '    Dim BankName As String=""
    '    Dim ExBankName As String=""
    '    Dim FirDate As String=""
    '    Dim Uname As String=""
    '    Dim Intaxno As String=""
    '    Dim Tel As String=""
    '    Dim Fax As String=""
    '    Dim Zip As String=""
    '    Dim Addr As String=""
    '    Dim ServDept As String=""
    '    Dim JobTitle As String=""
    '    Dim SDate As String=""
    '    Dim SJDate As String=""
    '    Dim SPDate As String=""
    '    Dim Q1 As String=""
    '    Dim Q2 As String=""
    '    Dim Q3 As String=""
    '    Dim Q3_Other As String=""
    '    Dim Q4 As String=""
    '    Dim Q5 As String=""
    '    Dim Q61 As String=""
    '    Dim Q62 As String=""
    '    Dim Q63 As String=""
    '    Dim Q64 As String=""
    '    '出生日期、學校名稱、科系、通訊地址、電子郵件帳號、參訓身分別、障礙類別、障礙等級、緊急通知人姓名、緊急通知人關係、緊急通知人電話、
    '    '緊急通知人地址郵遞區號、緊急通知人地址、報名管道、個資法意願、撥款方式、郵政_局號、金融_總代號、金融_分支代號、帳號、銀行名稱、分行名稱、
    '    '第一次投保勞保日、公司名稱、統編、公司電話、公司傳真、公司地址郵遞區號、公司地址、目前任職部門、職稱、個人到任目前任職公司起日、
    '    '個人到任目前職務起日、最近升遷日期、是否由公司推薦參訓、參訓動機、訓後動向、訓後動向其他說明、服務單位行業別、服務單位是否屬於中小企業、
    '    '個人工作年資、在這家公司的年資、在這職位的年資、最近升遷離本職幾年、是否提供基本資料查詢、戶籍地址郵遞區號、戶籍地址、預算別(經費來源)、補助比例。

    '    StudentID=Right(dr("StudentID").ToString, 2).Replace("'", "")
    '    Name=dr("Name").ToString().Replace("'", "")
    '    LastName=Convert.ToString(dr("EngName")).Replace("'", "")
    '    FirstName=""
    '    If dr("EngName").ToString <> "" Then
    '        If dr("EngName").ToString.IndexOf(" ") <> -1 Then
    '            LastName=Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
    '            FirstName=Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - dr("EngName").ToString.IndexOf(" ") - 1)).Replace("'", "")
    '        End If
    '    End If
    '    IDNO=TIMS.ChangeIDNO(dr("IDNO").ToString())
    '    If Not flgCIShow Then IDNO=TIMS.strMask(IDNO, 1) '不顯示個資。

    '    'If rblWorkMode.SelectedValue="1" Then IDNO=TIMS.strMask(IDNO, 1)
    '    SEX=Convert.ToString(If(v_shiftsort="1", dr("SEX"), dr("SEXName")))
    '    'PassPortNO=dr("PassPortNO").ToString

    '    PassPortNO=Convert.ToString(If(v_shiftsort="1", dr("PassPortNO"), dr("PassPortName"))).Replace("'", "")
    '    ChinaOrNot=Convert.ToString(If(v_shiftsort="1", dr("ChinaOrNot"), dr("ChinaOrNotName"))).Replace("'", "")
    '    Nationality=Convert.ToString(dr("Nationality")).Replace("'", "")
    '    PPNO=Convert.ToString(If(v_shiftsort="1", dr("PPNO"), dr("PPNOName"))).Replace("'", "")

    '    Birthday=""
    '    If Convert.ToString(dr("Birthday")) <> "" Then
    '        Birthday=TIMS.Cdate3(dr("Birthday")) 'DateFormat.ShortDate
    '        If Not flgCIShow Then Birthday=TIMS.strMask(Birthday, 2) '不顯示個資。
    '    End If
    '    MaritalStatus=Convert.ToString(If(v_shiftsort="1", dr("MaritalStatus"), dr("MaritalStatusName")))
    '    '最高學歷
    '    DegreeID=Convert.ToString(If(v_shiftsort="1", dr("DegreeID"), dr("DegreeName")))
    '    School=dr("School").ToString()

    '    '11~20
    '    Department=dr("Department").ToString()
    '    '畢業狀況
    '    GradID=Convert.ToString(If(v_shiftsort="1", dr("GradID"), dr("GradName"))).Replace("'", "")
    '    MilitaryID=Convert.ToString(If(v_shiftsort="1", dr("MilitaryID"), dr("MilitaryName"))).Replace("'", "")

    '    ServiceID=dr("ServiceID").ToString
    '    MilitaryAppointment=dr("MilitaryAppointment").ToString
    '    MilitaryRank=dr("MilitaryRank").ToString
    '    ServiceOrg=dr("ServiceOrg").ToString
    '    ChiefRankName=dr("ChiefRankName").ToString
    '    ServicePhone=dr("ServicePhone").ToString
    '    SServiceDate=dr("SServiceDate").ToString

    '    '21~30
    '    FServiceDate=dr("FServiceDate").ToString
    '    ZipCode4=Convert.ToString(dr("ZipCode4"))
    '    If dr("ZipCode4").ToString <> "" Then
    '        'ZipCode4=dr("ZipCode4").ToString
    '        If v_shiftsort="2" Then
    '            If dr("CTName4").ToString <> dr("ZipName4").ToString Then
    '                ZipCode4="(" & dr("ZipCode4").ToString & ")" & dr("CTName4").ToString & dr("ZipName4").ToString
    '            Else
    '                ZipCode4="(" & dr("ZipCode4").ToString & ")" & dr("CTName4").ToString
    '            End If
    '        End If
    '    End If
    '    Address4=dr("Address4").ToString

    '    JobState=Convert.ToString(If(v_shiftsort="1", dr("JobState"), dr("JobStateName")))

    '    'flgCIShow : 可顯示個資 / 不顯示個資。
    '    PhoneD=If(flgCIShow, dr("PhoneD").ToString().Replace("'", ""), "****")
    '    PhoneN=If(flgCIShow, dr("PhoneN").ToString().Replace("'", ""), "****")
    '    CellPhone=If(flgCIShow, dr("CellPhone").ToString().Replace("'", ""), "****")

    '    ZipCode1=Convert.ToString(dr("ZipCode1"))
    '    If dr("ZipCode1").ToString <> "" Then
    '        'ZipCode1=dr("ZipCode1").ToString()
    '        If v_shiftsort="2" Then
    '            If dr("CTName1").ToString <> dr("ZipName1").ToString Then
    '                ZipCode1="(" & dr("ZipCode1").ToString & ")" & dr("CTName1").ToString & dr("ZipName1").ToString
    '            Else
    '                ZipCode1="(" & dr("ZipCode1").ToString & ")" & dr("CTName1").ToString
    '            End If
    '        End If
    '    End If

    '    Address1=dr("Address1").ToString()
    '    'If rblWorkMode.SelectedValue="1" Then Address1=TIMS.strMask(Address1, 3)
    '    ZipCode2=Convert.ToString(dr("ZipCode2"))
    '    If dr("ZipCode2").ToString <> "" Then
    '        'ZipCode2=dr("ZipCode2").ToString()
    '        If v_shiftsort="2" Then
    '            If dr("CTName2").ToString <> dr("ZipName2").ToString Then
    '                ZipCode2="(" & dr("ZipCode2").ToString & ")" & dr("CTName2").ToString & dr("ZipName2").ToString
    '            Else
    '                ZipCode2="(" & dr("ZipCode2").ToString & ")" & dr("CTName2").ToString
    '            End If
    '        End If
    '    End If
    '    Address2=dr("Address2").ToString()
    '    'If rblWorkMode.SelectedValue="1" Then Address2=TIMS.strMask(Address2, 3)

    '    '31~41
    '    Email=dr("Email").ToString()
    '    IdentityID=dr("IdentityID").ToString
    '    MIdentityID=dr("MIdentityID").ToString
    '    If v_shiftsort="1" Then
    '        IdentityID=Replace(dr("IdentityID").ToString, ",", "，")
    '    Else
    '        IdentityID=TIMS.Get_IdentityName(dr("IdentityID").ToString, dtIdentity, "，")
    '        MIdentityID=TIMS.Get_IdentityName(dr("IdentityID").ToString, dtIdentity, "，")
    '    End If
    '    'by Vicient
    '    KNID=Convert.ToString(If(v_shiftsort="1", dr("KNID"), dr("knName")))

    '    SubsidyID=Convert.ToString(If(v_shiftsort="1", dr("SubsidyID"), dr("SubsidyName")))

    '    OpenDate=TIMS.Cdate3(dr("OpenDate"))
    '    CloseDate=TIMS.Cdate3(dr("CloseDate"))
    '    EnterDate=TIMS.Cdate3(dr("EnterDate"))

    '    HandTypeID=TIMS.ClearSQM(If(v_shiftsort="1", dr("HandTypeID"), dr("HandTypeName")))
    '    HandLevelID=TIMS.ClearSQM(If(v_shiftsort="1", dr("HandLevelID"), dr("HandLevelName")))
    '    'If HandTypeID <> "" Then HandTypeID=Trim(HandTypeID)
    '    'If HandLevelID <> "" Then HandLevelID=Trim(HandLevelID)
    '    '若是都空白看看有沒有 HandTypeID2, HandLevelID2
    '    If HandTypeID="" AndAlso HandLevelID="" Then
    '        HandTypeID=Convert.ToString(dr("HandTypeID2"))
    '        HandTypeID=Convert.ToString(If(v_shiftsort="1", dr("HandTypeID2"), TIMS.Get_HandTypeName2(HandTypeID)))
    '        HandLevelID=Convert.ToString(dr("HandLevelID2"))
    '        HandLevelID=TIMS.ClearSQM(If(v_shiftsort="1", dr("HandLevelID2"), TIMS.Get_HandLevelName2(HandLevelID)))
    '    End If

    '    EmergencyContact=dr("EmergencyContact").ToString()
    '    EmergencyRelation=dr("EmergencyRelation").ToString()

    '    '42~50
    '    EmergencyPhone=dr("EmergencyPhone").ToString()
    '    ZipCode3=Convert.ToString(dr("ZipCode3"))
    '    If dr("ZipCode3").ToString <> "" Then
    '        'ZipCode3=dr("ZipCode3").ToString()
    '        If v_shiftsort="2" Then
    '            If dr("CTName3").ToString <> dr("ZipName3").ToString Then
    '                ZipCode3="(" & dr("ZipCode3").ToString & ")" & dr("CTName3").ToString & dr("ZipName3").ToString
    '            Else
    '                ZipCode3="(" & dr("ZipCode3").ToString & ")" & dr("CTName3").ToString
    '            End If
    '        End If
    '    End If
    '    Address3=dr("Address3").ToString()
    '    'If rblWorkMode.SelectedValue="1" Then Address3=TIMS.strMask(Address3, 3)

    '    PriorWorkOrg1=dr("PriorWorkOrg1").ToString()
    '    Title1=dr("Title1").ToString
    '    SOfficeYM1=TIMS.Cdate3(dr("SOfficeYM1"))
    '    FOfficeYM1=TIMS.Cdate3(dr("FOfficeYM1"))

    '    PriorWorkOrg2=dr("PriorWorkOrg2").ToString()
    '    Title2=dr("Title2").ToString

    '    SOfficeYM2=TIMS.Cdate3(dr("SOfficeYM2"))
    '    '51~56
    '    FOfficeYM2=TIMS.Cdate3(dr("FOfficeYM2"))

    '    PriorWorkPay=dr("PriorWorkPay").ToString

    '    Traffic=Convert.ToString(If(v_shiftsort="1", dr("Traffic"), dr("TrafficName")))
    '    RealJobless=dr("RealJobless").ToString()
    '    JoblessID=Convert.ToString(If(v_shiftsort="1", dr("JoblessID"), dr("JoblessName")))
    '    ShowDetail=If(v_shiftsort="1", Convert.ToString(dr("ShowDetail")), If(Convert.ToString(dr("ShowDetail"))="Y", "是", "否"))
    '    LevelNo=dr("LevelNo").ToString
    '    EnterChannel=Convert.ToString(If(v_shiftsort="1", dr("EnterChannel"), dr("EnterChannelName")))
    '    TRNDMode=Convert.ToString(If(v_shiftsort="1", dr("TRNDMode"), dr("TRNDModeName")))
    '    TRNDType=Convert.ToString(If(v_shiftsort="1", dr("TRNDType"), dr("TRNDTypeName")))
    '    BudgetID=Convert.ToString(If(v_shiftsort="1", dr("BudgetID"), dr("BudName")))
    '    SupplyID=dr("SupplyID").ToString
    '    IsAgree=dr("IsAgree").ToString
    '    PMode=dr("PMode").ToString
    '    ForeName=dr("ForeName").ToString
    '    ForeTitle=dr("ForeTitle").ToString
    '    ForeSex=Convert.ToString(If(v_shiftsort="1", dr("ForeSex"), dr("ForeSexName")))
    '    ForeBirth=TIMS.Cdate3(dr("ForeBirth"))
    '    ForeIDNO=TIMS.ChangeIDNO(dr("ForeIDNO").ToString)
    '    'Dim v_shiftsort As String=TIMS.GetListValue(shiftsort)
    '    ForeZip=Convert.ToString(If(v_shiftsort="1", dr("ForeZip"), dr("ForeZipName")))
    '    ForeAddr=dr("ForeAddr").ToString

    '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        '企訓專用
    '        Dim iSOCID As Integer=Val(dr("SOCID"))
    '        Dim dr1 As DataRow=TIMS.GET_STUD_SERVICEPLACE(iSOCID, oConn)

    '        AcctMode=""
    '        AcctHeadNo=""
    '        AcctExNo=""
    '        PostNo=""
    '        AcctNo=""
    '        BankName=""
    '        ExBankName=""
    '        Uname=""
    '        Intaxno=""
    '        ServDept=""
    '        JobTitle=""
    '        Zip=""
    '        Addr=""
    '        Tel=""
    '        Fax=""
    '        SDate=""
    '        SJDate=""
    '        SPDate=""

    '        If dr1 IsNot Nothing Then
    '            If Not IsDBNull(dr1("AcctMode")) Then
    '                '**by Milor AMU 20080509--撥款方式多加入2-訓練單位代轉現金 start
    '                Select Case Convert.ToString(dr1("AcctMode"))
    '                    Case "0"
    '                        AcctMode=If(v_shiftsort="1", Convert.ToString(dr1("AcctMode")), "郵政")
    '                        PostNo=dr1("PostNo").ToString
    '                        AcctNo=dr1("AcctNo").ToString
    '                    Case "1"
    '                        AcctMode=If(v_shiftsort="1", Convert.ToString(dr1("AcctMode")), "金融")
    '                        AcctHeadNo=dr1("AcctHeadNo").ToString
    '                        AcctExNo=dr1("AcctExNo").ToString
    '                        AcctNo=dr1("AcctNo").ToString
    '                        BankName=dr1("BankName").ToString
    '                        ExBankName=dr1("ExBankName").ToString
    '                    Case "2"
    '                        AcctMode=If(v_shiftsort="1", Convert.ToString(dr1("AcctMode")), "訓練單位代轉現金")
    '                End Select
    '                '**by Milor AMU 20080509 end
    '            End If

    '            FirDate=If(Convert.ToString(dr1("FirDate")) <> "", TIMS.Cdate3(dr1("FirDate")), "")
    '            Uname=dr1("Uname").ToString
    '            Intaxno=dr1("Intaxno").ToString
    '            ServDept=dr1("ServDept").ToString
    '            JobTitle=dr1("JobTitle").ToString
    '            'Zip="(" & dr1("Zip").ToString & ")" & dr1("CTName").ToString & dr1("ZipName").ToString
    '            Zip="(" & dr1("Zip2").ToString & ")" & dr1("CTName").ToString & dr1("ZipName").ToString '原郵遞區號3碼+2碼
    '            Addr=dr1("Addr").ToString
    '            Tel=dr1("Tel").ToString
    '            Fax=dr1("Fax").ToString
    '            SDate=If(Convert.ToString(dr1("SDate")) <> "", TIMS.Cdate3(dr1("SDate")), "")
    '            SJDate=If(Convert.ToString(dr1("SJDate")) <> "", TIMS.Cdate3(dr1("SJDate")), "")
    '            SPDate=If(Convert.ToString(dr1("SPDate")) <> "", TIMS.Cdate3(dr1("SPDate")), "")
    '        End If

    '        Q1=""
    '        Q3=""
    '        Q3_Other=""
    '        Q4=""
    '        Q5=""
    '        Q61=""
    '        Q62=""
    '        Q63=""
    '        Q64=""

    '        Call TIMS.OpenDbConn(oConn)
    '        Dim sql2 As String="SELECT * FROM STUD_TRAINBG WITH(NOLOCK) WHERE SOCID=@SOCID"
    '        Dim sCmd2 As New SqlCommand(sql2, oConn)
    '        With sCmd2
    '            .Parameters.Clear()
    '            .Parameters.Add("SOCID", SqlDbType.BigInt).Value=iSOCID
    '        End With
    '        Dim dr2 As DataRow=TIMS.GetOneRow(sCmd2, oConn)
    '        'sql=String.Format(" SELECT * FROM STUD_TRAINBG WITH(NOLOCK) WHERE SOCID={0} ", iSOCID)
    '        'Dim dr2 As DataRow=DbAccess.GetOneRow(sql, oConn)
    '        If dr2 IsNot Nothing Then
    '            Q1=If(dr2("Q1")=1, "Y", "N")
    '            Q3=dr2("Q3").ToString
    '            Q3_Other=dr2("Q3_Other").ToString
    '            Q4=dr2("Q4").ToString
    '            Q5=If(Not IsDBNull(dr2("Q5")), If(dr2("Q5")=1, "Y", "N"), "")
    '            Q61=dr2("Q61").ToString
    '            Q62=dr2("Q62").ToString
    '            Q63=dr2("Q63").ToString
    '            Q64=dr2("Q64").ToString
    '        End If

    '        Dim sql3 As String="SELECT * FROM STUD_TRAINBGQ2 WITH(NOLOCK) WHERE SOCID=@SOCID"
    '        Dim sCmd3 As New SqlCommand(sql3, oConn)
    '        Dim dt3 As New DataTable
    '        With sCmd3
    '            .Parameters.Clear()
    '            .Parameters.Add("SOCID", SqlDbType.BigInt).Value=iSOCID
    '            dt3.Load(.ExecuteReader())
    '        End With
    '        Q2=""
    '        For Each dr3 As DataRow In dt3.Rows
    '            Q2 &= String.Concat(If(Q2 <> "", "，", ""), dr3("Q2"))
    '        Next
    '    End If

    '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        '企訓專用(產投) '~\SD\03\Sample2.xls
    '        t_sql=""
    '        t_sql &= String.Concat("INSERT INTO [Sheet1$] (", "學號,中文姓名,LastName,FirstName,身分證字號,性別,")
    '        t_sql &= "出生日期,最高學歷,學校名稱,"
    '        t_sql &= "科系,畢業狀況,"
    '        t_sql &= "聯絡電話_日,聯絡電話_夜,"
    '        t_sql &= "行動電話,通訊地址郵遞區號,通訊地址,"
    '        t_sql &= "電子郵件帳號,參訓身分別,主要參訓身分別,開訓日期,結訓日期,"
    '        t_sql &= "報到日期,障礙類別,障礙等級,緊急通知人姓名,緊急通知人關係,"
    '        t_sql &= "緊急通知人電話,緊急通知人地址郵遞區號,緊急通知人地址,"
    '        t_sql &= "報名管道,個資法意願,"
    '        t_sql &= "撥款方式,郵政_局號,金融_總代號,金融_分支代號,帳號,銀行名稱,分行名稱,第一次投保勞保日,公司名稱,統編,公司電話,"
    '        t_sql &= "公司傳真,公司地址郵遞區號,公司地址,目前任職部門,職稱,"
    '        t_sql &= "個人到任目前任職公司起日,個人到任目前職務起日,最近升遷日期,是否由公司推薦參訓,參訓動機,"
    '        t_sql &= "訓後動向,訓後動向其他說明,服務單位行業別,服務單位是否屬於中小企業,個人工作年資,"
    '        t_sql &= "在這家公司的年資,在這職位的年資,最近升遷離本職幾年,是否提供基本資料查詢"
    '        t_sql &= String.Concat(",[預算別(經費來源別)],補助比例", ")")
    '        t_sql &= "VALUES ("
    '        t_sql &= "'" & StudentID & "' , '" & Name.Replace("'", "") & "' , '" & LastName.Replace("'", "") & "', '" & FirstName.Replace("'", "") & "' , '" & TIMS.ChangeIDNO(IDNO) & "' , '" & SEX & "' ,"
    '        t_sql &= "'" & Birthday & "' , '" & DegreeID & "' , '" & School.Replace("'", "") & "' ,"
    '        t_sql &= "'" & Department.Replace("'", "") & "' , '" & GradID & "',"
    '        t_sql &= "'" & PhoneD.Replace("'", "") & "' , '" & PhoneN.Replace("'", "") & "' ,"
    '        t_sql &= "'" & CellPhone.Replace("'", "") & "' ,"
    '        t_sql &= "'" & ZipCode1 & "' , '" & Address1.Replace("'", "") & "' ,"
    '        t_sql &= "'" & Email.Replace("'", "") & "' , '" & IdentityID & "' , '" & MIdentityID & "' ,'" & OpenDate & "' , '" & CloseDate & "' ,"
    '        t_sql &= "'" & EnterDate & "' , '" & HandTypeID & "' , '" & HandLevelID & "' , '" & EmergencyContact.Replace("'", "") & "' , '" & EmergencyRelation.Replace("'", "") & "' ,"
    '        t_sql &= "'" & EmergencyPhone.Replace("'", "") & "' , '" & ZipCode3 & "' , '" & Address3.Replace("'", "") & "',"
    '        t_sql &= "'" & EnterChannel & "','" & IsAgree & "',"
    '        t_sql &= "'" & TIMS.ChangeSQM(AcctMode) & "','" & TIMS.ChangeSQM(PostNo) & "','" & TIMS.ChangeSQM(AcctHeadNo) & "','" & TIMS.ChangeSQM(AcctExNo) & "','" & TIMS.ChangeSQM(AcctNo) & "' ,'" & TIMS.ChangeSQM(BankName) & "','" & TIMS.ChangeSQM(ExBankName) & "', '" & TIMS.ChangeSQM(FirDate) & "' , '" & TIMS.ChangeSQM(Uname) & "' , '" & TIMS.ChangeSQM(Intaxno) & "' , '" & TIMS.ChangeSQM(Tel) & "' ,"
    '        t_sql &= "'" & TIMS.ChangeSQM(Fax) & "' , '" & TIMS.ChangeSQM(Zip) & "' , '" & TIMS.ChangeSQM(Addr) & "' , '" & TIMS.ChangeSQM(ServDept) & "' , '" & TIMS.ChangeSQM(JobTitle) & "' ,"
    '        t_sql &= "'" & SDate & "' , '" & SJDate & "' , '" & SPDate & "' , '" & Q1 & "' , '" & Q2 & "' ,"
    '        t_sql &= "'" & Q3 & "' , '" & Q3_Other.Replace("'", "") & "' , '" & Q4 & "' , '" & Q5 & "' , '" & Q61 & "' ,"
    '        t_sql &= "'" & Q62 & "' , '" & Q63 & "' , '" & Q64 & "' , '" & ShowDetail.Replace("'", "") & "'"
    '        t_sql &= ",'" & BudgetID & "' , '" & SupplyID & "'"
    '        t_sql &= ")"
    '    Else
    '        'TIMS '"~\SD\03\Sample.xls
    '        'If FirstName Is Nothing Then FirstName=""
    '        'If LastName Is Nothing Then LastName=""
    '        t_sql=""
    '        t_sql &= String.Concat("INSERT INTO [Sheet1$] (", "學號,中文姓名,LastName,FirstName,身分證字號,性別,")
    '        t_sql &= "身分別,非本國人身分別,原屬國籍,護照或工作證號,出生日期,婚姻狀況,最高學歷,學校名稱,"
    '        t_sql &= "科系,畢業狀況,兵役,軍種,職務_兵役,"
    '        t_sql &= "階級,服務單位名稱,主管階級姓名,單位電話,服役起日期,"
    '        t_sql &= "服役迄日期,服役單位地址郵遞區號,服役單位地址,在職狀況,聯絡電話_日,聯絡電話_夜,"
    '        t_sql &= "手機,通訊地址郵遞區號,通訊地址,戶籍地址郵遞區號,戶籍地址,"
    '        t_sql &= "電子郵件帳號,參訓身分別,主要參訓身分別,津貼類別,開訓日期,結訓日期,"
    '        t_sql &= "報到日期,障礙類別,障礙等級,緊急通知人姓名,緊急通知人關係,"
    '        t_sql &= "緊急通知人電話,緊急通知人地址郵遞區號,緊急通知人地址,受訓前服務單位1,受訓前服務單位1職稱,"
    '        t_sql &= "受訓前服務單位1任職起日,受訓前服務單位1任職迄日,受訓前服務單位2,受訓前服務單位2職稱,受訓前服務單位2任職起日,"
    '        t_sql &= "受訓前服務單位2任職迄日,受訓前薪資,受訓前真正失業週數,受訓前失業週數,交通方式,"
    '        t_sql &= "是否提供基本資料查詢,報名階段,報名管道,推介種類,券別種類,預算別,個資法意願,自費公費, "
    '        t_sql &= String.Concat("國內親屬資料_姓名,國內親屬資料_稱謂,國內親屬資料_性別,國內親屬資料_生日,國內親屬資料_身分證字號,國內親屬資料_郵遞區號,國內親屬資料_地址,原住民民族別", ")")
    '        t_sql &= "VALUES ("
    '        t_sql &= "'" & StudentID & "' , '" & Name.Replace("'", "") & "' , '" & LastName.Replace("'", "") & "', '" & FirstName.Replace("'", "") & "' , '" & TIMS.ChangeIDNO(IDNO) & "' , '" & SEX & "' ,"
    '        t_sql &= "'" & PassPortNO & "' ,'" & ChinaOrNot.Replace("'", "") & "','" & Nationality.Replace("'", "") & "','" & PPNO.Replace("'", "") & "', '" & Birthday & "' , '" & MaritalStatus.Replace("'", "") & "' , '" & DegreeID & "' , '" & School.Replace("'", "") & "' ,"
    '        t_sql &= "'" & Department.Replace("'", "") & "' , '" & GradID & "' , '" & MilitaryID & "' , '" & ServiceID & "' , '" & MilitaryAppointment.Replace("'", "") & "' ,"
    '        t_sql &= "'" & MilitaryRank.Replace("'", "") & "' , '" & ServiceOrg.Replace("'", "") & "' , '" & ChiefRankName.Replace("'", "") & "' , '" & ServicePhone.Replace("'", "") & "' , '" & SServiceDate & "' ,"
    '        t_sql &= "'" & FServiceDate & "' , '" & ZipCode4 & "' , '" & Address4.Replace("'", "") & "' ,'" & JobState.Replace("'", "") & "', '" & PhoneD.Replace("'", "") & "' , '" & PhoneN.Replace("'", "") & "' ,"
    '        t_sql &= "'" & CellPhone.Replace("'", "") & "' , '" & ZipCode1 & "' , '" & Address1.Replace("'", "") & "' , '" & ZipCode2 & "' , '" & Address2.Replace("'", "") & "' ,"
    '        t_sql &= "'" & Email.Replace("'", "") & "' , '" & IdentityID & "' , '" & MIdentityID & "' , '" & SubsidyID & "' , '" & OpenDate & "' , '" & CloseDate & "' ,"
    '        t_sql &= "'" & EnterDate & "' , '" & HandTypeID & "' , '" & HandLevelID & "' , '" & EmergencyContact.Replace("'", "") & "' , '" & EmergencyRelation.Replace("'", "") & "' ,"
    '        t_sql &= "'" & EmergencyPhone.Replace("'", "") & "' , '" & ZipCode3 & "' , '" & Address3.Replace("'", "") & "' , '" & PriorWorkOrg1.Replace("'", "") & "' , '" & Title1.Replace("'", "") & "' ,"
    '        t_sql &= "'" & SOfficeYM1.Replace("'", "") & "' , '" & FOfficeYM1.Replace("'", "") & "' , '" & PriorWorkOrg2.Replace("'", "") & "' , '" & Title2.Replace("'", "") & "' , '" & SOfficeYM2.Replace("'", "") & "' ,"
    '        t_sql &= "'" & FOfficeYM2.Replace("'", "") & "' , '" & PriorWorkPay.Replace("'", "") & "' , '" & RealJobless.Replace("'", "") & "' , '" & JoblessID & "' , '" & Traffic.Replace("'", "") & "' ,"
    '        t_sql &= "'" & ShowDetail.Replace("'", "") & "','" & LevelNo & "','" & EnterChannel & "','" & TRNDMode.Replace("'", "") & "','" & TRNDType.Replace("'", "") & "','" & BudgetID & "','" & IsAgree & "','" & PMode.Replace("'", "") & "',"
    '        t_sql &= "'" & ForeName.Replace("'", "") & "','" & ForeTitle.Replace("'", "") & "','" & ForeSex & "','" & ForeBirth & "','" & TIMS.ChangeIDNO(ForeIDNO) & "','" & ForeZip & "','" & ForeAddr.Replace("'", "") & "','" & KNID & "'"
    '        t_sql &= ")"
    '    End If

    '    Dim strErrmsg As String=""
    '    Using OleCmd1 As New OleDb.OleDbCommand(t_sql, MyConn)
    '        Try
    '            If MyConn.State=ConnectionState.Closed Then MyConn.Open()
    '            OleCmd1.ExecuteNonQuery()
    '            'If conn.State=ConnectionState.Open Then conn.Close()
    '        Catch ex As Exception
    '            strErrmsg=String.Concat("程式錯誤!!!", vbCrLf, "t_sql:", t_sql, vbCrLf, "ex.Message:", ex.Message, vbCrLf)
    '            strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入 'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '            Call TIMS.WriteTraceLog(strErrmsg, ex)

    '            If MyConn IsNot Nothing Then
    '                If MyConn.State=ConnectionState.Open Then MyConn.Close()
    '                MyConn=Nothing
    '            End If
    '            Throw ex
    '        End Try
    '    End Using

    'End Sub

    'Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click  OUTPUT 匯出名冊 (匯出學員資料)
    ' <summary> OUTPUT 匯出名冊 (匯出學員資料) </summary>
    'Sub Export1()
    '    Dim strErrmsg As String = ""

    '    Const cst_SampleXLS28 As String = "~\SD\03\Sample28.xls" '(產投)
    '    Const cst_SampleXLS06 As String = "~\SD\03\Sample06.xls" '(自辦在職)

    '    Dim SD32NM1 As String = TIMS.GetValidFileName(String.Format("SD32_{0}{1}.xls", TIMS.GetDateNo(), TIMS.GetRnd6Eng()))
    '    Dim sFileName As String = String.Format("~\SD\03\Temp\{0}", SD32NM1)
    '    Dim MyPath As String = Server.MapPath(sFileName)
    '    Try
    '        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '            If IO.File.Exists(Server.MapPath(cst_SampleXLS28)) Then
    '                TIMS.MyFileDelete(MyPath)
    '            Else
    '                Common.MessageBox(Me, "Sample檔案不存在")
    '                Return 'Exit Sub
    '            End If
    '            IO.File.Copy(Server.MapPath(cst_SampleXLS28), MyPath, True)
    '            ''除去sample檔的唯讀屬性
    '            'IO.File.SetAttributes(MyPath, IO.FileAttributes.Normal)
    '        Else
    '            If IO.File.Exists(Server.MapPath(cst_SampleXLS06)) Then
    '                TIMS.MyFileDelete(MyPath)
    '            Else
    '                Common.MessageBox(Me, "Sample檔案不存在")
    '                Return 'Exit Sub
    '            End If
    '            IO.File.Copy(Server.MapPath(cst_SampleXLS06), MyPath, True)
    '            '除去sample檔的唯讀屬性
    '            'IO.File.SetAttributes(MyPath, IO.FileAttributes.Normal)
    '        End If
    '    Catch ex As Exception
    '        strErrmsg = String.Concat("程式錯誤!", vbCrLf, "MyPath:", MyPath, vbCrLf, "ex.Message:", ex.Message, vbCrLf)
    '        strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入 'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '        Call TIMS.WriteTraceLog(strErrmsg, ex)

    '        strErrmsg = ""
    '        strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
    '        strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
    '        strErrmsg += ex.ToString & vbCrLf
    '        Common.MessageBox(Me, strErrmsg)
    '        Exit Sub
    '    End Try

    '    If strErrmsg <> "" Then Exit Sub

    '    '除去sample檔的唯讀屬性
    '    'MyFile.SetAttributes(Server.MapPath("~\SD\03\Temp\" & Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", "") & ".xls"), IO.FileAttributes.Normal)
    '    If IO.File.Exists(MyPath) Then IO.File.SetAttributes(MyPath, IO.FileAttributes.Normal)
    '    'copy一份sample資料 End

    '    '根據路徑建立資料庫連線，並取出學員資料填入 Start
    '    Dim v_rblWorkMode As String = TIMS.GetListValue(rblWorkMode) ' .SelectedValue
    '    '1:模糊顯示 2:可正常顯示個資。
    '    flgCIShow = If(v_rblWorkMode = TIMS.cst_wmdip2, True, False)

    '    drOCID = Nothing
    '    If OCIDValue1.Value <> "" Then
    '        drOCID = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
    '        If drOCID Is Nothing Then
    '            Common.MessageBox(Me, "班級查詢有誤。")
    '            Exit Sub
    '        End If
    '        'If Convert.ToString(drOCID("ShowOK14"))="Y" Then flgCIShow=True '可正常顯示個資。
    '    End If
    '    ViewState(cst_flgCIShow) = flgCIShow

    '    Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
    '    If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
    '        If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
    '        Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
    '    End If

    '    Dim dt As DataTable = SEARCH_DATA1_dt()

    '    'Dim Sql_iden As String=" SELECT IDENTITYID ,NAME FROM KEY_IDENTITY WITH(NOLOCK) ORDER BY IDENTITYID"
    '    'Dim dtIdentity As DataTable=DbAccess.GetDataTable(Sql_iden, objconn)

    '    'sMemo=GET_SEARCH_MEMO()
    '    '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
    '    '學號	中文姓名	LastName	FirstName	身分證字號	性別	最高學歷	畢業狀況	聯絡電話_日	聯絡電話_夜	行動電話	主要參訓身分別	開訓日期	結訓日期	報到日期
    '    Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "STUDID,NAME,IDNO,SEX,DEGREENAME,GRADNAME,PHONED,PHONEN,CELLPHONE,MIDENTITYID,OPENDATE,CLOSEDATE,ENTERDATE")
    '    Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, v_rblWorkMode, OCIDValue1.Value, "", objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

    '    If dt.Rows.Count = 0 Then
    '        Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
    '        Exit Sub
    '    End If

    '    Dim v_LOCALADDR As String = TIMS.Get_LOCALADDR(Me, 2)

    '    Dim objLock_ExportXLSX As Object = Nothing

    '    Using MyConn As New OleDb.OleDbConnection
    '        objLock_ExportXLSX = New Object
    '        SyncLock objLock_ExportXLSX
    '            MyConn.ConnectionString = TIMS.Get_OleDbStr(MyPath)
    '            Try
    '                MyConn.Open()
    '            Catch ex As Exception
    '                strErrmsg = String.Concat("(", v_LOCALADDR, ")Excel資料無法開啟連線!", vbCrLf, SD32NM1, vbCrLf, ex.Message)
    '                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入 'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '                Call TIMS.WriteTraceLog(strErrmsg, ex)

    '                strErrmsg = String.Concat("(", v_LOCALADDR, ")Excel資料無法開啟連線!", vbCrLf, SD32NM1, vbCrLf, ex.Message)
    '                Common.MessageBox(Me, strErrmsg)
    '                Exit Sub
    '            End Try
    '            dt.DefaultView.Sort = "StudentID"
    '            'shiftsort:1以代號匯出 'shiftsort:2以名稱匯出
    '            Dim v_SHIFTSORT As String = TIMS.GetListValue(shiftsort)
    '            Dim htPP As New Hashtable From {{"v_SHIFTSORT", v_SHIFTSORT}}
    '            For Each dr As DataRow In dt.Rows
    '                Dim t_sql As String = ""
    '                Call ExcelOutColumn(Page, sm, objconn, MyConn, dtIdentity, flgCIShow, htPP, dr, t_sql)
    '            Next
    '            If MyConn.State = ConnectionState.Open Then MyConn.Close()
    '            '根據路徑建立資料庫連線，並取出學員資料填入 End
    '        End SyncLock
    '    End Using

    '    Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
    '    Select Case V_ExpType
    '        Case "EXCEL"
    '            Dim ClsNM1 As String = TIMS.ChangeIDNO(Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", ""))
    '            Dim MyFileName As String = TIMS.GetValidFileName(TIMS.ClearSQM(String.Concat(ClsNM1, TIMS.GetDateNo(), ".xls")))
    '            Call ExpExccl_1(strErrmsg, MyPath, MyFileName)
    '            '刪除Temp中的資料
    '            Call TIMS.MyFileDelete(MyPath)
    '        Case "ODS"
    '            Using fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
    '                Dim br As New System.IO.BinaryReader(fr)
    '                Dim buf(fr.Length) As Byte
    '                fr.Read(buf, 0, fr.Length)
    '                fr.Close()

    '                '刪除Temp中的資料
    '                Call TIMS.MyFileDelete(MyPath)

    '                Dim sFileName1 As String = TIMS.GetValidFileName(String.Concat("ExpFile", TIMS.GetRnd6Eng()))
    '                'parmsExp.Add("strHTML", strHTML)
    '                Dim parmsExp As New Hashtable From {
    '                    {"ExpType", V_ExpType}, 'EXCEL/PDF/ODS
    '                    {"FileName", sFileName1},
    '                    {"xlsx_buf", buf},
    '                    {"ResponseNoEnd", "Y"}
    '                }
    '                TIMS.Utl_ExportRp1(Me, parmsExp)
    '            End Using

    '        Case Else
    '            '刪除Temp中的資料
    '            Call TIMS.MyFileDelete(MyPath)

    '            Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
    '            Common.MessageBox(Me, s_log1)
    '            Exit Sub
    '    End Select

    '    '刪除Temp中的資料
    '    'Call TIMS.MyFileDelete(MyPath)
    '    Call TIMS.CloseDbConn(objconn)
    '    If strErrmsg <> "" Then
    '        Common.MessageBox(Me, strErrmsg)
    '    Else
    '        'Response.Flush()
    '        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    '    End If
    'End Sub

    '新建立的excel存入記憶體下載
    'Sub ExpExccl_1(ByRef strErrmsg As String, ByRef MyPath As String, ByRef MyFileName As String)
    '    strErrmsg = ""
    '    Try
    '        Using fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
    '            Dim br As New System.IO.BinaryReader(fr)
    '            Dim buf(fr.Length) As Byte
    '            fr.Read(buf, 0, fr.Length)
    '            fr.Close()

    '            '刪除Temp中的資料
    '            Call TIMS.MyFileDelete(MyPath)

    '            Response.Clear()
    '            Response.ClearHeaders()
    '            Response.Buffer = False
    '            Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8))
    '            Response.ContentType = "Application/vnd.ms-Excel"
    '            'Common.RespWrite(Me, br.ReadBytes(fr.Length))
    '            Response.BinaryWrite(buf)
    '            Response.Flush()
    '            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    '        End Using
    '    Catch ex As Exception
    '        'Call TIMS.WriteTraceLog(ex.Message, ex)
    '        TIMS.LOG.Warn(ex.Message, ex)
    '        strErrmsg = String.Concat("無法存取該檔案!!!", vbCrLf, " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)", vbCrLf, ex.Message)
    '    End Try
    'End Sub

    '檢查輸入資料
    'Function CheckImportData(ByVal colArray As Array) As String
    '    Dim Reason As String=""
    '    'Dim SearchEngStr As String="ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
    '    'Dim sql As String=""
    '    'Dim dr As DataRow=Nothing

    '    '企訓專用'產投檢查。
    '    Dim flagTPlanID28a As Boolean=False '(產投 28.54)
    '    Dim flagTIMSNot28a As Boolean=True '(TIMS)
    '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        flagTPlanID28a=True
    '        flagTIMSNot28a=False
    '    End If

    '    '企訓專用
    '    If flagTPlanID28a Then
    '        Reason=CheckImportData28(colArray)
    '        Return Reason
    '    End If
    '    'sm.UserInfo.TPlanID != "28" 一般計劃專用
    '    If flagTIMSNot28a Then
    '        Reason=CheckImportDataTIMS(colArray)
    '        Return Reason
    '    End If

    '    Return Reason
    'End Function

    '查詢按鈕 
    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    'If ViewState("LastOCIDValue1")="" Then
    '    '    Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
    '    '    Exit Sub
    '    'End If
    '    'If ViewState("LastOCIDValue1") <> OCIDValue1.Value Then
    '    '    Common.MessageBox(Me, "班級選擇有誤，請重新選擇")
    '    '    Exit Sub
    '    'End If
    '    Call Search1()  '查詢按鈕 SQL
    'End Sub

#End Region

End Class
