Partial Class TC_01_004_add
    Inherits AuthBasePage

    '(TIMS)
    'INSERT/UPDATE/SAVE CLASS_CLASSINFO

    Const cst_TooltipT1 As String = "已執行計畫轉入修改時,不能修改!!"
    Dim flag_TPlanID70_1 As Boolean = False ' '70:區域產業據點職業訓練計畫(在職)

    '報名登錄最晚可作業時間 (FEnterDate2)
    Sub SUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("'ORG_ORGINFO','ID_CLASS','CLASS_CLASSINFO','KEY_TRAINTYPE'", objconn)
        If dt.Rows.Count = 0 Then Return 'Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "ORGNAME", TBplan) '計畫階層
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSID", TBclass_id) '班別代碼
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSCNAME", TB_ClassName) '班級中文名稱
        Call TIMS.sUtl_SetMaxLen(dt, "CYCLTYPE", TB_CyclType) '期別
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSENGNAME", ClassEngName) '班級英文名稱
        Call TIMS.sUtl_SetMaxLen(dt, "TRAINNAME", TB_career_id) '訓練職類
        Call TIMS.sUtl_SetMaxLen(dt, "TADDRESS", TBaddress) '上課地點
        Call TIMS.sUtl_SetMaxLen(dt, "EADDRESS", EADDRESS) '甄試地點
        Call TIMS.sUtl_SetMaxLen(dt, "CTNAME", CTName) '導師名稱
        Call TIMS.sUtl_SetMaxLen(dt, "OTHERREASON", OtherReason) '其他原因說明

        '70:區域產業據點職業訓練計畫(在職)
        flag_TPlanID70_1 = (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1)
        'OJT-23041104：區域據點-開班資料查詢：【班級英文名稱】改為非必填 
        td_ClassEngName.Attributes.Add("class", If(flag_TPlanID70_1, "bluecol", "bluecol_need"))
        'td_ExamDate'ExamDate'ImgExamDate'ExamPeriod'lab_msg_ExamDate
        'trEADDRESS         CheckBox1
        'EZip_Code'EADDRESSZIPB3'hidEADDRESSZIP6W'LitEZipCode'ECity'Ezipbtn'EADDRESS
        If flag_TPlanID70_1 Then
            'reqFVcity2 'reqFVaddress2
            reqFVcity2.Enabled = False '不啟用(甄試地點)
            reqFVaddress2.Enabled = False '不啟用(甄試地點)
            TIMS.SET_CLASS_BLUECOL_1(td_ExamDate, lab_msg_ExamDate)
            TIMS.SET_CLASS_BLUECOL_1(td_EADDRESS, lab_msg_EADDRESS) '(甄試地點)
        End If
    End Sub

    '報名登錄最晚可作業時間 (FEnterDate2)
    'Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
    '    Call sUtl_PageInit1()
    'End Sub

    'Dim dtSHAREJOB As DataTable 'SHARE_CJOB
    'Dim temp_table As DataTable
    'Dim temp_dr As DataRow
    'FROM CLASS_CLASSINFO
    'FROM PLAN_PLANINFO
    'Dim dtlevel As DataTable
    'Dim dr As DataRow

    Dim ff3 As String = ""
    Dim ProcessType As String 'PlanUpdate 'Update'Insert
    Const cst_PlanUpdate As String = "PlanUpdate" '計畫轉入
    Const cst_Update As String = "Update" '修改
    Const cst_Insert As String = "Insert" '新增 (新增開班資料)

    '(TIMS專用非產投) 'TC_01_004_InsertPlan.aspx 
    '產投專用。'TC_01_004_BusAdd.aspx
    'Response.Redirect("TC_01_004_BusAdd.aspx?ID=" & Request("ID") & "&STDate=" & vsSTDate)
    'strScript1 &= "location.href='TC_01_004_add.aspx?ProcessType=PlanUpdate&ID='+document.getElementById('Re_ID').value;" + vbCrLf
    Const cst_temp_classinfo As String = "temp_classinfo" 'Session(cst_temp_classinfo) TC_01_004_InsertPlan.aspx
    Const cst_ClassSearchStr As String = "ClassSearchStr"

    Dim iOCID_New As Integer = 0
    Dim rqOCID As String = ""
    'Dim gflag_test As Boolean=False '測試

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線
        Call SUtl_PageInit1()
        '檢查Session是否存在 End 'gflag_test=TIMS.sUtl_ChkTest() '測試

        '甄試日期/時間 'If TIMS.GFG_OJT_25050801_NoUse_ExamDateTime Then spExamDateTime.Visible=False:HR6.Visible=False : MM6.Visible=False

        ProcessType = TIMS.ClearSQM(Request("ProcessType"))
        rqOCID = TIMS.ClearSQM(Request("ocid")) '小寫
        If rqOCID = "" Then rqOCID = TIMS.ClearSQM(Request("OCID")) '試試看大寫

        '2005/3/7新增輸入郵遞區號回傳地區名稱的Javascript-Start
        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))
        ' TBCity.Attributes("onblur")="getzipname(this.value,'TBCity','city_code');"
        'Label1.Attributes("onclick")="javascript:return words();"
        'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
        '課程階段
        'tb_CLASSLEVEL.Visible=False '課程階段
        'add_but.Attributes("onclick")="javascript:return check_add();" LEVEL
        CheckBox1.Attributes.Add("onClick", "return ock_CheckBox1();")

        'bt_save.Attributes("onclick")="return CheckData();"
        If ProcessType = cst_PlanUpdate Then TPeriodList.Attributes("onchange") = "check_value();"

        '自辦在職顯示此功能 且為必填
        'LExamDate1.Visible=True : LExamDate2.Visible=False '甄試日期 自辦在職 為必填
        '訓練時段
        trTB_NOTE3.Visible = If(TIMS.Cst_TPlanID06AppPlan1.IndexOf(sm.UserInfo.TPlanID) > -1, True, False) 'True :顯示/False :不顯示

        'Litcity_code.Text=TIMS.Get_WorkZIPB3Link2()
        'LitEZipCode.Text=TIMS.Get_WorkZIPB3Link2()

        CompanyTR.Visible = False '企業名稱隱藏
        'If RB_TPropertyID.SelectedIndex=2 Then CompanyTR.Visible=True '企業名稱顯示

        If Not Page.IsPostBack Then
            Call Create1()
        Else
            '假如不開訓原因不能修改時
            '必須抓回數值
            If Not NORID.Enabled Then
                For i As Integer = 0 To Split(NORIDValue.Value, ",").Length - 1
                    For j As Integer = 0 To NORID.Items.Count - 1
                        If Split(NORIDValue.Value, ",")(i) = NORID.Items(j).Value Then NORID.Items(j).Selected = True
                    Next
                Next
            End If
        End If
    End Sub

    Sub Create1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Return 'Exit Sub
        If Session(cst_ClassSearchStr) IsNot Nothing Then
            Session(cst_ClassSearchStr) = Session(cst_ClassSearchStr)
            ViewState(cst_ClassSearchStr) = Session(cst_ClassSearchStr)
        End If

        '(訓練機構屬性設定)-郵遞區號查詢
        Litcity_code.Text = TIMS.Get_WorkZIPB3Link2()
        LitEZipCode.Text = TIMS.Get_WorkZIPB3Link2()

        Dim bt1_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code, TaddressZIPB3, hidTaddressZIP6W, TBCity, TBaddress)
        Bt1_city_zip.Attributes.Add("onclick", bt1_Attr_VAL)
        Dim bt2_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(EZip_Code, EADDRESSZIPB3, hidEADDRESSZIP6W, ECity, EADDRESS)
        Ezipbtn.Attributes.Add("onclick", bt2_Attr_VAL)

        'If Not Session("ClassSearchStr") Is Nothing Then
        '    ViewState("ClassSearchStr")=Session("ClassSearchStr")
        '    Session("ClassSearchStr")=Nothing
        'End If

        '接受企業委託
        'If RB_TPropertyID.Items.FindByValue("2") Is Nothing Then
        '    Dim sPlanKind As String=TIMS.Get_PlanKind(Me, objconn)
        '    Select Case sPlanKind
        '        Case "1"
        '            'RB_TPropertyID.Items.Insert(2, New ListItem("接受企業委託", "2"))
        '            RB_TPropertyID.Items.Insert(1, New ListItem("接受企業委託", "2"))        '(20181003 由於承辦人要求將"職前"的選項拿掉,所以後續的選項都需往前遞補)
        '        Case Else
        '            If sm.UserInfo.TPlanID="07" Then  '自辦計畫與接受企業委託計畫
        '                'RB_TPropertyID.Items.Insert(2, New ListItem("接受企業委託", "2"))
        '                RB_TPropertyID.Items.Insert(1, New ListItem("接受企業委託", "2"))    '(20181003 由於承辦人要求將"職前"的選項拿掉,所以後續的選項都需往前遞補)
        '            End If
        '    End Select
        'End If

        'Dim PlanKind As String
        'Dim sql As String="SELECT PLANKIND FROM ID_PLAN WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        'PlanKind=DbAccess.ExecuteScalar(sql, objconn)
        'If PlanKind="1" OrElse sm.UserInfo.TPlanID="07" Then RB_TPropertyID.Items.Insert(2, New ListItem("接受企業委託", "2"))  '自辦計畫與接受企業委託計畫
        CompanyTR.Visible = False '企業名稱隱藏

        '2005/8/12新增不開班選項只有系統管理者(1)and承辦人(5)才可修改--Melody
        If sm.UserInfo.RoleID <= 5 Then
            CB_NotOpen.Enabled = True
            NORID.Enabled = True
            OtherReason.Enabled = True
        Else
            '假如登入單位為管控單位,也可以修改 - Start
            Dim oIsConUnit As Object = 0 ' Boolean
            Dim sPMS As New Hashtable From {{"OrgID", sm.UserInfo.OrgID}}
            Dim sql As String = "SELECT IsConUnit FROM ORG_ORGINFO WHERE OrgID =@OrgID"
            oIsConUnit = DbAccess.ExecuteScalar(sql, objconn, sPMS)
            Dim fg_IsConUnit As Boolean = (oIsConUnit IsNot Nothing AndAlso Val(oIsConUnit) = 1)
            CB_NotOpen.Enabled = If(fg_IsConUnit, True, False)
            NORID.Enabled = If(fg_IsConUnit, True, False)
            OtherReason.Enabled = If(fg_IsConUnit, True, False)
            '假如登入單位為管控單位,也可以修改 - End
        End If

        'Session("ClassLevel")=Nothing '課程階段
        If ProcessType = cst_PlanUpdate Then Button1.Enabled = False '回上一頁

        '20100329 andy add  甄試時段
        '取出鍵詞-甄試時段鍵詞檔
        ExamPeriod = TIMS.GET_ExamPeriod(ExamPeriod, objconn)

        Select Case ProcessType
            Case cst_Insert
                lblProecessType.Text = "新增"
                'ExamPeriod.SelectedValue="01"
                'Common.SetListItem(ExamPeriod, "01")

            Case cst_Update, cst_PlanUpdate '修改 '轉入
                'PlanUpdate (班級轉入)
                lblProecessType.Text = "修改"
                '2005/6/20--Melody 計畫轉入and修改時,班別名稱,期別,開結訓日不能修改
                TB_ClassName.Enabled = False
                TB_CyclType.Enabled = False
                TB_STDate.Enabled = False
                TB_FTDate.Enabled = False
                TIMS.Tooltip(TB_ClassName, cst_TooltipT1)
                TIMS.Tooltip(TB_CyclType, cst_TooltipT1)
                TIMS.Tooltip(TB_STDate, cst_TooltipT1)
                TIMS.Tooltip(TB_FTDate, cst_TooltipT1)

                'OJT-21080601：自辦在職、接受企業委託、區域據點- 開班資料查詢：將「訓練人數」、「上課地址」欄位反灰不可修改
                TB_TNum.Enabled = False
                TIMS.Tooltip(TB_TNum, cst_TooltipT1)
                city_code.Disabled = True
                TaddressZIPB3.Disabled = True
                TBCity.Enabled = False
                Bt1_city_zip.Disabled = True
                TIMS.Tooltip(Bt1_city_zip, cst_TooltipT1)
                TBaddress.Enabled = False
                TIMS.Tooltip(city_code, cst_TooltipT1)
                TIMS.Tooltip(TaddressZIPB3, cst_TooltipT1)
                TIMS.Tooltip(TBCity, cst_TooltipT1)
                TIMS.Tooltip(TBaddress, cst_TooltipT1)

                '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                If sm.UserInfo.LID = "2" Then '委訓單位無法變更
                    TB_CheckInDate.Enabled = False
                    TIMS.Tooltip(TB_CheckInDate, cst_TooltipT1)
                End If

                'txtCJOB_NAME.ReadOnly=True '通俗職類
                'btu_sel2.Disabled=True '通俗職類(重新選擇開放) BY AMU 201107
                'check_Tplan="SELECT TPlanID from ID_Plan where PlanID='" & sm.UserInfo.PlanID & "'"
                'TPlan_str=Convert.ToString(DbAccess.ExecuteScalar(check_Tplan, objconn))

                'Select Case Convert.ToString(sm.UserInfo.TPlanID) 'TPlan_str
                '    Case "15" '學習券
                '        class_unit_button.Visible=True
                '        TB_ClassName.ReadOnly=True
                '        '2005/6/20--Melody 計畫轉入and修改時,班別名稱不能修改
                '        TB_ClassName.Enabled=True
                '        '不能選取"納入志願"選項
                '        CB_IsApplic.Enabled=False
                '        tb_TPlan_str.Value="15"
                '    Case Else
                '        class_unit_button.Visible=False
                '        TB_ClassName.ReadOnly=False
                '        CB_IsApplic.Enabled=True
                '        tb_TPlan_str.Value=""
                'End Select

                '(班別名稱不能修改)
                class_unit_button.Visible = False
                TB_ClassName.ReadOnly = False
                '選取"納入志願"選項
                CB_IsApplic.Enabled = True
                tb_TPlan_str.Value = ""

                For i As Integer = 0 To 23
                    Dim s_i As String = TIMS.AddZero(CStr(i), 2)
                    HR1.Items.Add(New ListItem(s_i, s_i))
                    HR2.Items.Add(New ListItem(s_i, s_i))
                    'HR3.Items.Add(New ListItem(s_i, s_i))
                    'HR4.Items.Add(New ListItem(s_i, s_i))
                    HR5.Items.Add(New ListItem(s_i, s_i))
                    HR6.Items.Add(New ListItem(s_i, s_i))
                Next

                For j As Integer = 0 To 59
                    Dim s_j As String = TIMS.AddZero(CStr(j), 2)
                    MM1.Items.Add(New ListItem(s_j, s_j))
                    MM2.Items.Add(New ListItem(s_j, s_j))
                    'MM3.Items.Add(New ListItem(s_j, s_j))
                    'MM4.Items.Add(New ListItem(s_j, s_j))
                    MM5.Items.Add(New ListItem(s_j, s_j))
                    MM6.Items.Add(New ListItem(s_j, s_j))
                Next
                Common.SetListItem(HR2, 23) : Common.SetListItem(MM2, 59)
                'Common.SetListItem(HR4, 23)'Common.SetListItem(MM4, 59)
                Common.SetListItem(HR5, 23) : Common.SetListItem(MM5, 59)
        End Select

        'If ProcessType=cst_Insert Then
        'ElseIf ProcessType=cst_Update OrElse ProcessType=cst_PlanUpdate Then
        'End If

        '課程階段
        'TB_LevelType.Items.Clear() '原資料 不使用清除功能
        'TB_LevelType.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        '課程階段 階段
        'LevelName.Items.Clear()
        'LevelName.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        '取出鍵詞-不開班理由代碼
        Call TIMS.Get_NotOpenReason(NORID, objconn)
        '取得鍵值-訓練時段
        TPeriodList = TIMS.GET_HOURRAN(TPeriodList, objconn, sm)
        '取出鍵詞-訓練期限代碼
        Call TIMS.GET_TRAINEXP(TDeadline_List, objconn, sm)

        trEADDRESS.Visible = True '其他計畫(非產投)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then trEADDRESS.Visible = False '不顯示甄試地點。'產投執行

        next_class.Enabled = False
        If ProcessType = cst_PlanUpdate Then
            next_class.Enabled = True '計畫轉入後帶到此頁時,"維護下一班"才可使用 TC_01_004_InsertPlan.aspx
            If Session(cst_temp_classinfo) Is Nothing Then Return 'Exit Sub
            Dim temp_table As DataTable = Session(cst_temp_classinfo)
            If temp_table Is Nothing OrElse temp_table.Rows.Count = 0 Then Return 'Exit Sub 'Session(cst_temp_classinfo)=Nothing
            'SHOW_TMP_CLASSINFO
            Call SHOW_TMP_CLASSINFO(temp_table) 'Session(cst_temp_classinfo)

        ElseIf ProcessType = cst_Update Then 'Or ProcessType=cst_PlanUpdate 
            '修改 'Call SHOW_CLASSLEVEL() '課程階段
            Call SHOW_CLASSINFO()

        End If
        Call SHOW_PLANINFO()
    End Sub

    '取得鍵值
    'Function Get_CLASSID(ByVal sCLSID As String) As String
    '    Dim rst As String=""
    '    If sCLSID="" Then Return rst
    '    Dim sql As String="SELECT CLASSID FROM ID_CLASS WHERE CLSID=@CLSID "
    '    Dim sCmd As New SqlCommand(sql, objconn)
    '    Call TIMS.OpenDbConn(objconn)
    '    With sCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("CLSID", SqlDbType.VarChar).Value=sCLSID
    '        rst=Convert.ToString(.ExecuteScalar())
    '    End With
    '    Return rst
    '    'classid=DbAccess.ExecuteScalar(className, objconn)
    'End Function

    '取得鍵值
    Function Get_TRAINNAME(ByVal oTMID As Object) As String
        Dim rst As String = ""
        Dim sTMID As String = Convert.ToString(oTMID)
        If sTMID = "" Then Return rst
        Dim sql As String = "SELECT TRAINNAME FROM KEY_TRAINTYPE WHERE TMID=@TMID "
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("TMID", SqlDbType.VarChar).Value = sTMID
            rst = Convert.ToString(.ExecuteScalar())
        End With
        Return rst
    End Function

    '取得鍵值
    Function Get_TB_CLASSID(ByVal oCLSID As Object) As String
        Dim rst As String = ""
        Dim sCLSID As String = Convert.ToString(oCLSID)
        If sCLSID = "" Then Return rst
        Dim sql As String = "SELECT CLASSID FROM ID_CLASS WHERE CLSID=@CLSID "
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("CLSID", SqlDbType.VarChar).Value = sCLSID
            rst = Convert.ToString(.ExecuteScalar())
        End With
        Return rst
    End Function

    '(轉班使用) 'SHOW_TMP_CLASSINFO  'Session(cst_temp_classinfo)  TC_01_004_InsertPlan.aspx
    Sub SHOW_TMP_CLASSINFO(ByVal temp_table As DataTable)
        Dim sql As String = ""
        For Each temp_dr As DataRow In temp_table.Rows
            'sql="SELECT CLASSID FROM ID_CLASS WHERE CLSID='" & temp_dr("CLSID") & "'"
            'Dim sCLASSID As String=DbAccess.ExecuteScalar(sql, objconn)
            PlanIDValue.Value = temp_dr("PlanID").ToString()
            P_ComIDNO.Value = temp_dr("ComIDNO").ToString()
            P_SeqNO.Value = temp_dr("SeqNO").ToString()
            P_Relship.Value = temp_dr("Relship").ToString()
            P_Years.Value = temp_dr("Years").ToString()
            clsid.Value = temp_dr("CLSID").ToString()

            'Dim classid As String
            TBclass_id.Text = Get_TB_CLASSID(temp_dr("CLSID")) 'sCLASSID
            Dim v_CyclType As String = TIMS.ClearSQM(temp_dr("CyclType"))
            If v_CyclType = "" Then v_CyclType = TIMS.cst_Default_CyclType
            TB_CyclType.Text = TIMS.FmtCyclType(v_CyclType)

            TB_ClassName.Text = Convert.ToString(temp_dr("ClassCName"))
            ClassEngName.Text = Convert.ToString(temp_dr("ClassEngName"))
            TB_Content.Text = Convert.ToString(temp_dr("Content"))
            RIDValue.Value = Convert.ToString(temp_dr("RID"))
            Hid_RID1.Value = Convert.ToString(temp_dr("RID")).Substring(0, 1)
            hid_classnum.Value = Convert.ToString(temp_dr("ClassNum"))
            Common.SetListItem(rblADVANCE, temp_dr("ADVANCE")) '訓練課程類型 ADVANCE
            TB_TNum.Text = Convert.ToString(temp_dr("TNum"))

            '2005/8/15-加上帶入訓練地點&目標 add by Jack
            TB_Purpose.Text = Convert.ToString(temp_dr("Purpose")) '課程目標
            Companyname.Text = Convert.ToString(temp_dr("Companyname")) '企業名稱
            '班級英文名稱
            ClassEngName.Text = Convert.ToString(temp_dr("ClassEngName"))
            '訓練時段'取得鍵值-訓練時段 'TPeriodList
            Common.SetListItem(TPeriodList, Convert.ToString(temp_dr("TPeriod")))
            TB_NOTE3.Text = Convert.ToString(temp_dr("NOTE3"))
            '「訓練期限」
            Common.SetListItem(TDeadline_List, Convert.ToString(temp_dr("TDeadline")))
            '導師名稱
            CTName.Text = Convert.ToString(temp_dr("CTName"))

            city_code.Value = Convert.ToString(temp_dr("TaddressZip")) 'TIMS.AddZero(Convert.ToString(temp_dr("TaddressZip")), 3)
            hidTaddressZIP6W.Value = Convert.ToString(temp_dr("TaddressZIP6W"))
            TaddressZIPB3.Value = TIMS.GetZIPCODEB3(hidTaddressZIP6W.Value)
            TBCity.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(temp_dr("TaddressZip")), Convert.ToString(temp_dr("TaddressZIP6W")))
            TBaddress.Text = Convert.ToString(temp_dr("TAddress"))

            EZip_Code.Value = TIMS.AddZero(Convert.ToString(temp_dr("EADDRESSZIP")), 3)
            hidEADDRESSZIP6W.Value = Convert.ToString(temp_dr("EADDRESSZIP6W"))
            EADDRESSZIPB3.Value = TIMS.GetZIPCODEB3(hidEADDRESSZIP6W.Value)
            ECity.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(temp_dr("EADDRESSZIP")), Convert.ToString(temp_dr("EADDRESSZIP6W")))
            EADDRESS.Text = Convert.ToString(temp_dr("EADDRESS"))

            hidEZip_Code.Value = EZip_Code.Value
            hidEADDRESSZIPB3.Value = EADDRESSZIPB3.Value
            hidhidEADDRESSZIP6W.Value = hidEADDRESSZIP6W.Value
            hidECity.Value = ECity.Text
            hidEADDRESS.Value = EADDRESS.Text

            '若沒有 甄試地點 資料，直接填入 上課資料
            CheckBox1.Checked = False

            LabAdd.Text = TIMS.Get_LabAddforRELSHIP(P_Relship.Value, objconn)
            'End of add
            'Dim tdid As String=temp_dr("TMID")
            trainValue.Value = Convert.ToString(temp_dr("TMID"))
            trainValue.Value = TIMS.ClearSQM(trainValue.Value)
            cjobValue.Value = Convert.ToString(temp_dr("CJOB_UNKEY")) '通俗職類

            TB_career_id.Text = ""
            If trainValue.Value <> "" Then
                Dim sPMS As New Hashtable From {{"TMID", trainValue.Value}}
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    sql = "SELECT dbo.NVL(JOBNAME,TRAINNAME) TRAINNAME FROM KEY_TRAINTYPE WHERE TMID=@TMID"
                Else
                    sql = "SELECT TRAINNAME FROM KEY_TRAINTYPE WHERE TMID=@TMID"
                End If
                Dim oTRAINNAME As Object = DbAccess.ExecuteScalar(sql, objconn, sPMS)
                TB_career_id.Text = If(oTRAINNAME IsNot Nothing, Convert.ToString(oTRAINNAME), "")
            End If

            '通俗職類 SHARE_CJOB
            Dim dr99 As DataRow = TIMS.Get_SHARECJOB(Convert.ToString(temp_dr("CJOB_UNKEY")), objconn)
            txtCJOB_NAME.Text = ""
            cjobValue.Value = ""
            If Not dr99 Is Nothing Then
                txtCJOB_NAME.Text = dr99("CJOB_NAME")
                cjobValue.Value = dr99("CJOB_UNKEY")
            End If

            TB_THours.Text = Convert.ToString(temp_dr("THours"))
            TB_STDate.Text = TIMS.Cdate3(temp_dr("STDate"))
            TB_FTDate.Text = TIMS.Cdate3(temp_dr("FTDate"))

            '起始日
            TB_SEnterDate.Text = ""
            If Convert.ToString(temp_dr("SEnterDate")) <> "" Then
                TB_SEnterDate.Text = TIMS.Cdate3(temp_dr("SEnterDate"))
                Common.SetListItem(HR1, TIMS.AddZero(CDate(temp_dr("SEnterDate")).Hour, 2))
                Common.SetListItem(MM1, TIMS.AddZero(CDate(temp_dr("SEnterDate")).Minute, 2))
                'If DateDiff(DateInterval.Day, CDate(temp_dr("SEnterDate")), CDate(Now())) >= 0 Then EnteredFlag=True
            End If

            '結束日
            TB_FEnterDate.Text = ""
            If Convert.ToString(temp_dr("FEnterDate")) <> "" Then
                TB_FEnterDate.Text = TIMS.Cdate3(temp_dr("FEnterDate"))
                Common.SetListItem(HR2, TIMS.AddZero(CDate(temp_dr("FEnterDate")).Hour, 2))
                Common.SetListItem(MM2, TIMS.AddZero(CDate(temp_dr("FEnterDate")).Minute, 2))
            End If
            '甄試日期2005/3/23 CLASS_CLASSINFO
            'If (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1) AndAlso v_GetTrain3="" Then
            '    ExamDate.Text="" 'TIMS.cdate3(dr("ExamDate")) 'ExamPeriod.SelectedIndex=-1
            '    Common.SetListItem(ExamPeriod, "")
            'ElseIf Convert.ToString(temp_dr("ExamDate")) <> "" Then
            '    ExamDate.Text=TIMS.cdate3(temp_dr("ExamDate"))  '20100329 andy add  甄試時段 '01-全天'02-上午'03-下午
            '    If Convert.ToString(temp_dr("ExamPeriod")) <> "" Then Common.SetListItem(ExamPeriod, temp_dr("ExamPeriod")) '甄試時段
            'Else
            '    ExamDate.Text="" 'TIMS.cdate3(dr("ExamDate")) 'ExamPeriod.SelectedIndex=-1
            '    Common.SetListItem(ExamPeriod, "")
            'End If

            If Convert.ToString(temp_dr("ExamDate")) <> "" Then
                'ExamDate.Text=Common.FormatDate(temp_dr("ExamDate"))  '20100329 andy add  甄試時段 '01-全天'02-上午'03-下午
                ExamDate.Text = TIMS.Cdate3(temp_dr("ExamDate"))
                TIMS.SET_DateHM(CDate(temp_dr("ExamDate")), HR6, MM6)
                If Convert.ToString(temp_dr("ExamPeriod")) <> "" Then Common.SetListItem(ExamPeriod, temp_dr("ExamPeriod"))
            Else
                ExamDate.Text = "" ' ExamPeriod.SelectedIndex=-1
                TIMS.SET_DateHM((Now.ToString("yyyy/MM/dd") & " 00:00"), HR6, MM6)
                Common.SetListItem(ExamPeriod, "")
            End If

            'FEnterDate2  試著計算
            If Convert.ToString(temp_dr("FEnterDate2")) = "" Then
                Dim sFENTERDATE As String = TB_FEnterDate.Text
                Dim sEXAMDATE As String = ExamDate.Text
                Dim SS1 As String = ""
                TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
                Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)
                'Dim sFENTERDATE2 As String=TIMS.GET_FENTERDATE2(Hid_RID1.Value, sFENTERDATE, sEXAMDATE, objconn)
                If sFENTERDATE2 <> "" Then
                    FEnterDate2.Text = TIMS.Cdate3(sFENTERDATE2)
                    Common.SetListItem(HR5, TIMS.AddZero(CDate(sFENTERDATE2).Hour, 2))
                    Common.SetListItem(MM5, TIMS.AddZero(CDate(sFENTERDATE2).Minute, 2))
                End If
            Else
                '不為空值
                FEnterDate2.Text = TIMS.Cdate3(temp_dr("FEnterDate2"))
                Common.SetListItem(HR5, TIMS.AddZero(CDate(temp_dr("FEnterDate2")).Hour, 2))
                Common.SetListItem(MM5, TIMS.AddZero(CDate(temp_dr("FEnterDate2")).Minute, 2))
            End If

            'FEnterDate2.Enabled=False
            'HR5.Enabled=False
            'MM5.Enabled=False
            'Dim sTip2 As String="請輸入「報名結束日期」及「甄試日期」計算後帶出「報名登錄最晚可作業日期」"
            Dim sTip2 As String = "請輸入「報名登錄最晚可作業日期」"
            TIMS.Tooltip(FEnterDate2, sTip2)
            TIMS.Tooltip(MM5, sTip2)
            TIMS.Tooltip(MM5, sTip2)

            '報到日期
            If Convert.ToString(temp_dr("CheckInDate")) <> "" Then TB_CheckInDate.Text = TIMS.Cdate3(temp_dr("CheckInDate"))

            If TB_SEnterDate.Text <> "" AndAlso TB_FEnterDate.Text <> "" AndAlso ExamDate.Text <> "" Then
                Dim flag_lock1 As Boolean = True
                '排除產投計畫
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_lock1 = False
                '排除自辦在職計畫
                If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_lock1 = False

                If flag_lock1 Then
                    TB_SEnterDate.Enabled = False
                    HR1.Enabled = False
                    MM1.Enabled = False
                    date1.Visible = False
                    TB_FEnterDate.Enabled = False
                    HR2.Enabled = False
                    MM2.Enabled = False
                    date2.Visible = False
                    ExamDate.Enabled = False
                    ImgExamDate.Visible = False
                    ExamPeriod.Enabled = False
                    HR6.Enabled = False
                    MM6.Enabled = False
                    Dim sTip As String = "若要修改,請依規定進行班級變更申請,並告知承辦人."
                    TIMS.Tooltip(TB_SEnterDate, sTip)
                    TIMS.Tooltip(HR1, sTip)
                    TIMS.Tooltip(MM1, sTip)
                    TIMS.Tooltip(TB_FEnterDate, sTip)
                    TIMS.Tooltip(HR2, sTip)
                    TIMS.Tooltip(MM2, sTip)
                    TIMS.Tooltip(ExamDate, sTip)
                    TIMS.Tooltip(ExamPeriod, sTip)
                    TIMS.Tooltip(HR6, sTip)
                    TIMS.Tooltip(MM6, sTip)
                End If
            End If

            CB_IsApplic.Checked = If(Convert.ToString(temp_dr("IsApplic")) = "Y", True, False)
            CB_NotOpen.Checked = If(Convert.ToString(temp_dr("NotOpen")) = "Y", True, False)

            Call TIMS.SetCblValue(NORID, Convert.ToString(temp_dr("NORID")))
            NORIDValue.Value = TIMS.GetCblValue(NORID)

            OtherReason.Text = temp_dr("OtherReason").ToString
            '2005/5/27班級單元for 學習卷-Melody

            'Select Case Convert.ToString(sm.UserInfo.TPlanID)
            '    Case "15" '學習券
            '        If Convert.IsDBNull(temp_dr("Class_Unit")) Then
            '            tb_class_unit.Value=""
            '        Else
            '            tb_class_unit.Value=temp_dr("Class_Unit")
            '        End If
            'End Select

            TBplan.Text = TIMS.Get_OrgNameInputRID(CStr(temp_dr("RID")), objconn)
        Next

#Region "(No Use)"

        'Dim plan_name As String
        'Dim sqlstr_PlanName As String="SELECT ORGNAME  FROM AUTH_RELSHIP a join ORG_ORGINFO b on a.orgid =b.orgid where a.rid='" & temp_dr("RID") & "'"
        'plan_name=Convert.ToString(DbAccess.ExecuteScalar(sqlstr_PlanName, objconn))
        'Me.TBplan.Text=plan_name
        'check_Tplan="SELECT TPLANID FROM ID_PLAN WHERE PlanID='" & PlanIDValue.Value & "'"
        'TPlan_str=DbAccess.ExecuteScalar(check_Tplan, objconn)
        'If TPlan_str="15" Then '學習券
        '    class_unit_button.Visible=True
        'Else
        '    class_unit_button.Visible=False
        'End If

#End Region

        class_unit_button.Visible = False
        'Select Case Convert.ToString(sm.UserInfo.TPlanID)
        '    Case "15" '學習券
        '        class_unit_button.Visible=True
        'End Select
    End Sub

    'SHOW CLASS_CLASSINFO
    Sub SHOW_CLASSINFO()
        'CLASS_CLASSINFO

        Dim parms As New Hashtable From {{"OCID", rqOCID}}
        Dim sqlstr_list As String = "SELECT * FROM CLASS_CLASSINFO WHERE OCID=@OCID" '& rqOCID
        Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr_list, objconn, parms)
        If row_list Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return 'Exit Sub
        End If

        PlanIDValue.Value = row_list("PlanID").ToString()
        P_ComIDNO.Value = row_list("ComIDNO").ToString()
        P_SeqNO.Value = row_list("SeqNO").ToString()
        P_Relship.Value = row_list("Relship").ToString()
        P_Years.Value = row_list("Years").ToString()
        clsid.Value = row_list("CLSID").ToString()

        hid_classnum.Value = Convert.ToString(row_list("ClassNum"))
        '2005/5/27班級單元for 學習卷-Melody
        tb_class_unit.Value = Convert.ToString(row_list("Class_Unit"))

        'planid=Convert.ToString(row_list("PlanID"))
        Hid_RID1.Value = Convert.ToString(row_list("RID")).Substring(0, 1)

        TBclass_id.Text = Get_TB_CLASSID(row_list("CLSID"))
        OldClassID.Value = TBclass_id.Text
        TB_ClassName.Text = Convert.ToString(row_list("ClassCName"))
        TB_CyclType.Text = TIMS.FmtCyclType(row_list("CyclType"))
        ClassEngName.Text = Convert.ToString(row_list("ClassEngName")) '班級英文名稱
        Companyname.Text = Convert.ToString(row_list("Companyname")) '企業名稱

        TB_career_id.Text = "" '
        trainValue.Value = "" '
        If Convert.ToString(row_list("TMID")) <> "" Then
            TB_career_id.Text = Get_TRAINNAME(Convert.ToString(row_list("TMID")))
            trainValue.Value = Convert.ToString(row_list("TMID"))
        End If

        '通俗職類 SHARE_CJOB
        Dim dr99 As DataRow = TIMS.Get_SHARECJOB(Convert.ToString(row_list("CJOB_UNKEY")), objconn)
        txtCJOB_NAME.Text = ""
        cjobValue.Value = ""
        If dr99 IsNot Nothing Then
            txtCJOB_NAME.Text = Convert.ToString(dr99("CJOB_NAME"))
            cjobValue.Value = Convert.ToString(dr99("CJOB_UNKEY"))
        End If

        Dim EnteredFlag As Boolean = False '是否報名時間已過
        '起始日
        TB_SEnterDate.Text = ""
        If Convert.ToString(row_list("SEnterDate")) <> "" Then
            TB_SEnterDate.Text = TIMS.Cdate3(row_list("SEnterDate"))
            Common.SetListItem(HR1, TIMS.AddZero(CDate(row_list("SEnterDate")).Hour, 2))
            Common.SetListItem(MM1, TIMS.AddZero(CDate(row_list("SEnterDate")).Minute, 2))
            If DateDiff(DateInterval.Day, CDate(row_list("SEnterDate")), CDate(Now())) >= 0 Then EnteredFlag = True
        End If
        '結束日
        TB_FEnterDate.Text = ""
        If Convert.ToString(row_list("FEnterDate")) <> "" Then
            TB_FEnterDate.Text = TIMS.Cdate3(row_list("FEnterDate"))
            Common.SetListItem(HR2, TIMS.AddZero(CDate(row_list("FEnterDate")).Hour, 2))
            Common.SetListItem(MM2, TIMS.AddZero(CDate(row_list("FEnterDate")).Minute, 2))
        End If
        '甄試日期2005/3/23
        If Convert.ToString(row_list("ExamDate")) <> "" Then
            'ExamDate.Text=Common.FormatDate(row_list("ExamDate"))  '20100329 andy add 甄試時段
            ExamDate.Text = TIMS.Cdate3(row_list("ExamDate"))
            Common.SetListItem(HR6, TIMS.AddZero(CDate(row_list("ExamDate")).Hour, 2))
            Common.SetListItem(MM6, TIMS.AddZero(CDate(row_list("ExamDate")).Minute, 2))
            'If Convert.ToString(row_list("ExamPeriod")) <> "" Then Common.SetListItem(ExamPeriod, Convert.ToString(row_list("ExamPeriod")))
            Common.SetListItem(ExamPeriod, Convert.ToString(row_list("ExamPeriod")))
        Else
            ExamDate.Text = ""
            Common.SetListItem(ExamPeriod, "") 'ExamPeriod.SelectedIndex=-1
            Common.SetListItem(HR6, "00")
            Common.SetListItem(MM6, "00")
        End If

        'FEnterDate2  試著計算
        If Convert.ToString(row_list("FEnterDate2")) = "" Then
            Dim sFENTERDATE As String = TB_FEnterDate.Text
            Dim sEXAMDATE As String = ExamDate.Text
            Dim SS1 As String = ""
            TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
            Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)
            'Dim sFENTERDATE2 As String=TIMS.GET_FENTERDATE2(Hid_RID1.Value, sFENTERDATE, sEXAMDATE, objconn)
            If sFENTERDATE2 <> "" Then
                FEnterDate2.Text = TIMS.Cdate3(sFENTERDATE2)
                Common.SetListItem(HR5, TIMS.AddZero(CDate(sFENTERDATE2).Hour, 2))
                Common.SetListItem(MM5, TIMS.AddZero(CDate(sFENTERDATE2).Minute, 2))
            End If
        Else
            '不為空值
            FEnterDate2.Text = TIMS.Cdate3(row_list("FEnterDate2"))
            Common.SetListItem(HR5, TIMS.AddZero(CDate(row_list("FEnterDate2")).Hour, 2))
            Common.SetListItem(MM5, TIMS.AddZero(CDate(row_list("FEnterDate2")).Minute, 2))
        End If

        'FEnterDate2  試著計算
        'If Convert.ToString(row_list("FEnterDate2"))="" Then
        '    Dim sFENTERDATE As String=TB_FEnterDate.Text
        '    Dim sEXAMDATE As String=ExamDate.Text
        '    Dim SS1 As String=""
        '    TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
        '    Dim sFENTERDATE2 As String=TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)
        '    'Dim sFENTERDATE2 As String=TIMS.GET_FENTERDATE2(Hid_RID1.Value, sFENTERDATE, sEXAMDATE, objconn)
        '    If sFENTERDATE2 <> "" Then
        '        FEnterDate2.Text=TIMS.cdate3(sFENTERDATE2)
        '        Common.SetListItem(HR5, CDate(sFENTERDATE2).Hour)
        '        Common.SetListItem(MM5, CDate(sFENTERDATE2).Minute)
        '    End If
        'End If
        'If Convert.ToString(row_list("FEnterDate2")) <> "" Then
        '    '不為空值
        '    FEnterDate2.Text=TIMS.cdate3(row_list("FEnterDate2"))
        '    Common.SetListItem(HR5, CDate(row_list("FEnterDate2")).Hour)
        '    Common.SetListItem(MM5, CDate(row_list("FEnterDate2")).Minute)
        'End If

        'FEnterDate2.Enabled=False
        'HR5.Enabled=False
        'MM5.Enabled=False
        'Dim sTip2 As String="請輸入「報名結束日期」及「甄試日期」計算後帶出「報名登錄最晚可作業日期」"
        Dim sTip2 As String = "請輸入「報名登錄最晚可作業日期」"
        TIMS.Tooltip(FEnterDate2, sTip2)
        TIMS.Tooltip(MM5, sTip2)
        TIMS.Tooltip(MM5, sTip2)

        Dim flag_lock1 As Boolean = True
        '排除產投計畫
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_lock1 = False
        '排除自辦在職計畫
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_lock1 = False

        If flag_lock1 Then
            TB_SEnterDate.Enabled = False
            HR1.Enabled = False
            MM1.Enabled = False
            date1.Visible = False
            TB_FEnterDate.Enabled = False
            HR2.Enabled = False
            MM2.Enabled = False
            date2.Visible = False
            ExamDate.Enabled = False
            ImgExamDate.Visible = False
            ExamPeriod.Enabled = False
            Dim sTip As String = "若要修改,請依規定進行班級變更申請,並告知承辦人."
            TIMS.Tooltip(TB_SEnterDate, sTip)
            TIMS.Tooltip(HR1, sTip)
            TIMS.Tooltip(HR2, sTip)
            TIMS.Tooltip(TB_FEnterDate, sTip)
            TIMS.Tooltip(MM1, sTip)
            TIMS.Tooltip(MM2, sTip)
            TIMS.Tooltip(ExamDate, sTip)
            TIMS.Tooltip(ExamPeriod, sTip)
        End If

#Region "(No Use)"

        '是否報名時間已過
        '報名時間已過，不可修改報名時間
        '且 sm.UserInfo.OrgLevel > 1 '委訓 or 縣市政府
        '排除產投計畫
        'If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    'EnteredFlag : 是否報名時間已過
        '    If EnteredFlag AndAlso sm.UserInfo.OrgLevel > 1 Then
        '        TB_SEnterDate.Enabled=False
        '        HR1.Enabled=False
        '        MM1.Enabled=False
        '        date1.Visible=False

        '        TB_FEnterDate.Enabled=False
        '        HR2.Enabled=False
        '        MM2.Enabled=False
        '        date2.Visible=False

        '        Dim sTip As String="若要修改,請依規定進行班級變更申請,並告知承辦人."
        '        TIMS.Tooltip(TB_SEnterDate, sTip)
        '        TIMS.Tooltip(TB_FEnterDate, sTip)
        '        TIMS.Tooltip(HR1, sTip)
        '        TIMS.Tooltip(HR2, sTip)
        '        TIMS.Tooltip(MM1, sTip)
        '        TIMS.Tooltip(MM2, sTip)
        '    End If
        'End If

        'If Convert.IsDBNull(row_list("LevelCount")) Then
        '    Common.SetListItem(TB_LevelType, "")  '課程階段
        '    LevelName.Items.Clear()
        '    LevelName.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        'Else
        '    If row_list("LevelCount")="0" Then '選擇無階段時
        '        Common.SetListItem(TB_LevelType, "0") '課程階段
        '        LevelName.Enabled=False
        '        LevelSDate.Enabled=False
        '        LevelEDate.Enabled=False
        '        LevelHour.Enabled=False
        '        add_but.Enabled=False
        '    Else
        '        Common.SetListItem(TB_LevelType, row_list("LevelCount")) '課程階段
        '    End If
        '    TB_LevelType_SelectedIndexChanged(sender, e)
        '    Call TB_LevelType_Selected1()
        'End If

#End Region

        TB_Content.Text = Convert.ToString(row_list("Content"))
        TB_Purpose.Text = Convert.ToString(row_list("Purpose"))
        Common.SetListItem(rblADVANCE, row_list("ADVANCE")) '訓練課程類型 ADVANCE
        TB_TNum.Text = Convert.ToString(row_list("TNum"))
        TB_THours.Text = Convert.ToString(row_list("THours"))
        TB_STDate.Text = TIMS.Cdate3(row_list("STDate"))
        TB_FTDate.Text = TIMS.Cdate3(row_list("FTDate"))

        '導師
        CTName.Text = TIMS.Get_CTNAME1(Convert.ToString(row_list("CTName")))

        city_code.Value = Convert.ToString(row_list("TaddressZip"))
        hidTaddressZIP6W.Value = Convert.ToString(row_list("TaddressZIP6W"))
        TaddressZIPB3.Value = TIMS.GetZIPCODEB3(hidTaddressZIP6W.Value)
        TBCity.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(row_list("TaddressZip")), Convert.ToString(row_list("TaddressZIP6W")))
        TBaddress.Text = Convert.ToString(row_list("TAddress"))

        EZip_Code.Value = Convert.ToString(row_list("EADDRESSZIP"))
        hidEADDRESSZIP6W.Value = Convert.ToString(row_list("EADDRESSZIP6W"))
        EADDRESSZIPB3.Value = TIMS.GetZIPCODEB3(hidEADDRESSZIP6W.Value)
        ECity.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(row_list("EADDRESSZIP")), Convert.ToString(row_list("EADDRESSZIP6W")))
        EADDRESS.Text = Convert.ToString(row_list("EADDRESS"))

        hidEZip_Code.Value = EZip_Code.Value
        hidEADDRESSZIPB3.Value = EADDRESSZIPB3.Value
        hidhidEADDRESSZIP6W.Value = hidEADDRESSZIP6W.Value
        hidECity.Value = ECity.Text
        hidEADDRESS.Value = EADDRESS.Text

        '若沒有 甄試地點 資料，直接填入 上課資料
        CheckBox1.Checked = False

        RIDValue.Value = row_list("RID")
        LabAdd.Text = TIMS.Get_LabAddforRELSHIP(row_list("RELSHIP"), objconn)

        Common.SetListItem(TDeadline_List, Convert.ToString(row_list("TDeadline")))
        'TPeriodList
        Common.SetListItem(TPeriodList, Convert.ToString(row_list("TPeriod")))
        TB_NOTE3.Text = Convert.ToString(row_list("NOTE3"))

        CB_IsApplic.Checked = False
        If Convert.ToString(row_list("IsApplic")) = "Y" Then CB_IsApplic.Checked = True

        CB_NotOpen.Checked = False
        If Convert.ToString(row_list("NotOpen")) = "Y" Then CB_NotOpen.Checked = True

        TIMS.SetCblValue(NORID, Convert.ToString(row_list("NORID")))
        NORIDValue.Value = TIMS.GetCblValue(NORID)

        'For i As Integer=0 To Split(row_list("NORID").ToString, ",").Length - 1
        '    For j As Integer=0 To NORID.Items.Count - 1
        '        If Split(row_list("NORID").ToString, ",")(i)=NORID.Items(j).Value Then
        '            NORID.Items(j).Selected=True
        '            If NORIDValue.Value <> "" Then NORIDValue.Value &= ","
        '            NORIDValue.Value += NORID.Items(j).Value
        '        End If
        '    Next
        'Next
        OtherReason.Text = Convert.ToString(row_list("OtherReason"))

        '報到日期(有值覆蓋)
        TB_CheckInDate.Text = ""
        If Convert.ToString(row_list("CheckInDate")) <> "" Then TB_CheckInDate.Text = Common.FormatDate(row_list("CheckInDate"))

        '2005/5/18 新增"是否為法定全日制"欄位--Melody
        'Dim vISFullDate As String=""
        'If Not Convert.IsDBNull(row_list("ISFullDate")) Then
        '    vISFullDate="N"
        '    If row_list("ISFullDate")="Y" Then vISFullDate="Y"
        'End If
        'Common.SetListItem(Radio_isfulldate, vISFullDate)

        '問卷調查起始日
        'TB_QaySDate.Text=""
        'If Convert.ToString(row_list("QaySDate")) <> "" Then
        '    TB_QaySDate.Text=TIMS.cdate3(row_list("QaySDate"))
        '    Common.SetListItem(HR3, TIMS.AddZero(CDate(row_list("QaySDate")).Hour, 2))
        '    Common.SetListItem(MM3, TIMS.AddZero(CDate(row_list("QaySDate")).Minute, 2))
        'End If
        '問卷調查結束日
        'TB_QayFDate.Text=""
        'If Convert.ToString(row_list("QayFDate")) <> "" Then
        '    TB_QayFDate.Text=TIMS.cdate3(row_list("QayFDate"))
        '    Common.SetListItem(HR4, TIMS.AddZero(CDate(row_list("QayFDate")).Hour, 2))
        '    Common.SetListItem(MM4, TIMS.AddZero(CDate(row_list("QayFDate")).Minute, 2))
        'End If

        'Dim plan_name As String
        'Dim sqlstr_PlanName As String="select OrgName  from  Auth_Relship a join ORG_ORGINFO b on a.orgid =b.orgid where a.rid='" & row_list("RID") & "'"
        'plan_name=Convert.ToString(DbAccess.ExecuteScalar(sqlstr_PlanName, objconn))
        'Me.TBplan.Text=plan_name
        TBplan.Text = TIMS.Get_OrgNameInputRID(CStr(row_list("RID")), objconn)
    End Sub

    Sub SHOW_PLANINFO()
        'Dim s_log1 As String=""
        's_log1 &= "##SHOW_PLANINFO" & vbCrLf
        's_log1 &= String.Format("PlanIDValue.Value: {0}", PlanIDValue.Value) & vbCrLf
        's_log1 &= String.Format("P_ComIDNO.Value: {0}", P_ComIDNO.Value) & vbCrLf
        's_log1 &= String.Format("P_SeqNO.Value: {0}", P_SeqNO.Value) & vbCrLf
        'TIMS.LOG.Debug(s_log1)

        'PlanIDValue.Value=temp_dr("PlanID").ToString()
        'P_ComIDNO.Value=temp_dr("ComIDNO").ToString()
        'P_SeqNO.Value=temp_dr("SeqNO").ToString()
        If PlanIDValue.Value = "" Then Return
        If P_ComIDNO.Value = "" Then Return
        If P_SeqNO.Value = "" Then Return

        Dim parms As New Hashtable From {{"PLANID", PlanIDValue.Value}, {"COMIDNO", P_ComIDNO.Value}, {"SEQNO", P_SeqNO.Value}}
        Dim sql_pp As String = "SELECT * FROM PLAN_PLANINFO WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        Dim dr_pp As DataRow = DbAccess.GetOneRow(sql_pp, objconn, parms)
        If dr_pp Is Nothing Then Return

        Dim v_rblADVANCE As String = TIMS.GetListValue(rblADVANCE) '訓練課程類型
        If v_rblADVANCE = "" Then Common.SetListItem(rblADVANCE, dr_pp("ADVANCE")) '訓練課程類型

        'rblADVANCE.Enabled=False
        'TIMS.Tooltip(rblADVANCE, cst_TooltipT1)
    End Sub

    ''' <summary> 檢核問題。先基礎檢核，再檢核邏輯問題 </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""
        'TB_SEnterDate,'TB_SEnterDate,'ExamDate,'FEnterDate2,'TB_QaySDate,'TB_QayFDate,'TB_STDate,'TB_FTDate,'TB_CheckInDate,'LevelSDate,'LevelEDate
        TB_TNum.Text = TIMS.ClearSQM(TB_TNum.Text)
        TB_SEnterDate.Text = TIMS.ClearSQM(TB_SEnterDate.Text)
        TB_FEnterDate.Text = TIMS.ClearSQM(TB_FEnterDate.Text)
        ExamDate.Text = TIMS.ClearSQM(ExamDate.Text)
        FEnterDate2.Text = TIMS.ClearSQM(FEnterDate2.Text)
        'TB_QaySDate.Text=TIMS.ClearSQM(TB_QaySDate.Text)
        'TB_QayFDate.Text=TIMS.ClearSQM(TB_QayFDate.Text)
        TB_STDate.Text = TIMS.ClearSQM(TB_STDate.Text)
        TB_FTDate.Text = TIMS.ClearSQM(TB_FTDate.Text)
        TB_CheckInDate.Text = TIMS.ClearSQM(TB_CheckInDate.Text)
        'LevelSDate.Text=TIMS.ClearSQM(LevelSDate.Text)
        'LevelEDate.Text=TIMS.ClearSQM(LevelEDate.Text)

        'OJT-23041104：區域據點-開班資料查詢：【班級英文名稱】改為非必填 
        If Not flag_TPlanID70_1 Then
            Dim vClassEngName As String = TIMS.ClearSQM(ClassEngName.Text)
            If vClassEngName = "" Then Errmsg &= "請輸入班級英文名稱!" & vbCrLf
        End If

        If TB_TNum.Text <> "" AndAlso Not TIMS.IsNumeric2(TB_TNum.Text) Then Errmsg &= "訓練人數 數字格式有誤，僅提供輸入大於0的數字!" & vbCrLf

        If Not TIMS.IsDate1(TB_SEnterDate.Text) Then Errmsg &= "報名開始日期 日期格式有誤!" & vbCrLf

        If Not TIMS.IsDate1(TB_FEnterDate.Text) Then Errmsg &= "報名結束日期 日期格式有誤!" & vbCrLf

        If Not TIMS.IsDate1(TB_STDate.Text) Then Errmsg &= "開訓日期 日期格式有誤!" & vbCrLf

        If Not TIMS.IsDate1(TB_FTDate.Text) Then Errmsg &= "結訓日期 日期格式有誤!" & vbCrLf

        If ExamDate.Text <> "" AndAlso Not TIMS.IsDate1(ExamDate.Text) Then Errmsg &= "甄試日期 日期格式有誤!" & vbCrLf

        If FEnterDate2.Text <> "" AndAlso Not TIMS.IsDate1(FEnterDate2.Text) Then Errmsg &= "報名登錄最晚 日期格式有誤!" & vbCrLf

        'If TB_QaySDate.Text <> "" AndAlso Not TIMS.IsDate1(TB_QaySDate.Text) Then Errmsg &= "問卷調查開始日期 日期格式有誤!" & vbCrLf
        'If TB_QayFDate.Text <> "" AndAlso Not TIMS.IsDate1(TB_QayFDate.Text) Then Errmsg &= "問卷調查結束日期 日期格式有誤!" & vbCrLf

        If TB_CheckInDate.Text <> "" AndAlso Not TIMS.IsDate1(TB_CheckInDate.Text) Then Errmsg &= "報到日期 日期格式有誤!" & vbCrLf

        'If LevelSDate.Text <> "" AndAlso Not TIMS.IsDate1(LevelSDate.Text) Then Errmsg &= "階段起始日 日期格式有誤!" & vbCrLf
        'If LevelEDate.Text <> "" AndAlso Not TIMS.IsDate1(LevelEDate.Text) Then Errmsg &= "階段結束日 日期格式有誤!" & vbCrLf

        rst = If(Errmsg <> "", False, True)
        If Errmsg <> "" Then Return rst

        TB_SEnterDate.Text = TIMS.Cdate3(TB_SEnterDate.Text)
        TB_FEnterDate.Text = TIMS.Cdate3(TB_FEnterDate.Text)
        TB_STDate.Text = TIMS.Cdate3(TB_STDate.Text)
        TB_FTDate.Text = TIMS.Cdate3(TB_FTDate.Text)
        'Dim strScript6 As String
        Dim oSEnterDate As Date = CDate(TB_SEnterDate.Text)
        Dim oFEnterDate As Date = CDate(TB_FEnterDate.Text)
        Dim oSTDate As Date = CDate(TB_STDate.Text)
        Dim oFTDate As Date = CDate(TB_FTDate.Text)

        'If RB_TPropertyID.SelectedValue="2" AndAlso Companyname.Text.ToString="" Then '企業名稱
        '    Errmsg &= "企業名稱為必填" & vbCrLf
        '    'Common.MessageBox(Me, "企業名稱為必填")
        '    'Return 'Exit Sub
        'End If
        'If RB_TPropertyID.SelectedValue="" Then
        '    '未設定 訓練性質
        '    Errmsg &= "請選擇訓練性質為必選!" & vbCrLf
        '    'Common.MessageBox(Me, "請選擇訓練性質為必選!")
        '    'Return 'Exit Sub
        'End If

        If (CDate(oSEnterDate) >= CDate(oFEnterDate)) Then
            Errmsg &= "[報名結束日期]必須大於[報名開始日期]!" & vbCrLf
            'Common.MessageBox(Me, "[報名結束日期]必須大於[報名開始日期]!")
            'Return 'Exit Sub
        End If
        If (CDate(oSTDate) <= CDate(oFEnterDate)) Then
            Errmsg &= "[開訓日期]必須大於[報名結束日期]!" & vbCrLf
            'Common.MessageBox(Me, "[開訓日期]必須大於[報名結束日期]!")
            'Return 'Exit Sub
        End If

        '090608 andy edit 
        ExamDate.Text = TIMS.Cdate3(ExamDate.Text)
        If ExamDate.Text <> "" Then
            If ExamPeriod.SelectedIndex = 0 Then '20100329 add 甄試時段
                Errmsg &= "「甄試日期」全天、上午、下午 時段請擇一選擇!" & vbCrLf
                'Common.MessageBox(Me, "「甄試日期」全天、上午、下午 時段請擇一選擇!")
                'Return 'Exit Sub
            End If
            If (CDate(ExamDate.Text) <= CDate(oFEnterDate)) Then
                Errmsg &= "「甄試日期」必須大於「報名結束日期」!" & vbCrLf
                'Common.MessageBox(Me, "「甄試日期」必須大於「報名結束日期」!")
                'Return 'Exit Sub
            End If
            If (CDate(ExamDate.Text) > CDate(oSTDate)) Then
                Errmsg &= "[甄試日期]必須小於或等於[開訓日期]!" & vbCrLf
                'Common.MessageBox(Me, "[甄試日期]必須小於或等於[開訓日期]!")
                'Return 'Exit Sub
            End If
            'If Errmsg="" Then
            '    Dim flagExamDateOk As Boolean=True
            '    If Not DateDiff(DateInterval.Day, CDate(FEnterDate).AddDays(2), CDate(ExamDate.Text)) >= 0 Then flagExamDateOk=False
            '    If Not DateDiff(DateInterval.Day, CDate(ExamDate.Text), CDate(FEnterDate).AddDays(7)) >= 0 Then flagExamDateOk=False
            '    If Not flagExamDateOk Then
            '        Errmsg &= "甄試日期：應為[報名結束日期]+2~7日曆天" & vbCrLf
            '    End If
            'End If
        End If

        '檢核-不開班
        If Not CB_NotOpen.Checked Then Errmsg = ""
        '確定-開班-檢核正確性
        rst = If(Errmsg <> "", False, True)
        If Errmsg <> "" Then Return rst
        '請選擇訓練性質 RB_TPropertyID
        '請選擇報名開始日期 TB_SEnterDate
        '請選擇報名結束日期 TB_FEnterDate
        '請選擇問卷調查開始日期 TB_QaySDate
        '請選擇問卷調查結束日期 TB_QayFDate
        '請選擇訓練期限 TDeadline_List
        '請選擇訓練時段 TPeriodList
        '請選擇是否為法定全日制 Radio_isfulldate
        '請選擇報到日期 TB_CheckInDate
        '請選擇課程階段 TB_LevelType

        'If TB_SEnterDate.Text <> "" Then TB_SEnterDate.Text=TIMS.ClearSQM(TB_SEnterDate.Text)
        'If TB_FEnterDate.Text <> "" Then TB_FEnterDate.Text=TIMS.ClearSQM(TB_FEnterDate.Text)
        'If ExamDate.Text <> "" Then ExamDate.Text=TIMS.ClearSQM(ExamDate.Text)
        'If FEnterDate2.Text <> "" Then FEnterDate2.Text=TIMS.ClearSQM(FEnterDate2.Text)
        'If TB_QaySDate.Text <> "" Then TB_QaySDate.Text=TIMS.ClearSQM(TB_QaySDate.Text)
        'If TB_QayFDate.Text <> "" Then TB_QayFDate.Text=TIMS.ClearSQM(TB_QayFDate.Text)
        'If TB_CheckInDate.Text <> "" Then TB_CheckInDate.Text=TIMS.ClearSQM(TB_CheckInDate.Text)
        'If RB_TPropertyID.SelectedValue="" Then Errmsg &= "請選擇訓練性質" & vbCrLf
        If TB_SEnterDate.Text = "" Then Errmsg &= "請選擇報名開始日期" & vbCrLf
        If TB_FEnterDate.Text = "" Then Errmsg &= "請選擇報名結束日期" & vbCrLf
        If Not flag_TPlanID70_1 Then
            If ExamDate.Text = "" Then Errmsg &= "請選擇甄試日期 " & vbCrLf
        End If

        '確定-開班-檢核正確性
        rst = If(Errmsg <> "", False, True)
        If Errmsg <> "" Then Return rst

        If Errmsg = "" Then
            'Dim sFENTERDATE2 As String=TIMS.GET_FENTERDATE2(sFENTERDATE, sEXAMDATE)
            Dim sFENTERDATE As String = TIMS.Cdate3(TB_FEnterDate.Text)
            Dim sEXAMDATE As String = TIMS.Cdate3(ExamDate.Text)
            Dim SS1 As String = ""
            TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
            'Dim chk_FENTERDATE2 As String=TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)

            Dim sHR5 As String = ""
            Dim sMM5 As String = ""
            Call TIMS.sUtl_GetHRMM(sHR5, sMM5, HR5.SelectedValue, MM5.SelectedValue)
            Dim strFEnterDate2x As String = String.Concat(FEnterDate2.Text, " ", sHR5, ":", sMM5) 'CDate()
            Dim strFENTERDATE2 As String = TIMS.Cdate3(If(TIMS.IsDate1(strFEnterDate2x), strFEnterDate2x, ""))

            If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                If strFENTERDATE2 = "" Then Errmsg &= "報名登錄最晚可作業時間 不可為空白，請確認報名結束日期正確性" & vbCrLf
            End If

            If Errmsg = "" AndAlso Not TIMS.Chk_FENTERDATE2(sFENTERDATE, strFENTERDATE2) Then
                Errmsg &= String.Format("報名登錄最晚可作業時間 報名結束日期 到 (報名結束日期+3)日曆天(ERROR:{0})", TIMS.STR2NUL(strFENTERDATE2)) & vbCrLf
            End If
            'FEnterDate2.Text=sFENTERDATE2 'FEnterDate2.Text
            'If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID)=-1 Then
            '    If sFENTERDATE2="" Then Errmsg &= "報名登錄最晚可作業時間 不可為空白，請確認報名結束日期與甄試日期正確性" & vbCrLf
            'End If
            'If sFENTERDATE2 <> "" Then
            '    FEnterDate2.Text=TIMS.cdate3(sFENTERDATE2)
            '    Common.SetListItem(HR5, CDate(sFENTERDATE2).Hour)
            '    Common.SetListItem(MM5, CDate(sFENTERDATE2).Minute)
            'End If
        End If
        'If FEnterDate2.Text="" Then
        'End If
        'If FEnterDate2.Text="" Then Errmsg &= "報名登錄最晚可作業時間 不可為空白，請點選計算" & vbCrLf
        'If TB_QaySDate.Text="" Then Errmsg &= "請選擇問卷調查開始日期" & vbCrLf
        'If TB_QayFDate.Text="" Then Errmsg &= "請選擇問卷調查結束日期" & vbCrLf
        'If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '甄試日期 自辦在職 為必填
        '    If ExamDate.Text="" Then Errmsg &= "請選擇甄試日期 " & vbCrLf
        'End If
        If TDeadline_List.SelectedValue = "" Then Errmsg &= "請選擇訓練期限" & vbCrLf

        '顯示此行才做必要判斷(TRUE/FLSE)
        Dim fg_CHECK_EADDRESS As Boolean = trEADDRESS.Visible '(顯示就檢測)
        If flag_TPlanID70_1 Then fg_CHECK_EADDRESS = False ''70:區域產業據點職業訓練計畫(在職) 不檢測 甄試地點
        If fg_CHECK_EADDRESS Then
            '顯示此行才做必要判斷
            If EZip_Code.Value = "" OrElse ECity.Text = "" Then Errmsg &= "甄試地點 郵遞區號前3碼不可為空。" & vbCrLf
            If EADDRESSZIPB3.Value = "" Then Errmsg &= "甄試地點 郵遞區號後3碼或2碼不可為空。" & vbCrLf
            If EADDRESS.Text = "" Then Errmsg &= "甄試地點 不可為空。" & vbCrLf
        End If

        If TPeriodList.SelectedValue = "" Then Errmsg &= "請選擇訓練時段" & vbCrLf

        '自辦在職顯示此功能 且為必填
        If trTB_NOTE3.Visible Then
            If TB_NOTE3.Text = "" Then Errmsg &= "請輸入訓練時段下方的上課時間" & vbCrLf
        End If

        'If Radio_isfulldate.SelectedValue="" Then Errmsg &= "請選擇是否為法定全日制" & vbCrLf
        If TB_CheckInDate.Text = "" Then Errmsg &= "請選擇報到日期" & vbCrLf
        'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
        'If TB_LevelType.SelectedValue="" Then Errmsg &= "請選擇課程階段" & vbCrLf

        Dim v_rblADVANCE As String = TIMS.GetListValue(rblADVANCE) '訓練課程類型
        If v_rblADVANCE = "" Then Errmsg &= "訓練課程類型 單選必填!" & vbCrLf

        '檢核-不開班
        If Not CB_NotOpen.Checked Then Errmsg = ""
        '確定-開班-檢核正確性
        rst = If(Errmsg <> "", False, True)
        If Errmsg <> "" Then Return rst

        Return rst
    End Function

    ''' <summary> 儲存前再次檢核 </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData2(ByRef Errmsg As String) As Boolean
        clsid.Value = TIMS.ClearSQM(clsid.Value)
        PlanIDValue.Value = TIMS.ClearSQM(PlanIDValue.Value)
        TB_CyclType.Text = TIMS.ClearSQM(TB_CyclType.Text)
        TB_ClassName.Text = TIMS.ClearSQM(TB_ClassName.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        P_ComIDNO.Value = TIMS.ClearSQM(P_ComIDNO.Value)
        P_SeqNO.Value = TIMS.ClearSQM(P_SeqNO.Value)

        TB_STDate.Text = TIMS.ClearSQM(TB_STDate.Text)
        TB_FTDate.Text = TIMS.ClearSQM(TB_FTDate.Text)

        Dim dt1 As DataTable = Nothing

        Select Case ProcessType
            Case cst_Insert
                Dim parms_B As New Hashtable From {
                    {"CLSID", clsid.Value},
                    {"PlanID", PlanIDValue.Value},
                    {"CyclType", TB_CyclType.Text},
                    {"CLASSCNAME", TB_ClassName.Text},
                    {"RID", RIDValue.Value}
                }
                Dim sqlstr_B As String = ""
                sqlstr_B &= " SELECT 1 FROM CLASS_CLASSINFO"
                sqlstr_B &= " WHERE NOTOPEN='N' AND CLSID=@CLSID AND PlanID=@PlanID AND ISNULL(CyclType,'')=@CyclType AND CLASSCNAME=@CLASSCNAME AND RID=@RID"
                'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
                'sqlstr_B &= " AND isnull(LevelType,'')='" & v_TB_LevelType & "' AND RID='" & RIDValue.Value & "') "
                dt1 = DbAccess.GetDataTable(sqlstr_B, objconn, parms_B)
                If dt1.Rows.Count > 0 Then '>0資料重複
                    Errmsg = "新增開班資料重複!!"
                    Return False
                End If

                parms_B.Clear()
                parms_B.Add("PLANID", PlanIDValue.Value)
                parms_B.Add("COMIDNO", P_ComIDNO.Value)
                parms_B.Add("SEQNO", P_SeqNO.Value)
                sqlstr_B = ""
                sqlstr_B &= " SELECT 1 FROM CLASS_CLASSINFO"
                sqlstr_B &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                dt1 = DbAccess.GetDataTable(sqlstr_B, objconn, parms_B)
                If dt1.Rows.Count > 0 Then '>0資料重複
                    Errmsg = "新增開班資料重複!!!"
                    Return False
                End If

                Return True
            Case cst_Update, cst_PlanUpdate
                'Session("ClassLevel")=dtlevel
                'Dim dtlevel1 As DataTable=Nothing
                'If Session("ClassLevel") IsNot Nothing Then dtlevel1=Session("ClassLevel")
                'Dim strScript As String=""
                rqOCID = TIMS.ClearSQM(rqOCID)
                If rqOCID <> "" Then
                    Dim drCC As DataRow = TIMS.GetOCIDDate(rqOCID, objconn)
                    If drCC Is Nothing Then
                        Errmsg = "開班資料傳入參數有誤!!"
                        rqOCID = ""
                        Return False
                    End If
                End If

                '(如果不開班，就不檢查了)
                If CB_NotOpen.Checked Then Return True

                Dim parms_B As New Hashtable From {
                    {"CLSID", clsid.Value},
                    {"PlanID", PlanIDValue.Value},
                    {"CyclType", TB_CyclType.Text},
                    {"CLASSCNAME", TB_ClassName.Text},
                    {"RID", RIDValue.Value}
                }
                Dim sqlstr_B As String = ""
                sqlstr_B &= " SELECT 1 FROM CLASS_CLASSINFO"
                sqlstr_B &= " WHERE NOTOPEN='N' AND CLSID=@CLSID AND PlanID=@PlanID AND ISNULL(CyclType,'')=@CyclType AND CLASSCNAME=@CLASSCNAME AND RID=@RID"
                If rqOCID <> "" Then
                    sqlstr_B &= String.Concat(" AND OCID != ", rqOCID, "")
                    sqlstr_B &= String.Concat(" AND STDATE='", TIMS.Cdate3(TB_STDate.Text), "'")
                    sqlstr_B &= String.Concat(" AND FTDATE='", TIMS.Cdate3(TB_FTDate.Text), "'")
                End If
                dt1 = DbAccess.GetDataTable(sqlstr_B, objconn, parms_B)
                If dt1.Rows.Count > 0 Then '開班資料重複!!!!
                    Errmsg &= "開班資料重複!!!!"
                    Return False
                End If

                parms_B.Clear()
                parms_B.Add("PLANID", PlanIDValue.Value)
                parms_B.Add("COMIDNO", P_ComIDNO.Value)
                parms_B.Add("SEQNO", P_SeqNO.Value)
                sqlstr_B = ""
                sqlstr_B &= " SELECT 1 FROM CLASS_CLASSINFO" & vbCrLf
                sqlstr_B &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
                If rqOCID <> "" Then sqlstr_B &= " AND OCID <> '" & rqOCID & "' " & vbCrLf
                dt1 = DbAccess.GetDataTable(sqlstr_B, objconn, parms_B)
                If dt1.Rows.Count > 0 Then '開班資料重複!!!
                    Errmsg = "開班資料重複!!!"
                    Return False
                End If

                'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
        End Select

        'If Session(cst_ClassSearchStr) Is Nothing AndAlso ViewState(cst_ClassSearchStr) IsNot Nothing Then
        '    Session(cst_ClassSearchStr)=ViewState(cst_ClassSearchStr)
        'End If
        Return True
    End Function

    '[儲存] CLASS_CLASSINFO/PLAN_PLANINFO/U
    Sub SaveData1()
        Dim Errmsg As String = ""
        '前端驗證是否有錯誤。
        Dim rstPageIsValid As Boolean = Page.IsValid
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Return 'Exit Sub
        End If

        'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
        'Dim v_TB_LevelType As String=TIMS.GetListValue(TB_LevelType)

        Dim flag_save_ok As Boolean = False 'true:轉入成功 /false:轉入失敗
        Dim str_save_data_All As String = "" '要儲存的資料全部資訊

        Dim sqlstr_update As String = "" 'sql string

        Dim sqldr As DataRow = Nothing 'CLASS_CLASSINFO
        Dim sqlAdapter As SqlDataAdapter = Nothing 'CLASS_CLASSINFO
        Dim sqlTable As New DataTable 'CLASS_CLASSINFO

        Dim sqldr_pp As DataRow = Nothing 'PlanInfo
        Dim daPlanInfo As SqlDataAdapter = Nothing 'PlanInfo
        Dim dtPlanInfo As DataTable = Nothing 'PlanInfo
        Dim sql As String = ""

        If Session(cst_ClassSearchStr) Is Nothing AndAlso ViewState(cst_ClassSearchStr) IsNot Nothing Then
            Session(cst_ClassSearchStr) = ViewState(cst_ClassSearchStr)
        End If

        TB_CyclType.Text = TIMS.FmtCyclType(TB_CyclType.Text)

        Dim strScript As String = ""
        Dim iCheck_Nextclass As Integer = 0 '(查詢是否有 未開班完成的資料)

        '再次檢核
        Dim flag_CHECK2 As Boolean = CheckData2(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Return 'Exit Sub
        End If

        Select Case ProcessType
            Case cst_Insert
                'CLASS_CLASSINFO-GetInsertRow
                sqldr = DbAccess.GetInsertRow("CLASS_CLASSINFO", sqlTable, sqlAdapter, objconn)
            Case cst_Update, cst_PlanUpdate
                If rqOCID <> "" Then
                    'CLASS_CLASSINFO-GetUpdateRow
                    sqlstr_update = String.Format(" SELECT * FROM CLASS_CLASSINFO WHERE OCID={0}", rqOCID)
                    sqldr = DbAccess.GetUpdateRow(sqlstr_update, sqlTable, sqlAdapter, objconn)
                Else
                    'CLASS_CLASSINFO-GetInsertRow
                    sqldr = DbAccess.GetInsertRow("CLASS_CLASSINFO", sqlTable, sqlAdapter, objconn)
                End If

                If ProcessType = cst_PlanUpdate AndAlso rqOCID <> "" Then
                    Dim chk_parms As New Hashtable From {
                        {"OCID", rqOCID},
                        {"CLSID", clsid.Value},
                        {"PlanID", PlanIDValue.Value},
                        {"CyclType", TB_CyclType.Text},
                        {"RID", RIDValue.Value}
                    }
                    Dim chk_nextclass As String = "SELECT OCID FROM CLASS_CLASSINFO WHERE OCID > @OCID and CLSID=@CLSID AND PlanID=@PlanID and ISNULL(CyclType,'')=@CyclType and RID=@RID"
                    Dim dt1_chk As DataTable = DbAccess.GetDataTable(chk_nextclass, objconn, chk_parms)
                    iCheck_Nextclass = dt1_chk.Rows.Count 'DbAccess.GetCount(check_nextclass, objconn)
                End If

        End Select

        Dim sHR1 As String = "" : Dim sHR2 As String = ""
        'Dim sHR3 As String="" 'Dim sHR4 As String=""
        Dim sMM1 As String = "" : Dim sMM2 As String = ""
        'Dim sMM3 As String="" 'Dim sMM4 As String=""
        Dim sHR5 As String = "" : Dim sMM5 As String = ""
        Dim sHR6 As String = "" : Dim sMM6 As String = ""
        Call TIMS.sUtl_GetHRMM(sHR1, sMM1, HR1.SelectedValue, MM1.SelectedValue)
        Call TIMS.sUtl_GetHRMM(sHR2, sMM2, HR2.SelectedValue, MM2.SelectedValue)
        'Call TIMS.sUtl_GetHRMM(sHR3, sMM3, HR3.SelectedValue, MM3.SelectedValue)
        'Call TIMS.sUtl_GetHRMM(sHR4, sMM4, HR4.SelectedValue, MM4.SelectedValue)
        Call TIMS.sUtl_GetHRMM(sHR5, sMM5, HR5.SelectedValue, MM5.SelectedValue)
        Call TIMS.sUtl_GetHRMM(sHR6, sMM6, HR6.SelectedValue, MM6.SelectedValue)

        Select Case ProcessType
            Case cst_Insert
                Dim sqlstr_A As String = "SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID='" & RIDValue.Value & "'"
                Dim relship As String = DbAccess.ExecuteScalar(sqlstr_A, objconn)
                Dim sqlstr_year As String = "SELECT YEARS FROM ID_PLAN WHERE PlanID='" & PlanIDValue.Value & "'"
                Dim years As String = DbAccess.ExecuteScalar(sqlstr_year, objconn)
                sqldr("Relship") = relship
                sqldr("Years") = years.Substring(2)

            Case cst_PlanUpdate
                sqldr("Relship") = P_Relship.Value
                sqldr("Years") = P_Years.Value
                sqldr("ComIDNO") = P_ComIDNO.Value
                sqldr("SeqNO") = P_SeqNO.Value
                'UPDATE PLAN_PLANINFO (dtPlanInfo)
                sqlstr_update = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & PlanIDValue.Value & "' AND ComIDNO='" & P_ComIDNO.Value & "' AND SeqNO='" & P_SeqNO.Value & "' AND AppliedResult='Y' "
                sqldr_pp = DbAccess.GetUpdateRow(sqlstr_update, dtPlanInfo, daPlanInfo, objconn)

            Case cst_Update
                PlanIDValue.Value = sqldr("PlanID")
                P_ComIDNO.Value = sqldr("ComIDNO")
                P_SeqNO.Value = sqldr("SeqNO")
                'UPDATE PLAN_PLANINFO (dtPlanInfo)
                sqlstr_update = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & PlanIDValue.Value & "' AND ComIDNO='" & P_ComIDNO.Value & "' AND SeqNO='" & P_SeqNO.Value & "' AND AppliedResult='Y' "
                sqldr_pp = DbAccess.GetUpdateRow(sqlstr_update, dtPlanInfo, daPlanInfo, objconn)

        End Select

        Dim v_rblADVANCE As String = TIMS.GetListValue(rblADVANCE) '訓練課程類型
        Dim v_ExamPeriod As String = TIMS.GetListValue(ExamPeriod)
        'UPDATE PLAN_PLANINFO
        If sqldr_pp IsNot Nothing Then
            sqldr_pp("TransFlag") = "Y"
            sqldr_pp("SEnterDate") = If(TB_SEnterDate.Text <> "", CDate(TB_SEnterDate.Text & " " & sHR1 & ":" & sMM1), Convert.DBNull)
            sqldr_pp("FEnterDate") = If(TB_FEnterDate.Text <> "", CDate(TB_FEnterDate.Text & " " & sHR2 & ":" & sMM2), Convert.DBNull)
            '甄試日期/時間 'If TIMS.GFG_OJT_25050801_NoUse_ExamDateTime Then  sqldr_pp("ExamDate")=If(ExamDate.Text <> "", TIMS.Cdate2(ExamDate.Text), Convert.DBNull)            'End If
            sqldr_pp("ExamDate") = If(ExamDate.Text <> "", CDate(ExamDate.Text & " " & sHR6 & ":" & sMM6), Convert.DBNull)
            '20100329 andy add 甄試日期(時段)
            sqldr_pp("ExamPeriod") = If(ExamDate.Text <> "" AndAlso v_ExamPeriod <> "", v_ExamPeriod, Convert.DBNull)
            'sqldr_new("FEnterDate2")=CDate(FEnterDate2.Text & " " & sHR5 & ":" & sMM5)
            sqldr_pp("ADVANCE") = If(v_rblADVANCE <> "", v_rblADVANCE, Convert.DBNull) '訓練課程類型
            sqldr_pp("ModifyAcct") = sm.UserInfo.UserID
            sqldr_pp("ModifyDate") = Now()
        End If

        'CLASS_CLASSINFO
        sqldr("PlanID") = PlanIDValue.Value
        sqldr("RID") = RIDValue.Value
        sqldr("ADVANCE") = If(v_rblADVANCE <> "", v_rblADVANCE, Convert.DBNull) '訓練課程類型
        'sqldr("LevelType")=TB_LevelType.SelectedValue
        '2005/5/27班級單元-Melody,for 學習卷
        sqldr("Class_Unit") = tb_class_unit.Value
        'sqldr("LevelCount")=If(v_TB_LevelType <> "", Val(v_TB_LevelType), Convert.DBNull)
        sqldr("CLSID") = clsid.Value
        sqldr("ClassCName") = TIMS.ClearSQM(TB_ClassName.Text)

        Dim vCyclType As String = TIMS.ClearSQM(TB_CyclType.Text)
        If vCyclType = "" Then vCyclType = TIMS.cst_Default_CyclType
        vCyclType = TIMS.FmtCyclType(vCyclType)
        sqldr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull) 'TB_CyclType.Text

        ClassEng.Value = TIMS.ClearSQM(ClassEng.Value)
        If ClassEng.Value <> "" AndAlso ClassEng.Value <> ClassEngName.Text Then ClassEngName.Text = ClassEng.Value
        sqldr("ClassEngName") = If(ClassEngName.Text <> "", ClassEngName.Text, Convert.DBNull)

        Dim vTPropertyID As String = Hid_RB_TPropertyID1.Value
        sqldr("TPropertyID") = Val(vTPropertyID) 'vTPropertyID 'RB_TPropertyID.SelectedValue
        sqldr("TMID") = trainValue.Value
        sqldr("CJOB_UNKEY") = cjobValue.Value
        sqldr("SEnterDate") = If(TB_SEnterDate.Text <> "", CDate(TB_SEnterDate.Text & " " & sHR1 & ":" & sMM1), Convert.DBNull)
        sqldr("FEnterDate") = If(TB_FEnterDate.Text <> "", CDate(TB_FEnterDate.Text & " " & sHR2 & ":" & sMM2), Convert.DBNull)
        '甄試日期/時間 'If TIMS.GFG_OJT_25050801_NoUse_ExamDateTime Then  sqldr("ExamDate")=If(ExamDate.Text <> "", TIMS.Cdate2(ExamDate.Text), Convert.DBNull) 
        sqldr("ExamDate") = If(ExamDate.Text <> "", CDate(ExamDate.Text & " " & sHR6 & ":" & sMM6), Convert.DBNull)
        '20100329 andy add 甄試日期(時段)
        sqldr("ExamPeriod") = If(ExamDate.Text <> "" AndAlso v_ExamPeriod <> "", v_ExamPeriod, Convert.DBNull)
        sqldr("FEnterDate2") = If(FEnterDate2.Text <> "", CDate(FEnterDate2.Text & " " & sHR5 & ":" & sMM5), Convert.DBNull)

        'sqldr("QaySDate")=If(TB_QaySDate.Text <> "", CDate(TB_QaySDate.Text & " " & sHR3 & ":" & sMM3), Convert.DBNull)
        'sqldr("QayFDate")=If(TB_QayFDate.Text <> "", CDate(TB_QayFDate.Text & " " & sHR4 & ":" & sMM4), Convert.DBNull)
        sqldr("Content") = TB_Content.Text
        sqldr("Purpose") = TB_Purpose.Text
        '訓練課程類型 ADVANCE
        'Dim v_rblADVANCE As String=TIMS.GetListValue(rblADVANCE) '訓練課程類型
        sqldr("ADVANCE") = If(v_rblADVANCE <> "", v_rblADVANCE, Convert.DBNull) '訓練課程類型
        sqldr("TNum") = CInt(TB_TNum.Text)

        hidTaddressZIP6W.Value = TIMS.GetZIPCODE6W(city_code.Value, TaddressZIPB3.Value)
        sqldr("TaddressZip") = city_code.Value
        sqldr("TaddressZIP6W") = hidTaddressZIP6W.Value
        sqldr("TAddress") = TBaddress.Text

        hidEADDRESSZIP6W.Value = TIMS.GetZIPCODE6W(EZip_Code.Value, EADDRESSZIPB3.Value)
        sqldr("EADDRESSZIP") = If(EZip_Code.Value <> "", EZip_Code.Value, Convert.DBNull) 'EZip_Code.Value
        sqldr("EADDRESSZIP6W") = If(hidEADDRESSZIP6W.Value <> "", hidEADDRESSZIP6W.Value, Convert.DBNull)
        sqldr("EADDRESS") = If(EADDRESS.Text <> "", EADDRESS.Text, Convert.DBNull) 'EADDRESS.Text

        sqldr("THours") = CInt(TB_THours.Text)
        sqldr("TDeadline") = TDeadline_List.SelectedValue
        sqldr("TPeriod") = TPeriodList.SelectedValue
        TB_NOTE3.Text = TIMS.ClearSQM(TB_NOTE3.Text)
        sqldr("NOTE3") = If(TB_NOTE3.Text = "", Convert.DBNull, TB_NOTE3.Text)
        sqldr("CTName") = TIMS.Get_CTNAME1(CTName.Text) 'CTName.Text

        sqldr("STDate") = CDate(TB_STDate.Text)
        sqldr("FTDate") = CDate(TB_FTDate.Text)
        sqldr("Companyname") = If(Companyname.Text <> "", Companyname.Text, Convert.DBNull) '企業名稱
        sqldr("CheckInDate") = CDate(TB_CheckInDate.Text)
        sqldr("IsApplic") = If(CB_IsApplic.Checked, "Y", "N") '納入志願
        If CB_NotOpen.Checked Then
            '不開班原因代碼
            Dim sNORID As String = ""
            If NORID.Enabled = True Then
                'sqldr("NORID")="" '不開班原因代碼
                For i As Integer = 0 To NORID.Items.Count - 1
                    If NORID.Items(i).Selected = True AndAlso NORID.Items(i).Value <> "" Then
                        If sNORID <> "" Then sNORID &= ","
                        sNORID &= NORID.Items(i).Value
                    End If
                Next
            End If
            sqldr("NotOpen") = "Y"
            sqldr("NORID") = If(sNORID = "", Convert.DBNull, sNORID) '不開班原因代碼
            sqldr("OtherReason") = If(OtherReason.Text = "", Convert.DBNull, OtherReason.Text) '不開班其他原因說明
            sqldr("LastState") = "D" 'D: 刪除(最後異動狀態)
        Else
            sqldr("NotOpen") = "N"
            sqldr("NORID") = Convert.DBNull '不開班原因代碼
            sqldr("OtherReason") = Convert.DBNull '不開班其他原因說明
            sqldr("LastState") = "M" 'M: 修改(最後異動狀態)
        End If

        '2005/5/18 新增"是否為法定全日制"欄位--Melody
        'Dim sISFullDate As String="N"
        'Dim v_Radio_isfulldate As String=TIMS.GetListValue(Radio_isfulldate)
        'If v_Radio_isfulldate="Y" Then sISFullDate="Y" '是否為法定全日制
        'sqldr("ISFullDate")=sISFullDate '是否為法定全日制(Y/N)
        sqldr("IsCalculate") = "N" '是否試算
        sqldr("IsClosed") = "N" '是否結訓
        sqldr("IsSuccess") = "Y" '是否轉入成功
        sqldr("BGTime") = 0 '勾稽次數
        sqldr("ModifyAcct") = sm.UserInfo.UserID
        sqldr("ModifyDate") = Now()

        Dim htPP As New Hashtable
        Select Case ProcessType
            Case cst_Insert, cst_PlanUpdate
                '新增一組OCID
                sqldr("LastState") = "A" 'A: 新增(最後異動狀態)
                iOCID_New = DbAccess.GetNewId(objconn, "CLASS_CLASSINFO_OCID_SEQ,CLASS_CLASSINFO,OCID")
                sqldr("OCID") = iOCID_New
                sqldr("ONSHELLDATE") = TIMS.Cdate2(CDate(Now).ToString("yyyy/MM/dd"))
                sqlTable.Rows.Add(sqldr)

                '2019-02-19 add 記操作歷程（sys_trans_log）'新增
                htPP.Clear()
                htPP.Add("TransType", TIMS.cst_TRANS_LOG_Insert)
                htPP.Add("TargetTable", "CLASS_CLASSINFO")
                htPP.Add("FuncPath", "/TC/01/TC_01_004_add")
                htPP.Add("s_WHERE", String.Format("OCID='{0}'", iOCID_New))
                TIMS.SaveTRANSLOG(sm, objconn, sqldr, htPP)

            Case Else
                'iOCID_New=rqOCID '使用相同的OCID
                '2019-02-19 add 記操作歷程（sys_trans_log）'修改
                'Dim htPP As New Hashtable
                htPP.Clear()
                htPP.Add("TransType", TIMS.cst_TRANS_LOG_Update)
                htPP.Add("TargetTable", "CLASS_CLASSINFO")
                htPP.Add("FuncPath", "/TC/01/TC_01_004_add")
                htPP.Add("s_WHERE", String.Format("OCID={0}", rqOCID))
                TIMS.SaveTRANSLOG(sm, objconn, sqldr, htPP) '修改後資料 CLASS_CLASSINFO

        End Select

        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            'INSERT/UPDATE CLASS_CLASSINFO
            DbAccess.UpdateDataTable(sqlTable, sqlAdapter, objTrans)

            Select Case ProcessType
                Case cst_PlanUpdate, cst_Update '計畫轉入成功 (修改)
                    If sqldr_pp IsNot Nothing Then
                        'UPDATE PLAN_PLANINFO
                        DbAccess.UpdateDataTable(dtPlanInfo, daPlanInfo, objTrans)
                    End If
            End Select

            '課程階段
            'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
            'Call SAVE_CLASS_CLASSLEVEL(objTrans)

            '假如有學員的話,要更新學員學號 Strat
            If OldClassID.Value <> TBclass_id.Text AndAlso rqOCID <> "" Then
                Dim da As SqlDataAdapter = Nothing
                Dim sqlCS1 As String = " SELECT SOCID,STUDENTID FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & rqOCID & "'"
                Dim dt As DataTable = DbAccess.GetDataTable(sqlCS1, da, objTrans)
                If dt.Rows.Count <> 0 Then
                    For Each dr As DataRow In dt.Rows
                        dr("StudentID") = Replace(dr("StudentID"), OldClassID.Value, TBclass_id.Text)
                    Next
                    DbAccess.UpdateDataTable(dt, da, objTrans)
                End If
            End If
            '假如有學員的話,要更新學員學號 End
            DbAccess.CommitTrans(objTrans)

            flag_save_ok = True
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            flag_save_ok = False
            '偵錯用儲存欄
            Dim strErrmsg As String = ""
            strErrmsg &= "/* ex.ToString: */" & vbCrLf & ex.ToString & vbCrLf
            strErrmsg &= "/* str_save_data_All: */" & vbCrLf & str_save_data_All & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Throw ex
        End Try

        If flag_save_ok Then
            '97年 TIMS 分署(中心)自辦計劃加入此規則，將自行增加賦予權限給系統管理者及承辦人 START 
            'Dim PlanKind As String
            Dim sPlanKind As String = TIMS.Get_PlanKind(Me, objconn)
            'sql=" SELECT PLANKIND FROM ID_PLAN WHERE PlanID='" & sm.UserInfo.PlanID & "' "
            'PlanKind=DbAccess.ExecuteScalar(sql, objconn)
            sql = " SELECT OCID FROM CLASS_CLASSINFO WHERE PlanID='" & PlanIDValue.Value & "' AND ComIDNO='" & P_ComIDNO.Value & "' AND SeqNO='" & P_SeqNO.Value & "' "
            rqOCID = DbAccess.ExecuteScalar(sql, objconn)
            If sPlanKind = "1" Then
                Select Case ProcessType
                    Case cst_PlanUpdate
                        If rqOCID <> "" Then TIMS.Insert_Auth_AccRWClass(sm, rqOCID, -1, objconn)
                End Select
            End If
        End If

        If Session(cst_ClassSearchStr) Is Nothing AndAlso ViewState(cst_ClassSearchStr) IsNot Nothing Then
            Session(cst_ClassSearchStr) = ViewState(cst_ClassSearchStr)
        End If

        'Dim strScript As String=""
        strScript = ""
        strScript = "<script language=""javascript"">" + vbCrLf
        Select Case ProcessType
            Case cst_PlanUpdate
                'Dim check_nextclass As String="SELECT OCID FROM CLASS_CLASSINFO WHERE OCID > '" & rqOCID & "' and CLSID='" & clsid.Value & "'  and (PlanID='" & PlanIDValue.Value & "' and CyclType='" & TB_CyclType.Text & "'  and RID='" & RIDValue.Value & "')"
                If iCheck_Nextclass > 0 Then
                    strScript &= "if(window.confirm('目前尚有未開班完成的資料,你確認要離開嗎?')){" + vbCrLf
                    strScript &= "alert('未完成開班資料的班級,將會造成後續學員動態管理查不到該班資料,\n請務必回到本作業,完成開班作業程序.');" + vbCrLf
                    strScript &= "location.href='TC_01_004.aspx?ID='+document.getElementById('Re_ID').value;}" + vbCrLf
                Else
                    strScript &= "alert('班級轉入成功');" + vbCrLf
                    strScript &= "location.href='TC_01_004.aspx?ID='+document.getElementById('Re_ID').value;" + vbCrLf
                End If
            Case cst_Update
                'Session("ClassSearchStr")=ViewState("ClassSearchStr")
                strScript &= "alert('開班資料修改成功!!');" + vbCrLf
                strScript &= "location.href='TC_01_004.aspx?ID='+document.getElementById('Re_ID').value;" + vbCrLf
        End Select
        strScript &= "</script>"
        Page.RegisterStartupScript(TIMS.xBlockName(), strScript)

        'Dim StrLevel As String
        'If iOCID_New <> 0 Then
        '    StrLevel="SELECT * FROM CLASS_CLASSLEVEL WHERE OCID='" & iOCID_New & "'"
        '    Session("ClassLevel")=DbAccess.GetDataTable(StrLevel, objconn)
        'End If
        'If rqOCID <> "" Then
        '    StrLevel="SELECT * FROM CLASS_CLASSLEVEL WHERE OCID='" & rqOCID & "'"
        '    Session("ClassLevel")=DbAccess.GetDataTable(StrLevel, objconn)
        'End If
    End Sub

    '儲存
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        Call SaveData1()
    End Sub

    '維護下一班
    Private Sub next_class_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles next_class.Click
        Dim strScript1 As String
        rqOCID = TIMS.ClearSQM(rqOCID)
        clsid.Value = TIMS.ClearSQM(clsid.Value)
        PlanIDValue.Value = TIMS.ClearSQM(PlanIDValue.Value)
        TB_CyclType.Text = TIMS.FmtCyclType(TB_CyclType.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim check_nextclass As String
        check_nextclass = ""
        check_nextclass &= " SELECT OCID FROM CLASS_CLASSINFO WHERE 1=1"
        check_nextclass &= " and OCID>'" & rqOCID & "'"
        check_nextclass &= " and CLSID='" & clsid.Value & "'"
        check_nextclass &= " and (PlanID='" & PlanIDValue.Value & "' and ISNULL(CyclType,'')='" & TB_CyclType.Text & "' and RID='" & RIDValue.Value & "')"
        check_nextclass &= " ORDER BY OCID"
        Dim next_ocid As String = Convert.ToString(DbAccess.ExecuteScalar(check_nextclass, objconn))
        'Call bt_save_Click(sender, e)
        Call SaveData1()
        If next_ocid = "" Then
            'next_class.Enabled=False
            strScript1 = "<script language=""javascript"">" + vbCrLf
            strScript1 &= "alert('已為最後一筆班級資料!!!');" + vbCrLf
            strScript1 &= "</script>"
            Page.RegisterStartupScript("", strScript1)
            Return 'Exit Sub
        End If

        'Response.Redirect("TC_01_004_add.aspx?ocid=" & next_ocid & "&ProcessType=PlanUpdate&ID=" & Re_ID.Value & "")
        Dim url1 As String = "TC_01_004_add.aspx?ocid=" & next_ocid & "&ProcessType=PlanUpdate&ID=" & Re_ID.Value & ""
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '回上一頁
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Session(cst_ClassSearchStr) Is Nothing AndAlso Not ViewState(cst_ClassSearchStr) Is Nothing Then Session(cst_ClassSearchStr) = ViewState(cst_ClassSearchStr)

        'Session("ClassSearchStr")=ViewState("ClassSearchStr")
        'If Session("ClassSearchStr") Is Nothing AndAlso Not ViewState("ClassSearchStr") Is Nothing Then
        '    Session("ClassSearchStr")=ViewState("ClassSearchStr")
        'End If
        'Response.Redirect("TC_01_004.aspx?ID=" & Request("ID") & "")
        Dim url1 As String = "TC_01_004.aspx?ID=" & Request("ID") & ""
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

#Region "NO USE"

    ' <summary>SHOW CLASS_CLASSLEVEL '課程階段</summary>
    'Sub SHOW_CLASSLEVEL()
    '    'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
    '    '課程階段
    '    If Not tb_CLASSLEVEL.Visible Then Return

    '    Dim StrLevel As String=""
    '    '建立課程階段的DATATABLE
    '    If rqOCID <> "" AndAlso rqOCID <> "0" Then
    '        StrLevel="SELECT * FROM CLASS_CLASSLEVEL WHERE OCID='" & rqOCID & "'"
    '    Else
    '        StrLevel="SELECT * FROM CLASS_CLASSLEVEL WHERE 1<>1"
    '    End If
    '    Dim dtlevel As DataTable=Nothing
    '    dtlevel=DbAccess.GetDataTable(StrLevel, objconn)
    '    dtlevel.Columns("CCLID").AutoIncrement=True
    '    dtlevel.Columns("CCLID").AutoIncrementSeed=-1
    '    dtlevel.Columns("CCLID").AutoIncrementStep=-1
    '    Session("ClassLevel")=dtlevel
    '    DG_ClassLevel.Visible=False
    '    If dtlevel.Rows.Count=0 Then Return

    '    DG_ClassLevel.Visible=True
    '    DG_ClassLevel.DataSource=dtlevel
    '    DG_ClassLevel.DataKeyField="CCLID"
    '    DG_ClassLevel.DataBind()
    'End Sub

    '課程階段
    'Sub SAVE_CLASS_CLASSLEVEL(ByRef objTrans As SqlTransaction)
    '    'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
    '    '課程階段
    '    If Not tb_CLASSLEVEL.Visible Then Return '課程階段

    '    If Session("ClassLevel") Is Nothing Then Return '課程階段
    '    Dim dtlevel As DataTable=Session("ClassLevel")
    '    If dtlevel Is Nothing Then Return
    '    If dtlevel.Rows.Count=0 Then Return

    '    Dim sql As String=""
    '    'update 開班階段檔(Class_ClassLevel)
    '    Dim dtTemp As New DataTable
    '    'Session("ClassLevel")
    '    '課程階段

    '    If Session("ClassLevel") IsNot Nothing Then dtlevel=Session("ClassLevel")
    '    If dtlevel IsNot Nothing Then '不為空時,才新增資料
    '        'dtlevel=Session("ClassLevel")
    '        If iOCID_New <> 0 Then
    '            Dim strClass As String=""
    '            strClass=" DELETE CLASS_CLASSLEVEL WHERE OCID=" & iOCID_New
    '            DbAccess.ExecuteNonQuery(strClass, objTrans)
    '            sql="" & vbCrLf
    '            sql &= " INSERT INTO CLASS_CLASSLEVEL ( " & vbCrLf
    '            sql &= " CCLID,OCID ,LEVELNAME ,LEVELSDATE ,LEVELEDATE ,LEVELHOUR ,MODIFYACCT ,MODIFYDATE " & vbCrLf
    '            'sql &= "  ,NUM ,LSDATE ,LEDATE " & vbCrLf
    '            sql &= " ) VALUES ( " & vbCrLf
    '            sql &= " @CCLID,@OCID ,@LEVELNAME ,@LEVELSDATE ,@LEVELEDATE ,@LEVELHOUR ,@MODIFYACCT ,GETDATE() " & vbCrLf
    '            'sql &= "  ,NUM ,LSDATE ,LEDATE " & vbCrLf
    '            sql &= " )" & vbCrLf
    '            Dim iCmd As New SqlCommand(sql, objconn, objTrans)
    '            For Each eItem As DataGridItem In DG_ClassLevel.Items
    '                Dim HidCCLID As HiddenField=eItem.FindControl("HidCCLID")
    '                Dim HidLevelName As HiddenField=eItem.FindControl("HidLevelName")
    '                Dim LevelName As Label=eItem.FindControl("LevelName")
    '                Dim LevelSDate As Label=eItem.FindControl("LevelSDate")
    '                Dim LevelEDate As Label=eItem.FindControl("LevelEDate")
    '                Dim LevelHour As Label=eItem.FindControl("LevelHour")
    '                'Dim btnDel As LinkButton=eItem.FindControl("btnDel")
    '                'Dim iCCLID As Integer=0
    '                If Val(HidCCLID.Value) <= 0 Then HidCCLID.Value=DbAccess.GetNewId(objTrans, "CLASS_CLASSLEVEL_CCLID_SEQ,CLASS_CLASSLEVEL,CCLID")
    '                With iCmd
    '                    .Parameters.Clear()
    '                    .Parameters.Add("CCLID", SqlDbType.Int).Value=Val(HidCCLID.Value)
    '                    .Parameters.Add("OCID", SqlDbType.Int).Value=Val(iOCID_New)
    '                    .Parameters.Add("LEVELNAME", SqlDbType.VarChar).Value=TIMS.ClearSQM(HidLevelName.Value)
    '                    .Parameters.Add("LEVELSDATE", SqlDbType.DateTime).Value=If(LevelSDate.Text <> "", CDate(TIMS.cdate2(LevelSDate.Text)), Convert.DBNull)
    '                    .Parameters.Add("LEVELEDATE", SqlDbType.DateTime).Value=If(LevelEDate.Text <> "", CDate(TIMS.cdate2(LevelEDate.Text)), Convert.DBNull)
    '                    .Parameters.Add("LEVELHOUR", SqlDbType.Int).Value=Val(LevelHour.Text)
    '                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value=sm.UserInfo.UserID
    '                    .ExecuteNonQuery()
    '                    'DbAccess.ExecuteNonQuery(iCmd.CommandText, objTrans, iCmd.Parameters)  'edit，by:20181024
    '                End With
    '            Next
    '        End If

    '        If rqOCID <> "" AndAlso rqOCID <> "0" Then
    '            Dim strClass As String=""
    '            strClass=" DELETE CLASS_CLASSLEVEL WHERE OCID=" & rqOCID
    '            DbAccess.ExecuteNonQuery(strClass, objTrans)

    '            sql="" & vbCrLf
    '            sql &= " INSERT INTO CLASS_CLASSLEVEL ( " & vbCrLf
    '            sql &= " CCLID,OCID ,LEVELNAME ,LEVELSDATE ,LEVELEDATE ,LEVELHOUR ,MODIFYACCT ,MODIFYDATE " & vbCrLf
    '            'sql &= "   ,NUM ,LSDATE ,LEDATE " & vbCrLf
    '            sql &= " ) VALUES ( " & vbCrLf
    '            sql &= " @CCLID ,@OCID ,@LEVELNAME ,@LEVELSDATE ,@LEVELEDATE ,@LEVELHOUR ,@MODIFYACCT ,GETDATE() " & vbCrLf
    '            'sql &= "   ,NUM ,LSDATE ,LEDATE " & vbCrLf
    '            sql &= " )" & vbCrLf
    '            Dim iCmd As New SqlCommand(sql, objconn, objTrans)
    '            For Each eItem As DataGridItem In DG_ClassLevel.Items
    '                Dim HidCCLID As HiddenField=eItem.FindControl("HidCCLID")
    '                Dim HidLevelName As HiddenField=eItem.FindControl("HidLevelName")
    '                Dim LevelName As Label=eItem.FindControl("LevelName")
    '                Dim LevelSDate As Label=eItem.FindControl("LevelSDate")
    '                Dim LevelEDate As Label=eItem.FindControl("LevelEDate")
    '                Dim LevelHour As Label=eItem.FindControl("LevelHour")
    '                'Dim btnDel As LinkButton=eItem.FindControl("btnDel")
    '                'Dim iCCLID As Integer=0
    '                If Val(HidCCLID.Value) <= 0 Then HidCCLID.Value=DbAccess.GetNewId(objTrans, "CLASS_CLASSLEVEL_CCLID_SEQ,CLASS_CLASSLEVEL,CCLID")
    '                With iCmd
    '                    .Parameters.Clear()
    '                    .Parameters.Add("CCLID", SqlDbType.Int).Value=Val(HidCCLID.Value)
    '                    .Parameters.Add("OCID", SqlDbType.Int).Value=Val(rqOCID)
    '                    .Parameters.Add("LEVELNAME", SqlDbType.VarChar).Value=TIMS.ClearSQM(HidLevelName.Value)
    '                    .Parameters.Add("LEVELSDATE", SqlDbType.DateTime).Value=If(LevelSDate.Text <> "", CDate(TIMS.cdate2(LevelSDate.Text)), Convert.DBNull)
    '                    .Parameters.Add("LEVELEDATE", SqlDbType.DateTime).Value=If(LevelEDate.Text <> "", CDate(TIMS.cdate2(LevelEDate.Text)), Convert.DBNull)
    '                    .Parameters.Add("LEVELHOUR", SqlDbType.Int).Value=Val(LevelHour.Text)
    '                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value=sm.UserInfo.UserID
    '                    .ExecuteNonQuery()
    '                    'DbAccess.ExecuteNonQuery(iCmd.CommandText, objTrans, iCmd.Parameters)  'edit，by:20181024
    '                End With
    '            Next
    '        End If
    '    End If
    'End Sub


    '課程階段選擇
    'Sub TB_LevelType_Selected1()
    '    'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
    '    '課程階段
    '    If Not tb_CLASSLEVEL.Visible Then Return '課程階段

    '    LevelName.Items.Clear()
    '    LevelName.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    '    Select Case TB_LevelType.SelectedValue '0~4
    '        Case "", "0"
    '            '選"無"階段時,不需輸入階段資料
    '            LevelName.Enabled=False
    '            LevelSDate.Enabled=False
    '            LevelEDate.Enabled=False
    '            LevelHour.Enabled=False
    '            add_but.Enabled=False
    '            Common.SetListItem(LevelName, "")
    '            LevelSDate.Text=""
    '            LevelEDate.Text=""
    '            LevelHour.Text=""
    '        Case "1", "2", "3", "4"
    '            LevelName.Enabled=True
    '            LevelSDate.Enabled=True
    '            LevelEDate.Enabled=True
    '            LevelHour.Enabled=True
    '            add_but.Enabled=True
    '            Dim i As Integer=1
    '            Dim str As String=""
    '            Do While i <= CInt(TB_LevelType.SelectedValue)
    '                Select Case i
    '                    Case 1
    '                        str="一"
    '                    Case 2
    '                        str="二"
    '                    Case 3
    '                        str="三"
    '                    Case 4
    '                        str="四"
    '                    Case Else
    '                        Exit Do
    '                End Select
    '                LevelName.Items.Add(New ListItem(str, "0" & i))
    '                i=i + 1
    '            Loop
    '            If TB_LevelType.SelectedValue="1" Then
    '                'LevelName.SelectedValue="01"
    '                Common.SetListItem(LevelName, "01")
    '                LevelSDate.Text=TB_STDate.Text
    '                LevelEDate.Text=TB_FTDate.Text
    '                LevelHour.Text=TB_THours.Text
    '            End If
    '    End Select
    '    If IsPostBack Then Page.RegisterStartupScript("1111", "<script>window.scroll(0,document.body.scrollHeight)</script>")
    'End Sub

    '課程階段
    'Private Sub TB_LevelType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_LevelType.SelectedIndexChanged
    '    Call TB_LevelType_Selected1()
    'End Sub

    '新增 (階段資料)
    'Private Sub add_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles add_but.Click
    '    'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
    '    '課程階段
    '    If Not tb_CLASSLEVEL.Visible Then Return 'SHOW CLASS_CLASSLEVEL '課程階段

    '    'Dim strScript, sqlstr As String
    '    'Dim dt As DataTable
    '    'Dim dr As DataRow
    '    Dim intTotalHour As Integer=0
    '    DG_ClassLevel.Visible=True
    '    Dim dt As DataTable
    '    If Session("ClassLevel") Is Nothing Then
    '        Dim sqlstr As String=""
    '        sqlstr="SELECT * FROM CLASS_CLASSLEVEL WHERE 1<>1"
    '        dt=DbAccess.GetDataTable(sqlstr, objconn)
    '        dt.Columns("CCLID").AutoIncrement=True
    '        dt.Columns("CCLID").AutoIncrementSeed=-1
    '        dt.Columns("CCLID").AutoIncrementStep=-1
    '        Session("ClassLevel")=dt
    '        intTotalHour=0
    '    Else
    '        dt=Session("ClassLevel")
    '        dt.Columns("CCLID").AutoIncrement=True
    '        dt.Columns("CCLID").AutoIncrementSeed=-1
    '        dt.Columns("CCLID").AutoIncrementStep=-1
    '        For Each dr1 As DataRow In dt.Rows
    '            If dr1.RowState <> DataRowState.Deleted Then intTotalHour += dr1("LevelHour")
    '        Next
    '    End If

    '    'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
    '    '課程階段
    '    'If tb_CLASSLEVEL.Visible Then
    '    '    If dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows).Length=TB_LevelType.SelectedValue Then  '當輸入筆數與課程階段相同時,則不可在新增
    '    '        Dim strScript1 As String
    '    '        strScript1="<script language=""javascript"">" + vbCrLf
    '    '        strScript1 &= "alert('階段資料設定筆數已滿!!');" + vbCrLf
    '    '        strScript1 &= "</script>"
    '    '        Page.RegisterStartupScript("", strScript1)
    '    '        LevelSDate.Text=""
    '    '        LevelEDate.Text=""
    '    '        LevelHour.Text=""
    '    '        Return 'Exit Sub
    '    '    End If
    '    'End If

    '    If CDate(LevelSDate.Text) < CDate(TB_STDate.Text) OrElse CDate(LevelEDate.Text) > CDate(TB_FTDate.Text) Then
    '        Common.MessageBox(Page, "階段起迄日不可超過班級開結訓日!!")
    '        Dim strScript As String=""
    '        strScript="<script language=""javascript"">" + vbCrLf
    '        strScript &= "</script>"
    '        Page.RegisterStartupScript("window_onload", strScript)
    '        Return 'Exit Sub
    '    End If

    '    If LevelHour.Text <> "" Then
    '        intTotalHour += LevelHour.Text
    '        If intTotalHour > Convert.ToInt16(TB_THours.Text) Then
    '            Common.MessageBox(Page, "階段時數加總不可超過班級訓練時數!!")
    '            Dim strScript As String=""
    '            strScript="<script language=""javascript"">" + vbCrLf
    '            strScript &= "</script>"
    '            Page.RegisterStartupScript("window_onload", strScript)
    '            Return 'Exit Sub
    '        End If
    '        If Convert.ToInt16(LevelHour.Text) > Convert.ToInt16(TB_THours.Text) Then
    '            Common.MessageBox(Page, "階段時數不可超過班級訓練時數!!")
    '            Dim strScript As String=""
    '            strScript="<script language=""javascript"">" + vbCrLf
    '            strScript &= "</script>"
    '            Page.RegisterStartupScript("window_onload", strScript)
    '            Return 'Exit Sub
    '        End If
    '    End If

    '    For Each dr1 As DataRow In dt.Rows
    '        If dr1.RowState <> DataRowState.Deleted Then
    '            If CDate(dr1("LevelSDate"))=CDate(LevelSDate.Text) Or CDate(dr1("LevelEDate"))=CDate(LevelEDate.Text) Or dr1("LevelName")=LevelName.SelectedValue Then
    '                Dim strScript1 As String
    '                strScript1="<script language=""javascript"">" + vbCrLf
    '                strScript1 &= "alert('階段資料設定重複!!!!');" + vbCrLf
    '                strScript1 &= "</script>"
    '                Page.RegisterStartupScript("", strScript1)
    '                Return 'Exit Sub
    '            End If
    '        End If
    '    Next

    '    Dim flag As Boolean=True
    '    For Each dr As DataRow In dt.Rows '判斷輸入的日期區間是否正確
    '        If dr.RowState <> DataRowState.Deleted Then
    '            If CDate(dr("LevelSDate")) <= CDate(LevelSDate.Text) And CDate(dr("LevelEDate")) >= CDate(LevelSDate.Text) Then flag=False
    '            If CDate(dr("LevelSDate")) <= CDate(LevelEDate.Text) And CDate(dr("LevelEDate")) >= CDate(LevelEDate.Text) Then flag=False
    '        End If
    '    Next
    '    If flag=False Then
    '        Dim strScript1 As String
    '        strScript1="<script language=""javascript"">" + vbCrLf
    '        strScript1 &= "alert('階段日期輸入範圍錯誤!!!');" + vbCrLf
    '        strScript1 &= "</script>"
    '        Page.RegisterStartupScript("", strScript1)
    '        Return 'Exit Sub
    '    End If
    '    'CLASS_CLASSLEVEL
    '    Dim dr2 As DataRow=dt.NewRow
    '    dt.Rows.Add(dr2)
    '    dr2("LevelSDate")=TIMS.cdate2(LevelSDate.Text)
    '    dr2("LevelEDate")=TIMS.cdate2(LevelEDate.Text)
    '    dr2("LevelName")=If(LevelName.SelectedValue <> "", LevelName.SelectedValue, "1")
    '    dr2("LevelHour")=If(LevelHour.Text <> "", Val(LevelHour.Text), 0)
    '    dr2("ModifyAcct")=sm.UserInfo.UserID
    '    dr2("ModifyDate")=Now()
    '    Session("ClassLevel")=dt
    '    DG_ClassLevel.DataSource=dt
    '    DG_ClassLevel.DataKeyField="CCLID"
    '    DG_ClassLevel.DataBind()

    '    '清空
    '    LevelSDate.Text=""
    '    LevelEDate.Text=""
    '    LevelHour.Text=""

    '    If IsPostBack Then Page.RegisterStartupScript("1111", "<script>window.scroll(0,document.body.scrollHeight)</script>")
    'End Sub


    'Private Sub DG_ClassLevel_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_ClassLevel.ItemCommand
    '    'OJT-21032503：在職進修訓練、接受企業委託訓練 - 開班資料查詢：部分欄位隱藏 by AMU 20210615
    '    '課程階段
    '    If Not tb_CLASSLEVEL.Visible Then Return 'SHOW CLASS_CLASSLEVEL '課程階段
    '    If Session("ClassLevel") Is Nothing Then Return 'Exit Sub 'SHOW CLASS_CLASSLEVEL '課程階段

    '    Dim dt As DataTable=Session("ClassLevel")
    '    Select Case e.CommandName
    '        Case "Del"
    '            Dim sCmdArg As String=Convert.ToString(e.CommandArgument)
    '            Dim CCLID As String=TIMS.GetMyValue(sCmdArg, "CCLID")
    '            ff3="CCLID=" & CCLID
    '            If dt.Select(ff3).Length=0 Then Return 'Exit Sub
    '            Dim dr As DataRow=dt.Select(ff3)(0)
    '            dr.Delete()
    '            'ComClass.Alert(Page, "刪除成功")
    '            'hidpyt_id.Value=""
    '            Session("ClassLevel")=dt
    '            'Call showPLAN_YOUNG_TRAIN()
    '            DG_ClassLevel.DataSource=dt
    '            DG_ClassLevel.DataKeyField="CCLID"
    '            DG_ClassLevel.DataBind()
    '    End Select
    'End Sub

    'Private Sub DG_ClassLevel_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_ClassLevel.ItemDataBound
    '    Select Case e.Item.ItemType
    '        Case ListItemType.AlternatingItem, ListItemType.Item
    '            Dim drv As DataRowView=e.Item.DataItem
    '            Dim HidCCLID As HiddenField=e.Item.FindControl("HidCCLID")
    '            Dim HidLevelName As HiddenField=e.Item.FindControl("HidLevelName")
    '            Dim LevelName As Label=e.Item.FindControl("LevelName")
    '            Dim LevelSDate As Label=e.Item.FindControl("LevelSDate")
    '            Dim LevelEDate As Label=e.Item.FindControl("LevelEDate")
    '            Dim LevelHour As Label=e.Item.FindControl("LevelHour")
    '            Dim btnDel As LinkButton=e.Item.FindControl("btnDel")
    '            HidCCLID.Value=Convert.ToString(drv("CCLID"))
    '            HidLevelName.Value=Convert.ToString(drv("LevelName"))
    '            If HidLevelName.Value <> "" Then
    '                'LevelName.Text=Convert.ToString(drv("LevelName"))
    '                LevelName.Text=TIMS.ChangeNum(Val(HidLevelName.Value))
    '            End If
    '            If Convert.ToString(drv("LevelSDate")) <> "" Then LevelSDate.Text=TIMS.cdate3(drv("LevelSDate"))
    '            If Convert.ToString(drv("LevelEDate")) <> "" Then LevelEDate.Text=TIMS.cdate3(drv("LevelEDate"))
    '            LevelHour.Text=Convert.ToString(drv("LevelHour"))
    '            Dim sCmdArg As String=""
    '            TIMS.SetMyValue(sCmdArg, "CCLID", Convert.ToString(drv("CCLID")))
    '            btnDel.CommandArgument=sCmdArg
    '            btnDel.Attributes.Add("onclick", "return confirm('確定刪除該筆資料?');")
    '    End Select
    'End Sub

#End Region

End Class
