Partial Class TC_01_027_add
    Inherits AuthBasePage

#Region "WEBFREM"

    Sub sUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("TEACH_TEACHERINFO", objconn)
        Call TIMS.sUtl_SetMaxLen(dt, "TEACHERID", TeacherID)
        Call TIMS.sUtl_SetMaxLen(dt, "MOBILE", Mobile)
        Call TIMS.sUtl_SetMaxLen(dt, "TEACHCNAME", TeachCName)
        Call TIMS.sUtl_SetMaxLen(dt, "TEACHENAME", TeachEName)
        Call TIMS.sUtl_SetMaxLen(dt, "IDNO", IDNO)
        Call TIMS.sUtl_SetMaxLen(dt, "SCHOOLNAME", SchoolName)
        Call TIMS.sUtl_SetMaxLen(dt, "DEPARTMENT", Department)
        Call TIMS.sUtl_SetMaxLen(dt, "PHONE", Phone)
        Call TIMS.sUtl_SetMaxLen(dt, "ADDRESS", Address)
        Call TIMS.sUtl_SetMaxLen(dt, "WORKORG", WorkOrg)
        Call TIMS.sUtl_SetMaxLen(dt, "INVEST", Invest2) '50 NVARCHAR2
        Call TIMS.sUtl_SetMaxLen(dt, "WORKADDR", WorkAddr)
        Call TIMS.sUtl_SetMaxLen(dt, "WORKPHONE", WorkPhone)
        Call TIMS.sUtl_SetMaxLen(dt, "EXPUNIT1", ExpUnit1)
        Call TIMS.sUtl_SetMaxLen(dt, "EXPUNIT2", ExpUnit2)
        Call TIMS.sUtl_SetMaxLen(dt, "EXPUNIT3", ExpUnit3)
        Call TIMS.sUtl_SetMaxLen(dt, "INV1", tINV1)
        Call TIMS.sUtl_SetMaxLen(dt, "INV2", tINV2)
        Call TIMS.sUtl_SetMaxLen(dt, "INV3", tINV3)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY1", Specialty1)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY2", Specialty2)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY3", Specialty3)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY4", Specialty4)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY5", Specialty5)
        Call TIMS.sUtl_SetMaxLen(dt, "EMAIL", Email)
        Call TIMS.sUtl_SetMaxLen(dt, "SERVDEPT", ServDept)
        Call TIMS.sUtl_SetMaxLen(dt, "FAX", Fax)
        Call TIMS.sUtl_SetMaxLen(dt, "TRANSBOOK", TransBook)
        Call TIMS.sUtl_SetMaxLen(dt, "PROLICENSE1", ProLicense1)
        Call TIMS.sUtl_SetMaxLen(dt, "PROLICENSE2", ProLicense2)
    End Sub

#End Region

    Const cst_rqProe_Insert As String = "Insert"
    Const cst_rqProe_add As String = "add"
    Const cst_rqProe_edit As String = "edit"
    Const Cst_ProeTxt_新增 As String = "-新增"
    Const Cst_ProeTxt_修改 As String = "-修改"

    '130是講師 2010/12/13 號改成委訓單位都可以看到師資別,但是不可以選,
    Const cst_KindID_130 As String = "130"
    Const cst_KindID_130_講師_Txt As String = "講師"
    Const cst_KindEngage_1_內聘 As String = "1"
    Const cst_KindEngage_2_外聘 As String = "2"

    Const Cst_msg1 As String = "建議使用 區碼-電話號碼"
    Const Cst_msg2 As String = "同一計劃，有相同身分證號碼，重複輸入" '依RIDValue.Value
    Const Cst_msg3 As String = "\n講師代碼發生重複，請注意"
    Const Cst_msg4 As String = "講師代碼最多10 碼(英文 / 數字)"
    Const Cst_msg5 As String = "資料新增成功!"
    Const Cst_msg6 As String = "資料修改成功!"
    Const Cst_msg7 As String = "資料儲存有誤，請再確認輸入資料正確性!!!"
    Const Cst_msg8 As String = "類別，請選擇講師或助教!!"

    Dim ff3 As String = ""
    Dim dtGOVCLASSCAST3 As DataTable = Nothing '主要職類table y2018
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        Call sUtl_PageInit1()
        '檢查Session是否存在 End

        Dim rqProecess As String = TIMS.ClearSQM(Request("proecess"))
        '按新增時代查詢之 教師姓名 & 身分證號碼
        Dim rqTechID As String = TIMS.ClearSQM(Request("serial"))
        Dim rqTeachCName As String = TIMS.ClearSQM(Request("TeachCName"))
        Dim rqTeachIDNO As String = TIMS.ClearSQM(Request("TeachIDNO"))

        dtGOVCLASSCAST3 = TIMS.Get_GOVCLASSCAST3dt(dtGOVCLASSCAST3, objconn)

        If Not IsPostBack Then
            msg.Text = ""
            '保存查詢值----------------------------------------Start
            ViewState("MySearchStr") = ""
            If Session("MySearchStr") IsNot Nothing Then ViewState("MySearchStr") = Session("MySearchStr")
            '保存查詢值----------------------------------------End
            TPlanID.Value = sm.UserInfo.TPlanID '登入-計畫
            TBplanOrgName.Text = sm.UserInfo.OrgName '登入/選擇
            RIDValue.Value = sm.UserInfo.RID '登入/選擇
            LID.Value = sm.UserInfo.LID '登入/選擇

            Call AddItem()

            Dim flag_vs_MySearchStr_CanUse As Boolean = False
            If ViewState("MySearchStr") Is Nothing Then flag_vs_MySearchStr_CanUse = False
            If flag_vs_MySearchStr_CanUse AndAlso ViewState("MySearchStr") <> "" Then
                Dim s_MySearchStr As String = Convert.ToString(ViewState("MySearchStr"))
                TBplanOrgName.Text = TIMS.GetMyValue(s_MySearchStr, "center")
                RIDValue.Value = TIMS.GetMyValue(s_MySearchStr, "RIDValue")
            End If

            Proecess.Text = Cst_ProeTxt_新增
            Select Case rqProecess
                Case cst_rqProe_add, cst_rqProe_Insert
                    Proecess.Text = Cst_ProeTxt_新增
                    '按新增時代查詢之 教師姓名 & 身分證號碼
                    TeachCName.Text = rqTeachCName
                    IDNO.Text = rqTeachIDNO
                Case Else
                    Proecess.Text = Cst_ProeTxt_修改
                    'Not IsPostBack
                    Call sCreate1(rqTechID)
                    '假如不同單位的師資資料則不允許修改
                    If Len(sm.UserInfo.RID.ToString) > 1 And RIDValue.Value <> sm.UserInfo.RID Then Button1.Visible = False
            End Select

            'LID【0=>署(局), 1=>分署(中心), 2=>委訓】
            If LID.Value = 2 Then
                'KindID_TD1.Style.Item("display") = "none" 
                'KindID_TD2.Style.Item("display") = "none"
                '原本委訓單位沒有師資別，後來改成有師資別可以看但是不能選,委訓單位的師資別統一只能是講師2010/12/13
                KindID.Items.Clear()
                KindID.Items.Insert(0, New ListItem(cst_KindID_130_講師_Txt, cst_KindID_130))
                'KindID.SelectedValue = 130
                Common.SetListItem(KindID, cst_KindID_130)
                '登入者為委訓單位鎖定 (FALSE)
                KindID.Enabled = Not (sm.UserInfo.LID = 2)
                '(不可作用，說明原因)
                If Not KindID.Enabled Then TIMS.Tooltip(KindID, "委訓單位的師資別只能是講師")
                '依 KindID.Enabled 若可作用，則啟用 AutoPostBack
                KindEngage.AutoPostBack = (KindID.Enabled) '可變

                IVID.Visible = False
                R1.Visible = False '不顯示
                R2.Visible = False
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '企訓專用
                    Invest1_TD1.Visible = False
                    Invest1_TD2.Visible = False
                    Invest2_TD1.Visible = True
                    Invest2_TD2.Visible = True
                    Invest1.Visible = False '產業人才投資方案
                    star1.Visible = True
                    star2.Visible = True
                    star3.Visible = True
                Else
                    Invest1_TD1.Visible = True
                    Invest1_TD2.Visible = True
                    Invest2_TD1.Visible = False
                    Invest2_TD2.Visible = False
                    Invest1.Visible = True '非產業人才投資方案
                    star1.Visible = False
                    star2.Visible = False
                    star3.Visible = False
                End If
            Else
                IVID.Visible = True
                Invest1.Visible = False
                Invest2_TD1.Visible = False
                Invest2_TD2.Visible = False
                R1.Visible = False
                R2.Visible = False
                'KindEngage.AutoPostBack = True
            End If
        End If

        Dim v_KindEngage As String = TIMS.GetListValue(KindEngage)
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '不是產業人才投資方案也不是委訓單位
            If LID.Value <> 2 Then
                If v_KindEngage = "2" Then '如果是外聘
                    IVID.Visible = True '選擇職稱
                    Invest1.Visible = True '輸入職稱文字
                    R1.Visible = True '顯示
                    R2.Visible = True '顯示
                Else '內聘或是未選擇時
                    IVID.Visible = True
                    Invest1.Visible = False
                    R1.Visible = False
                    R2.Visible = False
                End If
            End If
        End If

        '依身分證號載入舊資料(非新增功能)(不顯示)
        Button3.Visible = False
        Select Case rqProecess
            Case cst_rqProe_add, cst_rqProe_Insert
                Button3.Visible = True '依身分證號載入舊資料(顯示)
                If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If LID.Value <> 2 Then   '不是產業人才投資方案也不是委訓單位
                        R1.Checked = True '職稱 (新增預設值)
                        Invest1.Enabled = False '輸入職稱文字(停用)  (新增預設值)
                    End If
                End If
        End Select

        Button4.Style("display") = "none"

        TIMS.Tooltip(Phone, Cst_msg1)
        TIMS.Tooltip(WorkPhone, Cst_msg1)

        '郵遞區號查詢
        Litcity_code.Text = TIMS.Get_WorkZIPB3Link2()
        Litcity_code1.Text = TIMS.Get_WorkZIPB3Link2()

        city_code.Attributes.Add("onblur", "getZipName('TBCity',this,this.value);")
        city_code1.Attributes.Add("onblur", "getZipName('TBCity1',this,this.value);")

        Dim bt1_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code, AddressZIPB3, hidAddressZIP6W, TBCity, Address)
        bt_openZip1.Attributes.Add("onclick", bt1_Attr_VAL)
        Dim bt2_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code1, WorkZIPB3, hidWorkZIP6W, TBCity1, WorkAddr)
        bt_openZip2.Attributes.Add("onclick", bt2_Attr_VAL)

        Button1.Attributes("onclick") = "javascript:return chkdata();"
        Button3.Attributes("onclick") = "wopen('TC_01_027_copy.aspx?IDNO='+document.getElementById('IDNO').value,'copy',450,400,1);"
        IDNO.Attributes("onBlur") = "return SexChoice();"
        Sex.Attributes("onclick") = "return SexChoice();"
        KindID.Attributes("onclick") = "return KindIDChoice();"
        R2.Attributes("onclick") = "return change();"
        R1.Attributes("onclick") = "return change2();"

        If Not IsPostBack Then
            '有可能重複2次，所以加入判斷是否為空值
            If ProLicense1.Text = "" Then
                ProLicense1.Text = hid_PLMsgX1.Value
                ProLicense1.ForeColor = ColorTranslator.FromHtml("#666666")
            End If
            ProLicense1.Attributes("onfocus") = "PL_focusState1();"
            ProLicense1.Attributes("onblur") = "PL_focusState1();"
            If ProLicense2.Text = "" Then
                ProLicense2.Text = hid_PLMsgX1.Value
                ProLicense2.ForeColor = ColorTranslator.FromHtml("#666666")
            End If
            ProLicense2.Attributes("onfocus") = "PL_focusState2();"
            ProLicense2.Attributes("onblur") = "PL_focusState2();"
            TIMS.PL_placeholder(tINV1)
            TIMS.PL_placeholder(tINV2)
            TIMS.PL_placeholder(tINV3)
        End If
    End Sub

    Sub AddItem()
        tr_techtype12.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then tr_techtype12.Visible = True   '產投 '顯示 講師 助教 (類別)
        IVID = TIMS.Get_Invest(IVID, objconn)
        DegreeID = TIMS.Get_Degree(DegreeID, 1, objconn)
        GraduateStatus = TIMS.Get_GradState2(GraduateStatus, objconn)
        ExpMonths = TIMS.Get_Month(ExpMonths, "", 11)
        ExpMonths1 = TIMS.Get_Month(ExpMonths1, "", 11)
        ExpMonths2 = TIMS.Get_Month(ExpMonths2, "", 11)
        ExpMonths3 = TIMS.Get_Month(ExpMonths3, "", 11)
    End Sub

    '查詢 有效資料
    Sub sCreate1(ByVal TechID As String)
        If TechID = "" Then Exit Sub

        Dim rqProecess As String = TIMS.ClearSQM(Request("proecess"))
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        sql = ""
        sql &= " SELECT a.*"
        sql &= " ,c.OrgName"
        sql &= " ,d.TrainID"
        sql &= " ,d.TrainName"
        sql &= " ,d.JobID"
        sql &= " ,d.JobName"
        sql &= " FROM TEACH_TEACHERINFO a"
        sql &= " JOIN AUTH_RELSHIP b ON a.RID=b.RID "
        sql &= " JOIN ORG_ORGINFO c ON b.OrgID=c.OrgID "
        sql &= " LEFT JOIN VIEW_TRAINTYPE d ON a.TMID=d.TMID "
        sql &= " WHERE a.TechID='" & TechID & "' "
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then Exit Sub

        If Convert.ToString(dr("RID")) <> "" Then
            If Convert.ToString(dr("RID")) = "A" Then
                LID.Value = "0"
            Else
                LID.Value = If(Convert.ToString(dr("RID")).Length = 1, "1", "2")
            End If
        End If

        If rqProecess = cst_rqProe_edit Then
            TBplanOrgName.Text = dr("OrgName").ToString
            RIDValue.Value = dr("RID").ToString
        End If
        TeacherID.Text = dr("TeacherID").ToString
        Me.ViewState("TID") = dr("TeacherID")  '暫存老師代碼
        TeachCName.Text = dr("TeachCName").ToString
        TeachEName.Text = dr("TeachEName").ToString
        IDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
        If Not IsDBNull(dr("Birthday")) Then
            birthday.Text = If(flag_ROC, TIMS.Cdate17(dr("Birthday")), TIMS.Cdate3(dr("Birthday")))
        End If
        If Convert.ToString(dr("Sex")) <> "" Then
            Common.SetListItem(Sex, dr("Sex"))
        Else
            '以身分證號作判斷
            Select Case Mid(Convert.ToString(dr("IDNO")), 2, 1)
                Case "1" '男
                    Common.SetListItem(Sex, "M")
                Case "2" '女
                    Common.SetListItem(Sex, "F")
            End Select
        End If

        'TMID-OLD
        Dim v_TB_career_id_txt As String = ""
        If Convert.ToString(dr("TrainID")) <> "" Then
            '職類
            v_TB_career_id_txt = "[" & Convert.ToString(dr("TrainID")) & "]" & Convert.ToString(dr("TrainName"))
            TB_career_id.Text = v_TB_career_id_txt 'TMID-OLD
        End If
        If v_TB_career_id_txt = "" AndAlso Convert.ToString(dr("JobID")) <> "" Then
            '若取不到職類-但有業別-顯示業別
            v_TB_career_id_txt = "[" & Convert.ToString(dr("JobID")) & "]" & Convert.ToString(dr("JobName"))
            TB_career_id.Text = v_TB_career_id_txt 'TMID-OLD
        End If
        'GCODE2-NEW
        Dim v_GCODE2 As String = ""
        ff3 = "TMID='" & Convert.ToString(dr("TMID")) & "'"
        If dtGOVCLASSCAST3.Select(ff3).Length > 0 Then v_GCODE2 = "[" & dtGOVCLASSCAST3.Select(ff3)(0)("GCODE2") & "]" & dtGOVCLASSCAST3.Select(ff3)(0)("CNAME")
        If v_GCODE2 <> "" Then TB_career_id.Text = v_GCODE2 'v_TB_career_id_txt

        trainValue.Value = Convert.ToString(dr("TMID"))
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            jobValue.Value = Convert.ToString(dr("TMID"))
        End If

        If LID.Value <= 1 Then '如果是分署(中心)或署(局)
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 OrElse dr("KindEngage").ToString <> "2" Then   '如果是內聘或是產學訓
                Common.SetListItem(IVID, dr("IVID"))
            Else                                                            '如果是外聘
                If dr("Invest").ToString <> "" Then
                    Invest1.Text = dr("Invest").ToString
                    R2.Checked = True
                    IVID.Enabled = False
                ElseIf dr("IVID").ToString <> "" Then
                    Common.SetListItem(IVID, dr("IVID"))
                    R1.Checked = True
                    Invest1.Enabled = False
                End If
            End If
        Else    '委訓單位
            If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Invest1.Text = dr("Invest").ToString '企訓專用
        End If
        Common.SetListItem(KindEngage, dr("KindEngage"))
        Dim v_KindEngage As String = Convert.ToString(dr("KindEngage"))

        If tr_techtype12.Visible Then
            '產投(顯示) 才存取此功能
            Me.cb_techtype1.Checked = False
            If Convert.ToString(dr("TechType1")) = "Y" Then Me.cb_techtype1.Checked = True
            Me.cb_techtype2.Checked = False
            If Convert.ToString(dr("TechType2")) = "Y" Then Me.cb_techtype2.Checked = True
        End If

        KindID = TIMS.Get_KindOfTeacher(KindID, v_KindEngage, "", objconn)
        Common.SetListItem(KindID, dr("KindID").ToString)

        If Convert.ToString(dr("KindID")) = cst_KindID_130 Then
            TIMS.Tooltip(KindID, "委訓單位的師資別只能是講師", True)
        End If

        Common.SetListItem(DegreeID, dr("DegreeID"))
        Common.SetListItem(GraduateStatus, dr("GraduateStatus"))
        SchoolName.Text = dr("SchoolName").ToString
        Department.Text = dr("Department").ToString
        Phone.Text = dr("Phone").ToString
        Mobile.Text = dr("Mobile").ToString
        Email.Text = dr("Email").ToString

        city_code.Value = Convert.ToString(dr("AddressZip"))
        AddressZIPB3.Value = TIMS.GetZIPCODEB3(dr("AddressZIP6W"))
        hidAddressZIP6W.Value = Convert.ToString(dr("AddressZIP6W"))
        TBCity.Text = TIMS.GET_FullCCTName(objconn, city_code.Value, AddressZIPB3.Value)
        Address.Text = dr("Address").ToString

        WorkOrg.Text = dr("WorkOrg").ToString
        ExpYears.Text = dr("ExpYears").ToString
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Invest2.Text = dr("Invest").ToString '企訓專用
        ServDept.Text = dr("ServDept").ToString
        WorkPhone.Text = dr("WorkPhone").ToString
        Fax.Text = dr("Fax").ToString

        city_code1.Value = Convert.ToString(dr("WorkZip"))
        WorkZIPB3.Value = TIMS.GetZIPCODEB3(dr("WorkZip6W"))
        hidWorkZIP6W.Value = Convert.ToString(dr("WorkZip6W"))
        TBCity1.Text = TIMS.GET_FullCCTName(objconn, city_code1.Value, WorkZIPB3.Value)
        WorkAddr.Text = dr("WorkAddr").ToString

        ExpUnit1.Text = dr("ExpUnit1").ToString
        ExpSDate1.Text = If(flag_ROC, TIMS.Cdate17(dr("ExpSDate1")), TIMS.Cdate3(dr("ExpSDate1")))
        ExpEDate1.Text = If(flag_ROC, TIMS.Cdate17(dr("ExpEDate1")), TIMS.Cdate3(dr("ExpEDate1")))
        ExpYears1.Text = dr("ExpYears1").ToString
        tINV1.Text = Convert.ToString(dr("INV1"))

        ExpUnit2.Text = dr("ExpUnit2").ToString
        ExpSDate2.Text = If(flag_ROC, TIMS.Cdate17(dr("ExpSDate2")), TIMS.Cdate3(dr("ExpSDate2")))
        ExpEDate2.Text = If(flag_ROC, TIMS.Cdate17(dr("ExpEDate2")), TIMS.Cdate3(dr("ExpEDate2")))
        ExpYears2.Text = dr("ExpYears2").ToString
        tINV2.Text = Convert.ToString(dr("INV2"))

        ExpUnit3.Text = dr("ExpUnit3").ToString
        ExpSDate3.Text = If(flag_ROC, TIMS.Cdate17(dr("ExpSDate3")), TIMS.Cdate3(dr("ExpSDate3")))
        ExpEDate3.Text = If(flag_ROC, TIMS.Cdate17(dr("ExpEDate3")), TIMS.Cdate3(dr("ExpEDate3")))
        ExpYears3.Text = dr("ExpYears3").ToString
        tINV3.Text = Convert.ToString(dr("INV3"))

        If dr("ExpMonths").ToString <> "" Then
            Common.SetListItem(ExpMonths, Convert.ToString(dr("ExpMonths")))
        End If
        If dr("ExpMonths1").ToString <> "" Then
            Common.SetListItem(ExpMonths1, Convert.ToString(dr("ExpMonths1")))
        End If
        If dr("ExpMonths2").ToString <> "" Then
            Common.SetListItem(ExpMonths2, Convert.ToString(dr("ExpMonths2")))
        End If
        If dr("ExpMonths3").ToString <> "" Then
            Common.SetListItem(ExpMonths3, Convert.ToString(dr("ExpMonths3")))
        End If
        Specialty1.Text = dr("Specialty1").ToString
        Specialty2.Text = dr("Specialty2").ToString
        Specialty3.Text = dr("Specialty3").ToString
        Specialty4.Text = dr("Specialty4").ToString
        Specialty5.Text = dr("Specialty5").ToString
        Specialty1.Text = TIMS.ClearSQM(Specialty1.Text)
        Specialty2.Text = TIMS.ClearSQM(Specialty2.Text)
        Specialty3.Text = TIMS.ClearSQM(Specialty3.Text)
        Specialty4.Text = TIMS.ClearSQM(Specialty4.Text)
        Specialty5.Text = TIMS.ClearSQM(Specialty5.Text)
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Specialty1.Text = TIMS.Get_Substr1(Specialty1.Text, 500)
            Specialty2.Text = TIMS.Get_Substr1(Specialty2.Text, 500)
            Specialty3.Text = TIMS.Get_Substr1(Specialty3.Text, 500)
            Specialty4.Text = TIMS.Get_Substr1(Specialty4.Text, 500)
            Specialty5.Text = TIMS.Get_Substr1(Specialty5.Text, 500)
        End If
        TransBook.Text = dr("TransBook").ToString
        ProLicense1.Text = hid_PLMsgX1.Value
        ProLicense1.ForeColor = ColorTranslator.FromHtml("#666666")
        If Convert.ToString(dr("ProLicense1")) <> "" Then
            ProLicense1.Text = Convert.ToString(dr("ProLicense1"))
            ProLicense1.ForeColor = ColorTranslator.FromHtml("#000000")
        End If
        ProLicense2.Text = hid_PLMsgX1.Value
        ProLicense2.ForeColor = ColorTranslator.FromHtml("#666666")
        If Convert.ToString(dr("ProLicense2")) <> "" Then
            ProLicense2.Text = Convert.ToString(dr("ProLicense2"))
            ProLicense2.ForeColor = ColorTranslator.FromHtml("#000000")
        End If
        '排課使用
        Common.SetListItem(WorkStatus, dr("WorkStatus"))
        If dr("PassPortNO").ToString <> "" Then Common.SetListItem(PassPortNO, dr("PassPortNO").ToString)
        '師資資料建立後，姓名與身分證字號均不得修正，以避免個資使用疑慮。
        '開放分署以上層級可做姓名、身分證號欄位修改。
        IDNO.Enabled = True
        TeachCName.Enabled = True
        TeachEName.Enabled = True
        Select Case sm.UserInfo.LID
            Case "2"    '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                IDNO.Enabled = False
                TeachCName.Enabled = False
                TeachEName.Enabled = False
        End Select
    End Sub

    '(新增)資料檢核 若有錯誤 則離開
    Sub CHK_INPUT_DATA(ByRef vERRMSG As String, ByVal rqTechID As String) 'As Boolean
        Dim Rst As Boolean = True
        'IF TRUE IS OK , FALSE IS ERROR
        'Dim ERRMSG As String = ""
        vERRMSG = ""
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        '資料檢核 若有錯誤 則離開
        'Dim vERRMSG As String = ""
        'If vERRMSG <> "" Then Exit Sub 'Return False

        jobValue.Value = TIMS.ClearSQM(jobValue.Value)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        Dim flag_NG_trainValue As Boolean = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If jobValue.Value = "" AndAlso trainValue.Value = "" Then flag_NG_trainValue = True 'vERRMSG = "主要職類為必填，請重新確認!!"
        Else
            If trainValue.Value = "" Then flag_NG_trainValue = True 'vERRMSG = "主要職類為必填，請重新確認!!"
        End If
        'If vERRMSG <> "" Then Exit Sub 'Return False
        If flag_NG_trainValue Then vERRMSG &= "主要職類為必填，請重新確認!!" & vbCrLf

        ExpUnit1.Text = TIMS.ClearSQM(ExpUnit1.Text)
        ExpYears1.Text = TIMS.ClearSQM(ExpYears1.Text)
        ExpMonths1.Text = TIMS.ClearSQM(ExpMonths1.Text)
        Specialty1.Text = TIMS.ClearSQM(Specialty1.Text)
        tINV1.Text = TIMS.ClearSQM(tINV1.Text)
        If (ExpUnit1.Text = "") Then vERRMSG &= "請輸入 服務單位1" & vbCrLf
        If (ExpYears1.Text = "") Then vERRMSG &= "請輸入 年資的年1" & vbCrLf
        If (ExpMonths1.Text = "") Then vERRMSG &= "請輸入 年資的月1" & vbCrLf
        If (Specialty1.Text = "") Then vERRMSG &= "請輸入 專長1" & vbCrLf
        If (tINV1.Text = "") Then vERRMSG &= "請輸入 職稱1" & vbCrLf
        If (vERRMSG = "" AndAlso tINV1.Text <> "" AndAlso tINV1.Text = TIMS.Get_placeholder(tINV1)) Then vERRMSG &= "請輸入 職稱1" & vbCrLf

        TeacherID.Text = TIMS.ClearSQM(TeacherID.Text)
        TeachCName.Text = TIMS.ClearSQM(TeachCName.Text)
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        If (TeacherID.Text = "") Then vERRMSG &= "請輸入講師代碼" & vbCrLf
        If (TeachCName.Text = "") Then vERRMSG &= "請輸入講師姓名" & vbCrLf
        If (IDNO.Text = "") Then vERRMSG &= "請輸入身分證號碼" & vbCrLf
        'IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)

        Dim flag_NG_idno As Boolean = False
        If IDNO.Text = "" Then flag_NG_idno = True 'vERRMSG = "身分證號碼有誤，請確認!!" & vbCrLf
        If IDNO.Text.IndexOf("*") > -1 Then flag_NG_idno = True 'vERRMSG = "身分證號碼有誤，請確認!!" & vbCrLf
        Dim v_PassPortNO As String = TIMS.GetListValue(PassPortNO)
        Select Case v_PassPortNO'PassPortNO.SelectedValue
            Case "1" '本國
                If Not TIMS.CheckIDNO(IDNO.Text) Then flag_NG_idno = True'vERRMSG = "身分證號碼有誤，請確認!!"
            Case "2" '外藉
                Dim nsIDNO As String = IDNO.Text
                '2:居留證 4:居留證2021
                Dim flag2 As Boolean = TIMS.CheckIDNO2(nsIDNO, 2)
                Dim flag4 As Boolean = TIMS.CheckIDNO2(nsIDNO, 4)
                If Not flag2 AndAlso Not flag4 Then
                    flag_NG_idno = True 'vERRMSG = "身分別為外藉，居留證號有誤，請確認!!"
                End If
                'If Not TIMS.CheckIDNO2(IDNO.Text, 8) Then flag_NG_idno = True 'vERRMSG = "身分別為外藉，居留證號有誤，請確認!!"
            Case Else
        End Select
        If (flag_NG_idno) Then vERRMSG &= "身分證號碼或居留證號 有誤，請確認!!" & vbCrLf

        Dim fg_CHKDATE As Boolean = False

        birthday.Text = TIMS.ClearSQM(birthday.Text)
        If birthday.Text = "" Then
            vERRMSG &= "請輸入出生日期!" & vbCrLf
        Else
            fg_CHKDATE = If(flag_ROC, TIMS.IsDate7(birthday.Text), TIMS.IsDate1(birthday.Text))
            If Not fg_CHKDATE Then vERRMSG &= "出生日期格式不正確!" & vbCrLf
        End If

        If (TB_career_id.Text = "") Then vERRMSG &= "請選擇主要職類!" & vbCrLf

        '選擇內外聘
        Dim v_KindEngage As String = TIMS.GetListValue(KindEngage)
        '師資別
        Dim v_KindID As String = TIMS.GetListValue(KindID)
        Dim v_IVID As String = TIMS.GetListValue(IVID)
        Invest1.Text = TIMS.ClearSQM(Invest1.Text)
        If v_KindEngage = "" OrElse v_KindEngage = "0" Then vERRMSG &= "請選擇內外聘!" & vbCrLf ' = msg + '請選擇內外聘!\n';
        Select Case sm.UserInfo.LID
            Case 2
                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 AndAlso Invest1.Visible Then
                    If Invest1.Text = "" Then vERRMSG &= "請輸入職稱!" & vbCrLf
                End If
            Case Else
                If IVID.Visible Then
                    If v_KindEngage <> cst_KindEngage_2_外聘 OrElse TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        If v_IVID = "" Then vERRMSG &= "請選擇職稱!" & vbCrLf
                    Else
                        If v_IVID = "" AndAlso Invest1.Text = "" Then vERRMSG &= "請選擇 職稱!" & vbCrLf
                    End If
                End If
                '非委訓單位提供選擇 
                If Not (sm.UserInfo.LID = 2 OrElse LID.Value = "2") Then
                    'If v_KindID = "" OrElse v_KindID = "0" Then
                    '    KindID.Enabled = True
                    '    vERRMSG &= "請選擇師資別!!" & vbCrLf
                    'End If
                End If
        End Select

        '最高學歷
        Dim v_DegreeID As String = TIMS.GetListValue(DegreeID)
        If v_DegreeID = "" Then vERRMSG &= "請選擇 最高學歷 !!" & vbCrLf
        '畢業狀況
        Dim v_GraduateStatus As String = TIMS.GetListValue(GraduateStatus)
        If v_GraduateStatus = "" Then vERRMSG &= "請選擇 畢業狀況 !!" & vbCrLf
        '    If (document.form1.DegreeID.selectedIndex == '0') msg = msg + '請選擇最高學歷!\n';
        '    If(document.form1.GraduateStatus.selectedIndex == '0') msg = msg + '請選擇畢業狀況!\n';

        SchoolName.Text = TIMS.ClearSQM(SchoolName.Text)
        If SchoolName.Text = "" Then vERRMSG &= "請輸入 學校名稱 !!" & vbCrLf
        Department.Text = TIMS.ClearSQM(Department.Text)
        If Department.Text = "" Then vERRMSG &= "請輸入 科系名稱 !!" & vbCrLf
        '    If(SchoolName.value == "") { msg += '請輸入 學校名稱!\n'; }
        '    If (Department.value == "") Then { msg += '請輸入 科系名稱!\n'; }

        Phone.Text = TIMS.ClearSQM(Phone.Text)
        If Phone.Text = "" Then vERRMSG &= "請輸入聯絡電話 !!" & vbCrLf

        'ADDRESSZIP
        city_code.Value = TIMS.ClearSQM(city_code.Value)
        If city_code.Value = "" Then vERRMSG &= "請輸入通訊地址的郵遞區號3碼 !!" & vbCrLf
        AddressZIPB3.Value = TIMS.ClearSQM(AddressZIPB3.Value)
        TIMS.CheckZipCODEB3(AddressZIPB3.Value, "通訊地址郵遞區號後2碼或後3碼", True, vERRMSG)
        Address.Text = TIMS.ClearSQM(Address.Text)
        If Address.Text = "" Then vERRMSG &= "請輸入通訊地址 !!" & vbCrLf

        '服務單位名稱
        WorkOrg.Text = TIMS.ClearSQM(WorkOrg.Text)
        If WorkOrg.Text = "" Then vERRMSG &= "服務單位名稱為必填，請重新確認!!" & vbCrLf
        '    If (document.form1.Address.value == '') msg += '請輸入通訊地址\n';
        '    If(document.form1.WorkOrg.value == '') msg += '請輸入服務單位\n';
        '服務單位電話
        WorkPhone.Text = TIMS.ClearSQM(WorkPhone.Text)
        If WorkPhone.Text = "" Then vERRMSG &= "請輸入服務單位電話 !!" & vbCrLf

        'WorkZip
        city_code1.Value = TIMS.ClearSQM(city_code1.Value)
        'If city_code1.Value = "" Then vERRMSG &= "請輸入服務單位地址的郵遞區號3碼 !!" & vbCrLf
        WorkZIPB3.Value = TIMS.ClearSQM(WorkZIPB3.Value)
        TIMS.CheckZipCODEB3(WorkZIPB3.Value, "服務單位地址郵遞區號後2碼或後3碼", False, vERRMSG)
        WorkAddr.Text = TIMS.ClearSQM(WorkAddr.Text)
        'If WorkAddr.Text = "" Then vERRMSG &= "請輸入服務單位地址 !!" & vbCrLf

        ExpYears1.Text = TIMS.ClearSQM(ExpYears1.Text)
        ExpYears2.Text = TIMS.ClearSQM(ExpYears2.Text)
        ExpYears3.Text = TIMS.ClearSQM(ExpYears3.Text)
        If ExpYears1.Text <> "" AndAlso Not TIMS.IsInt(ExpYears1.Text) Then
            vERRMSG &= " 經歷年資1必須為數字!" & vbCrLf
        End If
        If ExpYears2.Text <> "" AndAlso Not TIMS.IsInt(ExpYears2.Text) Then
            vERRMSG &= " 經歷年資2必須為數字!" & vbCrLf
        End If
        If ExpYears3.Text <> "" AndAlso Not TIMS.IsInt(ExpYears3.Text) Then
            vERRMSG &= " 經歷年資3必須為數字!" & vbCrLf
        End If

        ExpSDate1.Text = TIMS.ClearSQM(ExpSDate1.Text)
        ExpEDate1.Text = TIMS.ClearSQM(ExpEDate1.Text)
        If ExpSDate1.Text <> "" Then
            fg_CHKDATE = If(flag_ROC, TIMS.IsDate7(ExpSDate1.Text), TIMS.IsDate1(ExpSDate1.Text))
            If Not fg_CHKDATE Then vERRMSG &= " 服務期間1起始 日期格式不正確!" & vbCrLf
        End If
        If ExpEDate1.Text <> "" Then
            fg_CHKDATE = If(flag_ROC, TIMS.IsDate7(ExpEDate1.Text), TIMS.IsDate1(ExpEDate1.Text))
            If Not fg_CHKDATE Then vERRMSG &= " 服務期間1終止 日期格式不正確!" & vbCrLf
        End If

        ExpSDate2.Text = TIMS.ClearSQM(ExpSDate2.Text)
        ExpEDate2.Text = TIMS.ClearSQM(ExpEDate2.Text)
        If ExpSDate2.Text <> "" Then
            fg_CHKDATE = If(flag_ROC, TIMS.IsDate7(ExpSDate2.Text), TIMS.IsDate1(ExpSDate2.Text))
            If Not fg_CHKDATE Then vERRMSG &= " 服務期間2起始 日期格式不正確!" & vbCrLf
        End If
        If ExpEDate2.Text <> "" Then
            fg_CHKDATE = If(flag_ROC, TIMS.IsDate7(ExpEDate2.Text), TIMS.IsDate1(ExpEDate2.Text))
            If Not fg_CHKDATE Then vERRMSG &= " 服務期間2終止 日期格式不正確!" & vbCrLf
        End If

        ExpSDate3.Text = TIMS.ClearSQM(ExpSDate3.Text)
        ExpEDate3.Text = TIMS.ClearSQM(ExpEDate3.Text)
        If ExpSDate3.Text <> "" Then
            fg_CHKDATE = If(flag_ROC, TIMS.IsDate7(ExpSDate3.Text), TIMS.IsDate1(ExpSDate3.Text))
            If Not fg_CHKDATE Then vERRMSG &= " 服務期間3起始 日期格式不正確!" & vbCrLf
        End If
        If ExpEDate3.Text <> "" Then
            fg_CHKDATE = If(flag_ROC, TIMS.IsDate7(ExpEDate3.Text), TIMS.IsDate1(ExpEDate3.Text))
            If Not fg_CHKDATE Then vERRMSG &= " 服務期間3終止 日期格式不正確!" & vbCrLf
        End If
        Dim v_WorkStatus As String = TIMS.GetListValue(WorkStatus)
        If v_WorkStatus = "" OrElse v_WorkStatus = "0" Then vERRMSG &= "請選擇 排課使用!" & vbCrLf ' = msg + '請選擇內外聘!\n';
        '    If (document.form1.WorkStatus.selectedIndex == '0') {
        '        msg = msg + '請選擇任職狀況!\n';
        '    }
        TransBook.Text = TIMS.ClearSQM(TransBook.Text)
        If (TransBook.Text.Length > 100) Then vERRMSG &= "【譯著】長度不可超過100字元!" & vbCrLf

        ProLicense1.Text = TIMS.ClearSQM(ProLicense1.Text)
        If (ProLicense1.Text = hid_PLMsgX1.Value OrElse ProLicense1.Text = "") Then
            vERRMSG &= "請輸入 專業證照-政府機關辦理相關證照或檢定!" & vbCrLf
        Else
            If (ProLicense1.Text.Length > 200) Then vERRMSG &= "【專業證照-政府機關辦理相關證照或檢定】長度不可超過200字元!" & vbCrLf
        End If

        ProLicense2.Text = TIMS.ClearSQM(ProLicense2.Text)
        If (ProLicense2.Text = hid_PLMsgX1.Value OrElse ProLicense2.Text = "") Then
            vERRMSG &= "請輸入 專業證照-其他證照或檢定!" & vbCrLf
        Else
            If (ProLicense2.Text.Length > 200) Then vERRMSG &= "【專業證照-其他證照或檢定】長度不可超過200字元!" & vbCrLf
        End If
        If vERRMSG <> "" Then Exit Sub 'Return False

        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        '驗證IDNO 是否已建立
        Select Case Proecess.Text
            Case Cst_ProeTxt_新增
                sql = ""
                sql &= " SELECT 'x' FROM TEACH_TEACHERINFO "
                sql &= " WHERE RID = '" & RIDValue.Value & "' "
                sql &= " AND IDNO = '" & IDNO.Text & "' "
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count > 0 Then
                    vERRMSG = Cst_msg2
                    Exit Sub 'Return False 'Common.MessageBox(Me, Cst_msg2)     Exit Function
                End If
                'dr = DbAccess.GetOneRow(sql, objconn)
                'If Not dr Is Nothing Then
                '    If CInt(dr("CNT").ToString) > 0 Then
                '        vERRMSG = Cst_msg2
                '        Return False 'Common.MessageBox(Me, Cst_msg2)   Exit Function
                '    End If
                'End If

            Case Cst_ProeTxt_修改
                'sql = ""
                'sql &= " SELECT 'x' FROM TEACH_TEACHERINFO"
                'sql &= " WHERE RID='" & If(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value) & "'"
                'sql &= " and IDNO='" & IDNO.Text & "' "
                'sql &= " and TechID!='" & rqTechID & "'"
                'dr = DbAccess.GetOneRow(sql, objconn)
                'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
                'If dt.Rows.Count > 0 Then
                '    vERRMSG = Cst_msg2
                '    Return False 'Common.MessageBox(Me, Cst_msg2)                        Exit Function
                'End If
                ''講師代碼重複 (允許儲存)
                'Sql = "SELECT * FROM TEACH_TEACHERINFO WHERE RID='" & If(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value) & "' and TeacherID='" & TeacherID.Text & "' and TechID<>'" & rqTechID & "'"
                'dr = DbAccess.GetOneRow(Sql, objconn)
                'If Not dr Is Nothing Then
                '    double_flag = True '講師代碼重複
                'End If
        End Select
        If vERRMSG <> "" Then Exit Sub 'Return False

        If TIMS.LENB(TeacherID.Text) > 10 Then vERRMSG += Cst_msg4 & vbCrLf
        If TeachCName.MaxLength > 0 AndAlso TIMS.LENB(TeachCName.Text) > TeachCName.MaxLength Then vERRMSG += "講師姓名 長度最多" & TeachCName.MaxLength & "碼" & vbCrLf
        If TeachEName.MaxLength > 0 AndAlso TIMS.LENB(TeachEName.Text) > TeachEName.MaxLength Then vERRMSG += "講師英文姓名 長度最多" & TeachEName.MaxLength & "碼" & vbCrLf
        If vERRMSG <> "" Then Exit Sub
        If tr_techtype12.Visible Then
            Dim iTechType As Integer = 0
            '產投(顯示) 才存取此功能
            iTechType = 0
            If Me.cb_techtype1.Checked = True Then iTechType += 1
            If Me.cb_techtype2.Checked = True Then iTechType += 1
            If iTechType = 0 Then vERRMSG += Cst_msg8 & vbCrLf '至少要有一項
        End If
        Specialty1.Text = TIMS.ClearSQM(Specialty1.Text)
        Specialty2.Text = TIMS.ClearSQM(Specialty2.Text)
        Specialty3.Text = TIMS.ClearSQM(Specialty3.Text)
        Specialty4.Text = TIMS.ClearSQM(Specialty4.Text)
        Specialty5.Text = TIMS.ClearSQM(Specialty5.Text)
        If Specialty5.Text <> "" AndAlso Specialty4.Text = "" Then vERRMSG &= "專長5有填, 專長4不可為空" & vbCrLf
        If Specialty4.Text <> "" AndAlso Specialty3.Text = "" Then vERRMSG &= "專長4有填, 專長3不可為空" & vbCrLf
        If Specialty3.Text <> "" AndAlso Specialty2.Text = "" Then vERRMSG &= "專長3有填, 專長2不可為空" & vbCrLf
        If Specialty2.Text <> "" AndAlso Specialty1.Text = "" Then vERRMSG &= "專長2有填, 專長1不可為空" & vbCrLf
        If vERRMSG <> "" Then Rst = False
        Dim vSpecialty1 As String = TIMS.Get_Substr1(Specialty1.Text, 500)
        Dim vSpecialty2 As String = TIMS.Get_Substr1(Specialty2.Text, 500)
        Dim vSpecialty3 As String = TIMS.Get_Substr1(Specialty3.Text, 500)
        Dim vSpecialty4 As String = TIMS.Get_Substr1(Specialty4.Text, 500)
        Dim vSpecialty5 As String = TIMS.Get_Substr1(Specialty5.Text, 500)
        If vSpecialty1 <> Specialty1.Text Then vERRMSG &= "專長1，字串長度不可超過500!" & vbCrLf
        If vSpecialty2 <> Specialty2.Text Then vERRMSG &= "專長2，字串長度不可超過500!" & vbCrLf
        If vSpecialty3 <> Specialty3.Text Then vERRMSG &= "專長3，字串長度不可超過500!" & vbCrLf
        If vSpecialty4 <> Specialty4.Text Then vERRMSG &= "專長4，字串長度不可超過500!" & vbCrLf
        If vSpecialty5 <> Specialty5.Text Then vERRMSG &= "專長5，字串長度不可超過500!" & vbCrLf
        If vERRMSG <> "" Then Rst = False
        'Return Rst
    End Sub

    '儲存
    Sub SaveData1(ByVal double_flag As Boolean)
        'double_flag'講師代碼重複 (true:(重複)允許儲存 false:沒有重複)
        Dim rqProecess As String = TIMS.ClearSQM(Request("proecess"))
        Dim rqTechID As String = TIMS.ClearSQM(Request("serial"))
        'rqProecess = TIMS.ClearSQM(rqProecess)
        'rqTechID = TIMS.ClearSQM(rqTechID)
        Phone.Text = TIMS.ClearSQM(Phone.Text)
        Mobile.Text = TIMS.ClearSQM(Mobile.Text)
        Fax.Text = TIMS.ClearSQM(Fax.Text)
        TransBook.Text = TIMS.ClearSQM(TransBook.Text)
        ProLicense1.Text = TIMS.ClearSQM(ProLicense1.Text)
        ProLicense2.Text = TIMS.ClearSQM(ProLicense2.Text)
        Specialty1.Text = TIMS.Get_Substr1(TIMS.ClearSQM(Specialty1.Text), 500)
        Specialty2.Text = TIMS.Get_Substr1(TIMS.ClearSQM(Specialty2.Text), 500)
        Specialty3.Text = TIMS.Get_Substr1(TIMS.ClearSQM(Specialty3.Text), 500)
        Specialty4.Text = TIMS.Get_Substr1(TIMS.ClearSQM(Specialty4.Text), 500)
        Specialty5.Text = TIMS.Get_Substr1(TIMS.ClearSQM(Specialty5.Text), 500)

        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        'Dim double_flag As Boolean = False '講師代碼重複 (允許儲存)
        Select Case Proecess.Text
            Case Cst_ProeTxt_新增
                'sql = "SELECT * FROM TEACH_TEACHERINFO WHERE 1<>1"
            Case Else 'Cst_修改
                If rqTechID = "" Then
                    Common.MessageBox(Me, "查無有效師資資料，請檢查輸入資料!!")
                    Exit Sub
                End If
        End Select

        '選擇內外聘
        Dim v_KindEngage As String = TIMS.GetListValue(KindEngage)
        '師資別
        Dim v_KindID As String = TIMS.GetListValue(KindID)
        Dim v_IVID As String = TIMS.GetListValue(IVID)
        Dim iTECHID As Integer = 0
        Select Case Proecess.Text
            Case Cst_ProeTxt_新增
                iTECHID = DbAccess.GetNewId(objconn, "TEACH_TEACHERINFO_TECHID_SEQ,TEACH_TEACHERINFO,TECHID")
                sql = " SELECT * FROM TEACH_TEACHERINFO WHERE 1<>1 "
            Case Else 'Cst_修改
                iTECHID = Val(rqTechID)
                sql = " SELECT * FROM TEACH_TEACHERINFO WHERE TechID = '" & rqTechID & "' "
        End Select
        If iTECHID = 0 Then Exit Sub

        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("TECHID") = iTECHID 'TEACH_TEACHERINFO_TECHID_SEQ
        Else
            dr = dt.Rows(0)
            iTECHID = dr("TECHID")
        End If
        dr("RID") = RIDValue.Value
        dr("TeacherID") = TeacherID.Text
        dr("TeachCName") = TeachCName.Text
        dr("TeachEName") = If(TeachEName.Text = "", Convert.DBNull, TeachEName.Text)
        dr("Birthday") = If(birthday.Text = "", Convert.DBNull, If(flag_ROC, TIMS.Cdate18(birthday.Text), birthday.Text))  'edit，by:20181022
        dr("IDNO") = TIMS.ChangeIDNO(IDNO.Text)
        Dim v_TMID As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            v_TMID = If(jobValue.Value <> "", jobValue.Value, trainValue.Value)
        Else
            v_TMID = If(trainValue.Value <> "", trainValue.Value, "")
        End If
        dr("TMID") = If(v_TMID <> "", v_TMID, Convert.DBNull)

        Dim v_DegreeID As String = TIMS.GetListValue(DegreeID)
        Dim v_GraduateStatus As String = TIMS.GetListValue(GraduateStatus)
        dr("DegreeID") = If(v_DegreeID <> "", v_DegreeID, "") 'DegreeID.SelectedValue
        dr("SchoolName") = If(SchoolName.Text = "", Convert.DBNull, SchoolName.Text)
        dr("Department") = If(Department.Text = "", Convert.DBNull, Department.Text)
        dr("GraduateStatus") = v_GraduateStatus 'GraduateStatus.SelectedValue
        dr("Phone") = Phone.Text
        hidAddressZIP6W.Value = TIMS.GetZIPCODE6W(city_code.Value, AddressZIPB3.Value)
        dr("AddressZip") = If(city_code.Value <> "", city_code.Value, Convert.DBNull)
        dr("AddressZIP6W") = If(hidAddressZIP6W.Value <> "", hidAddressZIP6W.Value, Convert.DBNull)
        dr("Address") = If(Address.Text <> "", Address.Text, Convert.DBNull)

        dr("WorkOrg") = WorkOrg.Text
        Invest2.Text = TIMS.ClearSQM(Invest2.Text)
        Invest1.Text = TIMS.ClearSQM(Invest1.Text)
        If LID.Value = 2 Then   '委訓單位
            dr("IVID") = Convert.DBNull
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                dr("Invest") = Invest2.Text '企訓專用
            Else
                dr("Invest") = Invest1.Text
            End If
        Else  '分署(中心)
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 OrElse v_KindEngage = "1" Then '假如是產業人才投資方案或是內聘
                dr("IVID") = v_IVID 'IVID.SelectedValue
                dr("Invest") = Convert.DBNull
            Else '外聘
                'If IVID.SelectedValue <> -1 Or IVID.SelectedValue <> "" Then '職稱的選單選項是否為空值,不是的話就存選單選項
                If IVID.SelectedIndex <> 0 Then
                    dr("IVID") = v_IVID 'IVID.SelectedValue
                    dr("Invest") = Convert.DBNull
                ElseIf Invest1.Text <> "" Then '職稱的textbox是否為空值,不是的話就存textbox
                    dr("IVID") = Convert.DBNull
                    dr("Invest") = Invest1.Text
                End If
            End If
        End If
        If tr_techtype12.Visible Then
            '產投(顯示) 才存取此功能
            dr("TechType1") = If(cb_techtype1.Checked, "Y", Convert.DBNull) '講師(類別) 清空
            dr("TechType2") = If(cb_techtype2.Checked, "Y", Convert.DBNull) '助教(類別) 清空
        End If
        dr("ExpYears") = If(ExpYears.Text = "", Convert.DBNull, ExpYears.Text)

        hidWorkZIP6W.Value = TIMS.GetZIPCODE6W(city_code1.Value, WorkZIPB3.Value)
        dr("WorkZip") = If(city_code1.Value <> "", city_code1.Value, Convert.DBNull)
        dr("WorkZIP6W") = If(hidWorkZIP6W.Value <> "", hidWorkZIP6W.Value, Convert.DBNull)
        dr("WorkAddr") = If(WorkAddr.Text <> "", WorkAddr.Text, Convert.DBNull)

        tINV1.Text = TIMS.ClearSQM(tINV1.Text)
        tINV2.Text = TIMS.ClearSQM(tINV2.Text)
        tINV3.Text = TIMS.ClearSQM(tINV3.Text)
        dr("WorkPhone") = If(WorkPhone.Text = "", Convert.DBNull, WorkPhone.Text)

        dr("ExpUnit1") = If(ExpUnit1.Text = "", Convert.DBNull, ExpUnit1.Text)
        dr("ExpSDate1") = If(ExpSDate1.Text = "", Convert.DBNull, TIMS.Cdate2(If(flag_ROC, TIMS.Cdate18(ExpSDate1.Text), ExpSDate1.Text)))  'edit，by:20181022
        dr("ExpEDate1") = If(ExpEDate1.Text = "", Convert.DBNull, TIMS.Cdate2(If(flag_ROC, TIMS.Cdate18(ExpEDate1.Text), ExpEDate1.Text)))  'edit，by:20181022
        dr("ExpYears1") = If(ExpYears1.Text = "", Convert.DBNull, ExpYears1.Text)

        dr("ExpUnit2") = If(ExpUnit2.Text = "", Convert.DBNull, ExpUnit2.Text)
        dr("ExpSDate2") = If(ExpSDate2.Text = "", Convert.DBNull, TIMS.Cdate2(If(flag_ROC, TIMS.Cdate18(ExpSDate2.Text), ExpSDate2.Text)))  'edit，by:20181022
        dr("ExpEDate2") = If(ExpEDate2.Text = "", Convert.DBNull, TIMS.Cdate2(If(flag_ROC, TIMS.Cdate18(ExpEDate2.Text), ExpEDate2.Text)))  'edit，by:20181022
        dr("ExpYears2") = If(ExpYears2.Text = "", Convert.DBNull, ExpYears2.Text)

        dr("ExpUnit3") = If(ExpUnit3.Text = "", Convert.DBNull, ExpUnit3.Text)
        dr("ExpSDate3") = If(ExpSDate3.Text = "", Convert.DBNull, TIMS.Cdate2(If(flag_ROC, TIMS.Cdate18(ExpSDate3.Text), ExpSDate3.Text)))  'edit，by:20181022
        dr("ExpEDate3") = If(ExpEDate3.Text = "", Convert.DBNull, TIMS.Cdate2(If(flag_ROC, TIMS.Cdate18(ExpEDate3.Text), ExpEDate3.Text)))  'edit，by:20181022
        dr("ExpYears3") = If(ExpYears3.Text = "", Convert.DBNull, ExpYears3.Text)

        TIMS.Chk_placeholder(tINV1)
        TIMS.Chk_placeholder(tINV2)
        TIMS.Chk_placeholder(tINV3)
        dr("INV1") = If(tINV1.Text = "", Convert.DBNull, tINV1.Text)
        dr("INV2") = If(tINV2.Text = "", Convert.DBNull, tINV2.Text)
        dr("INV3") = If(tINV3.Text = "", Convert.DBNull, tINV3.Text)

        Dim v_ExpMonths As String = TIMS.GetListValue(ExpMonths)
        Dim v_ExpMonths1 As String = TIMS.GetListValue(ExpMonths1)
        Dim v_ExpMonths2 As String = TIMS.GetListValue(ExpMonths2)
        Dim v_ExpMonths3 As String = TIMS.GetListValue(ExpMonths3)
        dr("ExpMonths") = If(v_ExpMonths <> "", v_ExpMonths, Convert.DBNull) ', ExpMonths.SelectedValue)
        dr("ExpMonths1") = If(v_ExpMonths1 <> "", v_ExpMonths1, Convert.DBNull) 'If(ExpMonths1.SelectedIndex = 0, Convert.DBNull, ExpMonths1.SelectedValue)
        dr("ExpMonths2") = If(v_ExpMonths2 <> "", v_ExpMonths2, Convert.DBNull) 'If(ExpMonths2.SelectedIndex = 0, Convert.DBNull, ExpMonths2.SelectedValue)
        dr("ExpMonths3") = If(v_ExpMonths3 <> "", v_ExpMonths3, Convert.DBNull) 'If(ExpMonths3.SelectedIndex = 0, Convert.DBNull, ExpMonths3.SelectedValue)
        dr("Specialty1") = If(Specialty1.Text = "", Convert.DBNull, Specialty1.Text)
        dr("Specialty2") = If(Specialty2.Text = "", Convert.DBNull, Specialty2.Text)
        dr("Specialty3") = If(Specialty3.Text = "", Convert.DBNull, Specialty3.Text)
        dr("Specialty4") = If(Specialty4.Text = "", Convert.DBNull, Specialty4.Text)
        dr("Specialty5") = If(Specialty5.Text = "", Convert.DBNull, Specialty5.Text)

        'Dim v_KindID As String = TIMS.GetListValue(KindID)
        'Dim v_KindEngage As String = TIMS.GetListValue(KindEngage)
        Dim v_WorkStatus As String = TIMS.GetListValue(WorkStatus)
        Dim v_Sex As String = TIMS.GetListValue(Sex)
        Dim v_PassPortNO As String = TIMS.GetListValue(PassPortNO)

        '確認是委訓單位
        'dr("KindID") = Convert.DBNull 
        '130是講師 2010/12/13 號改成委訓單位都可以看到師資別,但是不可以選,
        'flag_LID_2 = True '確認是委訓單位
        Dim flag_LID_2 As Boolean = (sm.UserInfo.LID = 2 OrElse LID.Value = "2")
        dr("KindID") = If(flag_LID_2, cst_KindID_130, v_KindID) '130

        dr("KindEngage") = If(v_KindEngage <> "", v_KindEngage, Convert.DBNull) 'KindEngage.SelectedValue
        dr("WorkStatus") = If(v_WorkStatus <> "", v_WorkStatus, Convert.DBNull) 'WorkStatus.SelectedValue '排課使用
        dr("Sex") = If(v_Sex <> "", v_Sex, Convert.DBNull) 'Sex.SelectedValue
        dr("Mobile") = If(Mobile.Text = "", Convert.DBNull, Mobile.Text)
        dr("Email") = If(Email.Text = "", Convert.DBNull, Email.Text)
        dr("ServDept") = If(ServDept.Text = "", Convert.DBNull, ServDept.Text)
        dr("Fax") = If(Fax.Text = "", Convert.DBNull, Fax.Text)
        dr("TransBook") = If(TransBook.Text = "", Convert.DBNull, TransBook.Text)
        '未輸入有效資訊!!
        If ProLicense1.Text = hid_PLMsgX1.Value Then ProLicense1.Text = ""
        dr("ProLicense1") = If(ProLicense1.Text = "", Convert.DBNull, ProLicense1.Text)
        If ProLicense2.Text = hid_PLMsgX1.Value Then ProLicense2.Text = ""
        dr("ProLicense2") = If(ProLicense2.Text = "", Convert.DBNull, ProLicense2.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        Select Case v_PassPortNO'PassPortNO.SelectedValue
            Case "1", "2"
                dr("PassPortNO") = v_PassPortNO 'PassPortNO.SelectedValue
            Case Else
                dr("PassPortNO") = "2"
        End Select
        DbAccess.UpdateDataTable(dt, da)

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Select Case Proecess.Text
            Case Cst_ProeTxt_新增
                Dim sMemo As String = ""
                sMemo = ""
                sMemo &= "&NAME=" & TeachCName.Text
                '寫入Log查詢(SubInsAccountLog1(Auth_Accountlog))
                Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm新增, Session(TIMS.gcst_rblWorkMode), "", sMemo)
            Case Else 'Cst_修改
                Dim sMemo As String = ""
                sMemo = ""
                sMemo &= "&NAME=" & TeachCName.Text
                '寫入Log查詢(SubInsAccountLog1(Auth_Accountlog))
                Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm修改, Session(TIMS.gcst_rblWorkMode), "", sMemo)
        End Select
        If Session("MySearchStr") Is Nothing Then Session("MySearchStr") = ViewState("MySearchStr")
        Dim msg_tmp As String = ""
        msg_tmp = ""
        '訊息選擇
        Select Case rqProecess
            Case cst_rqProe_add, cst_rqProe_Insert
                'Cst_新增
                msg_tmp += Cst_msg5
            Case Else
                'Cst_修改
                msg_tmp += Cst_msg6
        End Select
        If double_flag Then
            '講師代碼重複 (允許儲存)
            msg_tmp += Cst_msg3
        End If
        Common.RespWrite(Me, "<script language=javascript>")
        Common.RespWrite(Me, "window.alert('" & msg_tmp & "');")
        Common.RespWrite(Me, "window.location.href='TC_01_027.aspx?ID=" & MRqID & "';")
        Common.RespWrite(Me, "</script>")
#Region "(No Use)"

        'Try
        'Catch ex As Exception
        '    Common.RespWrite(Me, "<script language=javascript>window.alert('" & Cst_msg7 & "');")
        '    Common.MessageBox(Me, ex.ToString)
        'End Try

#End Region
    End Sub

    '存檔。
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim rqTechID As String = TIMS.ClearSQM(Request("serial"))

        Dim double_flag As Boolean = False '(無重複) '講師代碼重複 (允許儲存)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        If IDNO.Text <> "" Then IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)

        '資料檢核 若有錯誤 則離開
        Dim vERRMSG As String = ""
        Call CHK_INPUT_DATA(vERRMSG, rqTechID)
        If vERRMSG <> "" Then
            Common.MessageBox(Me, vERRMSG)
            Exit Sub
        End If

        '驗證IDNO 是否已建立
        Select Case Proecess.Text
            Case Cst_ProeTxt_新增
                'sql = "SELECT COUNT(1) CNT FROM TEACH_TEACHERINFO WHERE RID='" & If(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value) & "' and IDNO='" & IDNO.Text & "' "
                'dr = DbAccess.GetOneRow(sql, objconn)
                'If Not dr Is Nothing Then
                '    If CInt(dr("CNT").ToString) > 0 Then
                '        vERRMSG = Cst_msg2
                '        Return False 'Common.MessageBox(Me, Cst_msg2) Exit Function
                '    End If
                'End If

            Case Cst_ProeTxt_修改
                '講師代碼重複 (允許儲存)
                Dim dr As DataRow = Nothing
                Dim sql As String = ""
                sql = ""
                sql &= " SELECT * FROM TEACH_TEACHERINFO "
                sql &= " WHERE RID = '" & If(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value) & "' "
                sql &= " AND TeacherID = '" & TeacherID.Text & "' "
                sql &= " AND TechID <> '" & rqTechID & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If dr IsNot Nothing Then double_flag = True '講師代碼重複
        End Select

        Call SaveData1(double_flag)
    End Sub

    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
        'Session("MySearchStr") = ViewState("MySearchStr")
        If Session("MySearchStr") Is Nothing Then Session("MySearchStr") = ViewState("MySearchStr")
        DbAccess.CloseDbConn(objconn)
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect1(Me, "TC_01_027.aspx?ID=" & MRqID & "")
    End Sub

    Private Sub KindEngage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KindEngage.SelectedIndexChanged
        Dim v_KindEngage As String = TIMS.GetListValue(KindEngage)
        Select Case v_KindEngage
            Case "1", "2"
                '動態產生師資別
                KindID = TIMS.Get_KindOfTeacher(KindID, v_KindEngage, "", objconn) 'KindEngage.SelectedValue)
            Case Else
                KindID.Items.Clear()
                KindID.Items.Add(New ListItem("請選擇內外聘", "0"))
        End Select

        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso LID.Value <= 1 Then '如果不是產學訓而且是分署(中心)或署(局) 【TIMS計畫】
            R1.Checked = True
            R2.Checked = False
            IVID.SelectedIndex = 0
            IVID.Enabled = True
            Invest1.Text = ""
            Invest1.Enabled = False
        End If
    End Sub

    '複製產生(隱藏)
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If TechID.Value <> "" Then Call sCreate1(TechID.Value)
    End Sub
End Class