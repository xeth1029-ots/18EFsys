Partial Class TC_01_007_add
    Inherits AuthBasePage

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
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY1", Specialty1)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY2", Specialty2)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY3", Specialty3)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY4", Specialty4)
        'Call TIMS.sUtl_SetMaxLen(dt, "SPECIALTY5", Specialty5)
        Call TIMS.sUtl_SetMaxLen(dt, "EMAIL", Email)
        Call TIMS.sUtl_SetMaxLen(dt, "SERVDEPT", ServDept)
        Call TIMS.sUtl_SetMaxLen(dt, "FAX", Fax)
        Call TIMS.sUtl_SetMaxLen(dt, "TRANSBOOK", TransBook)
        Call TIMS.sUtl_SetMaxLen(dt, "PROLICENSE", ProLicense)
    End Sub

    Const cst_rqProe_Insert As String = "Insert"
    Const cst_rqProe_add As String = "add"
    Const cst_rqProe_edit As String = "edit"
    Const Cst_ProeTxt_新增 As String = "-新增"
    Const Cst_ProeTxt_修改 As String = "-修改"
    Const Cst_msg1 As String = "建議使用 區碼-電話號碼"
    Const Cst_msg2 As String = "同一計劃，有相同身分證號碼，重複輸入" '依RIDValue.Value
    Const Cst_msg3 As String = "\n講師代碼發生重複，請注意"
    Const Cst_msg4 As String = "講師代碼最多10 碼(英文 / 數字)"
    Const Cst_msg5 As String = "資料新增成功!"
    Const Cst_msg6 As String = "資料修改成功!"
    Const Cst_msg7 As String = "資料儲存有誤，請再確認輸入資料正確性!!!"
    Const Cst_msg8 As String = "類別，請選擇講師或助教!!"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
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

        If Not IsPostBack Then
            msg.Text = ""
            '保存查詢值---Start
            ViewState("MySearchStr") = ""
            If Session("MySearchStr") IsNot Nothing Then ViewState("MySearchStr") = Session("MySearchStr")
            'Session("MySearchStr") = Nothing
            '保存查詢值---End
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
                    If Len(sm.UserInfo.RID) > 1 And RIDValue.Value <> sm.UserInfo.RID Then
                        Button1.Visible = False
                    End If

            End Select

            'LID【0=>署(局), 1=>分署(中心), 2=>委訓】
            If LID.Value = 2 Then
                'KindID_TD1.Style.Item("display") = "none" 
                'KindID_TD2.Style.Item("display") = "none"
                '原本委訓單位沒有師資別，後來改成有師資別可以看但是不能選,委訓單位的師資別統一只能是講師2010/12/13
                KindID.Items.Clear()
                KindID.Items.Insert(0, New ListItem("講師", 130))
                KindID.SelectedValue = 130
                KindID.Enabled = False

                KindEngage.AutoPostBack = False
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
                KindEngage.AutoPostBack = True
            End If
        End If


        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '不是產業人才投資方案也不是委訓單位
            If LID.Value <> 2 Then
                If KindEngage.SelectedValue = 2 Then '如果是外聘
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

        Dim bt_openZip1_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code, AddressZIPB3, hidAddressZIP6W, TBCity, Address)
        bt_openZip1.Attributes.Add("onclick", bt_openZip1_Attr_VAL)

        Dim bt_openZip2_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code1, WorkZIPB3, hidWorkZIP6W, TBCity1, WorkAddr)
        bt_openZip2.Attributes.Add("onclick", bt_openZip2_Attr_VAL)

        Button1.Attributes("onclick") = "javascript:return chkdata();"
        Button3.Attributes("onclick") = "wopen('TC_01_007_copy.aspx?IDNO='+document.getElementById('IDNO').value,'copy',450,400,1);"

        IDNO.Attributes("onBlur") = "return SexChoice();"
        Sex.Attributes("onclick") = "return SexChoice();"
        KindID.Attributes("onclick") = "return KindIDChoice();"
        R2.Attributes("onclick") = "return change();"
        R1.Attributes("onclick") = "return change2();"

        If Not IsPostBack Then
            '有可能重複2次，所以加入判斷是否為空值
            If ProLicense.Text = "" Then
                ProLicense.Text = hid_PLMsgX1.Value
                ProLicense.ForeColor = ColorTranslator.FromHtml("#666666")
            End If
        End If
        ProLicense.Attributes("onfocus") = "PL_focusState();"
        ProLicense.Attributes("onblur") = "PL_focusState();"

    End Sub

    Sub AddItem()
        tr_techtype12.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投
            tr_techtype12.Visible = True '顯示 '講師 助教 (類別) 
        End If

        tr_techtype34.Visible = False
        If TIMS.Cst_TPlanID06Plan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '2018 add 自辦在職類別 (專用)  教師 / 第二教師 
            tr_techtype34.Visible = True
        End If

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
        Dim dr As DataRow = Nothing

        Dim sql As String = ""
        sql &= " SELECT a.*"
        sql &= " ,c.OrgName"
        sql &= " ,d.TrainID"
        sql &= " ,d.TrainName"
        sql &= " ,d.JobID"
        sql &= " ,d.JobName"
        'sql &= " ,ISNULL(d.TrainID,d.JobID) TrainID"
        'sql &= " ,ISNULL(d.TrainName,d.JobName) TrainName"
        'sql &= " ,ISNULL(d.TrainID,d.JobID) JobID"
        'sql &= " ,ISNULL(d.TrainName,d.JobName) JobName"
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
        ViewState("TID") = Convert.ToString(dr("TeacherID")) '暫存老師代碼-?
        TeachCName.Text = dr("TeachCName").ToString
        TeachEName.Text = dr("TeachEName").ToString
        IDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
        If Not IsDBNull(dr("Birthday")) Then
            birthday.Text = FormatDateTime(dr("Birthday"), 2)
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

        Dim v_TB_career_id_txt As String = ""
        If Convert.ToString(dr("TrainID")) <> "" Then
            v_TB_career_id_txt = "[" & Convert.ToString(dr("TrainID")) & "]" & Convert.ToString(dr("TrainName"))
        End If
        If v_TB_career_id_txt = "" AndAlso Convert.ToString(dr("JobID")) <> "" Then
            v_TB_career_id_txt = "[" & Convert.ToString(dr("JobID")) & "]" & Convert.ToString(dr("JobName"))
        End If
        TB_career_id.Text = v_TB_career_id_txt
        trainValue.Value = Convert.ToString(dr("TMID"))
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            jobValue.Value = Convert.ToString(dr("TMID"))
        End If

        If LID.Value <= 1 Then '如果是分署(中心)或署(局)
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 _
                OrElse dr("KindEngage").ToString <> "2" Then   '如果是內聘或是產學訓
                Common.SetListItem(IVID, dr("IVID"))
            Else                                               '如果是外聘
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
            If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '企訓專用
                Invest1.Text = dr("Invest").ToString
            End If
        End If

        Common.SetListItem(KindEngage, dr("KindEngage"))

        If tr_techtype12.Visible Then
            '產投(顯示) 才存取此功能
            cb_techtype1.Checked = ($"{dr("TechType1")}" = "Y")
            cb_techtype2.Checked = ($"{dr("TechType2")}" = "Y")
        End If

        If tr_techtype34.Visible Then
            '2018 add 自辦在職師資類別
            cb_techtype3.Checked = ($"{dr("TechType3")}" = "Y")
            cb_techtype4.Checked = ($"{dr("TechType4")}" = "Y")
        End If

        If LID.Value <= 1 Then
            Dim v_KindEngage As String = TIMS.GetListValue(KindEngage)
            KindID = TIMS.Get_KindOfTeacher(KindID, v_KindEngage, "", objconn)
            Common.SetListItem(KindID, dr("KindID").ToString)
        End If
        Common.SetListItem(DegreeID, dr("DegreeID"))
        Common.SetListItem(GraduateStatus, dr("GraduateStatus"))
        SchoolName.Text = $"{dr("SchoolName")}" ''dr("SchoolName").ToString
        Department.Text = $"{dr("Department")}" ''dr("Department").ToString
        Phone.Text = $"{dr("Phone")}" ''dr("Phone").ToString
        Mobile.Text = $"{dr("Mobile")}" '' dr("Mobile").ToString
        Email.Text = $"{dr("Email")}" ''dr("Email").ToString
        TBCity.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("AddressZip")), Convert.ToString(dr("AddressZIP6W")))
        city_code.Value = $"{dr("AddressZip")}" 'Convert.ToString(dr("AddressZip"))
        hidAddressZIP6W.Value = $"{dr("AddressZIP6W")}" 'Convert.ToString(dr("AddressZIP6W"))
        AddressZIPB3.Value = TIMS.GetZIPCODEB3(hidAddressZIP6W.Value)
        Address.Text = $"{dr("Address")}" 'dr("Address").ToString
        WorkOrg.Text = dr("WorkOrg")
        ExpYears.Text = dr("ExpYears").ToString
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用
            Invest2.Text = $"{dr("Invest")}" 'dr("Invest").ToString
        End If
        ServDept.Text = $"{dr("ServDept")}" 'dr("ServDept").ToString
        WorkPhone.Text = $"{dr("WorkPhone")}" 'dr("WorkPhone").ToString
        Fax.Text = $"{dr("Fax")}" 'dr("Fax").ToString
        TBCity1.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("WorkZip")), Convert.ToString(dr("WorkZIP6W")))
        city_code1.Value = $"{dr("WorkZip")}" 'Convert.ToString(dr("WorkZip"))
        hidWorkZIP6W.Value = $"{dr("WorkZIP6W")}" 'Convert.ToString(dr("WorkZIP6W"))
        WorkZIPB3.Value = TIMS.GetZIPCODEB3(dr("WorkZIP6W"))
        WorkAddr.Text = $"{dr("WorkAddr")}" 'dr("WorkAddr").ToString

        ExpUnit1.Text = dr("ExpUnit1").ToString
        If IsDate(dr("ExpSDate1")) Then
            ExpSDate1.Text = FormatDateTime(dr("ExpSDate1"), 2)
        End If
        If IsDate(dr("ExpEDate1")) Then
            ExpEDate1.Text = FormatDateTime(dr("ExpEDate1"), 2)
        End If
        ExpYears1.Text = dr("ExpYears1").ToString

        ExpUnit2.Text = dr("ExpUnit2").ToString
        If IsDate(dr("ExpSDate2")) Then
            ExpSDate2.Text = FormatDateTime(dr("ExpSDate2"), 2)
        End If
        If IsDate(dr("ExpEDate2")) Then
            ExpEDate2.Text = FormatDateTime(dr("ExpEDate2"), 2)
        End If
        ExpYears2.Text = dr("ExpYears2").ToString

        ExpUnit3.Text = dr("ExpUnit3").ToString
        If IsDate(dr("ExpSDate3")) Then
            ExpSDate3.Text = FormatDateTime(dr("ExpSDate3"), 2)
        End If
        If IsDate(dr("ExpEDate3")) Then
            ExpEDate3.Text = FormatDateTime(dr("ExpEDate3"), 2)
        End If
        ExpYears3.Text = dr("ExpYears3").ToString

        Common.SetListItem(ExpMonths, Convert.ToString(dr("ExpMonths")))
        Common.SetListItem(ExpMonths1, Convert.ToString(dr("ExpMonths1")))
        Common.SetListItem(ExpMonths2, Convert.ToString(dr("ExpMonths2")))
        Common.SetListItem(ExpMonths3, Convert.ToString(dr("ExpMonths3")))

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
            Specialty1.Text = TIMS.Get_Substr1(Specialty1.Text, 250)
            Specialty2.Text = TIMS.Get_Substr1(Specialty2.Text, 100)
            Specialty3.Text = TIMS.Get_Substr1(Specialty3.Text, 100)
            Specialty4.Text = TIMS.Get_Substr1(Specialty4.Text, 100)
            Specialty5.Text = TIMS.Get_Substr1(Specialty5.Text, 100)
        End If

        TransBook.Text = dr("TransBook").ToString

        If Convert.ToString(dr("ProLicense")) <> "" Then
            ProLicense.Text = Convert.ToString(dr("ProLicense"))
            ProLicense.ForeColor = ColorTranslator.FromHtml("#000000")
        Else
            ProLicense.Text = hid_PLMsgX1.Value
            ProLicense.ForeColor = ColorTranslator.FromHtml("#666666")
        End If
        '排課使用
        Common.SetListItem(WorkStatus, dr("WorkStatus"))

        If Convert.ToString(dr("PassPortNO")) <> "" Then
            Common.SetListItem(PassPortNO, dr("PassPortNO"))
        End If


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
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        '資料檢核 若有錯誤 則離開
        'Dim vERRMSG As String = ""
        If IDNO.Text = "" Then
            vERRMSG = "身分證號碼有誤，請確認!!"
        End If
        If IDNO.Text.IndexOf("*") > -1 Then
            vERRMSG = "身分證號碼有誤，請確認!!"
        End If
        Select Case PassPortNO.SelectedValue
            Case "1" '本國
                If Not TIMS.CheckIDNO(IDNO.Text) Then
                    vERRMSG = "身分證號碼有誤，請確認!!"
                End If
            Case "2" '外藉
                Dim nsIDNO As String = IDNO.Text
                '2:居留證 4:居留證2021
                Dim flag2 As Boolean = TIMS.CheckIDNO2(nsIDNO, 2)
                Dim flag4 As Boolean = TIMS.CheckIDNO2(nsIDNO, 4)
                If Not flag2 AndAlso Not flag4 Then
                    vERRMSG = "身分別為外藉，居留證號有誤，請確認!!"
                End If
            Case Else

        End Select
        If vERRMSG <> "" Then Exit Sub 'Return False

        jobValue.Value = TIMS.ClearSQM(jobValue.Value)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If jobValue.Value = "" AndAlso trainValue.Value = "" Then
                vERRMSG = "主要職類為必填，請重新確認!!"
            End If
        Else
            If trainValue.Value = "" Then
                vERRMSG = "主要職類為必填，請重新確認!!"
            End If
        End If
        If vERRMSG <> "" Then Exit Sub 'Return False

        If WorkOrg.Text = "" Then
            '服務單位名稱
            vERRMSG = "服務單位名稱為必填，請重新確認!!"
            Exit Sub 'Return False
        End If

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Dim dr As DataRow = Nothing
        '驗證IDNO 是否已建立
        Select Case Proecess.Text
            Case Cst_ProeTxt_新增
                Dim pmsx1 As New Hashtable From {{"RID", RIDValue.Value}, {"IDNO", IDNO.Text}}
                Dim sqlx1 As String = ""
                sqlx1 &= " SELECT 'x' FROM TEACH_TEACHERINFO"
                sqlx1 &= " WHERE RID=@RID and IDNO=@IDNO"
                Dim dt As DataTable = DbAccess.GetDataTable(sqlx1, objconn, pmsx1)
                If dt.Rows.Count > 0 Then
                    vERRMSG = Cst_msg2
                    Exit Sub 'Return False 'Common.MessageBox(Me, Cst_msg2)                        Exit Function
                End If
                'dr = DbAccess.GetOneRow(sql, objconn)
                'If Not dr Is Nothing Then
                '    If CInt(dr("CNT").ToString) > 0 Then
                '        vERRMSG = Cst_msg2
                '        Return False 'Common.MessageBox(Me, Cst_msg2)                        Exit Function
                '    End If
                'End If

        End Select

        If vERRMSG <> "" Then Exit Sub 'Return False
        If TIMS.LENB(TeacherID.Text) > 10 Then
            vERRMSG += Cst_msg4 & vbCrLf
        End If
        If TeachCName.MaxLength > 0 AndAlso TIMS.LENB(TeachCName.Text) > TeachCName.MaxLength Then
            vERRMSG += "講師姓名 長度最多" & TeachCName.MaxLength & "碼" & vbCrLf
        End If
        If TeachEName.MaxLength > 0 AndAlso TIMS.LENB(TeachEName.Text) > TeachEName.MaxLength Then
            vERRMSG += "講師英文姓名 長度最多" & TeachEName.MaxLength & "碼" & vbCrLf
        End If
        If vERRMSG <> "" Then Exit Sub

        If tr_techtype12.Visible Then
            Dim iTechType As Integer = 0
            '產投(顯示) 才存取此功能
            iTechType = 0
            If cb_techtype1.Checked = True Then
                iTechType += 1
            End If
            If cb_techtype2.Checked = True Then
                iTechType += 1
            End If

            If iTechType = 0 Then
                '至少要有一項
                vERRMSG += Cst_msg8 & vbCrLf
            End If
        End If

        If tr_techtype34.Visible Then
            '2018 add 自辦在職師資類別輸入檢核
            If Not cb_techtype3.Checked AndAlso Not cb_techtype4.Checked Then
                vERRMSG += "類別，請選擇教師或第二教師!!" & vbCrLf
            End If
        End If

        Specialty1.Text = TIMS.ClearSQM(Specialty1.Text)
        Specialty2.Text = TIMS.ClearSQM(Specialty2.Text)
        Specialty3.Text = TIMS.ClearSQM(Specialty3.Text)
        Specialty4.Text = TIMS.ClearSQM(Specialty4.Text)
        Specialty5.Text = TIMS.ClearSQM(Specialty5.Text)
        If Specialty5.Text <> "" AndAlso Specialty4.Text = "" Then
            vERRMSG &= "專長5有填, 專長4不可為空" & vbCrLf
        End If
        If Specialty4.Text <> "" AndAlso Specialty3.Text = "" Then
            vERRMSG &= "專長4有填, 專長3不可為空" & vbCrLf
        End If
        If Specialty3.Text <> "" AndAlso Specialty2.Text = "" Then
            vERRMSG &= "專長3有填, 專長2不可為空" & vbCrLf
        End If
        If Specialty2.Text <> "" AndAlso Specialty1.Text = "" Then
            vERRMSG &= "專長2有填, 專長1不可為空" & vbCrLf
        End If
        If vERRMSG <> "" Then Rst = False

        Dim vSpecialty1 As String = TIMS.Get_Substr1(Specialty1.Text, 250)
        Dim vSpecialty2 As String = TIMS.Get_Substr1(Specialty2.Text, 100)
        Dim vSpecialty3 As String = TIMS.Get_Substr1(Specialty3.Text, 100)
        Dim vSpecialty4 As String = TIMS.Get_Substr1(Specialty4.Text, 100)
        Dim vSpecialty5 As String = TIMS.Get_Substr1(Specialty5.Text, 100)
        If vSpecialty1 <> Specialty1.Text Then
            vERRMSG &= "專長1，字串長度不可超過250!" & vbCrLf
        End If
        If vSpecialty2 <> Specialty2.Text Then
            vERRMSG &= "專長2，字串長度不可超過100!" & vbCrLf
        End If
        If vSpecialty3 <> Specialty3.Text Then
            vERRMSG &= "專長3，字串長度不可超過100!" & vbCrLf
        End If
        If vSpecialty4 <> Specialty4.Text Then
            vERRMSG &= "專長4，字串長度不可超過100!" & vbCrLf
        End If
        If vSpecialty5 <> Specialty5.Text Then
            vERRMSG &= "專長5，字串長度不可超過100!" & vbCrLf
        End If

        If vERRMSG <> "" Then Rst = False
        'Return Rst
    End Sub

    '儲存
    Sub SaveData1(ByVal double_flag As Boolean)
        'double_flag'講師代碼重複 (true:(重複)允許儲存 false:沒有重複)
        Dim rqProecess As String = TIMS.ClearSQM(Request("proecess"))
        Dim rqTechID As String = TIMS.ClearSQM(Request("serial"))

        Mobile.Text = TIMS.ClearSQM(Mobile.Text)
        Fax.Text = TIMS.ClearSQM(Fax.Text)
        TransBook.Text = TIMS.ClearSQM(TransBook.Text)
        ProLicense.Text = TIMS.ClearSQM(ProLicense.Text)

        Specialty1.Text = TIMS.ClearSQM(Specialty1.Text)
        Specialty2.Text = TIMS.ClearSQM(Specialty2.Text)
        Specialty3.Text = TIMS.ClearSQM(Specialty3.Text)
        Specialty4.Text = TIMS.ClearSQM(Specialty4.Text)
        Specialty5.Text = TIMS.ClearSQM(Specialty5.Text)

        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Specialty1.Text = TIMS.Get_Substr1(Specialty1.Text, 250)
            Specialty2.Text = TIMS.Get_Substr1(Specialty2.Text, 100)
            Specialty3.Text = TIMS.Get_Substr1(Specialty3.Text, 100)
            Specialty4.Text = TIMS.Get_Substr1(Specialty4.Text, 100)
            Specialty5.Text = TIMS.Get_Substr1(Specialty5.Text, 100)
        End If

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

        'sql = "SELECT * FROM TEACH_TEACHERINFO WHERE TechID='" & Request("serial") & "'"
        Dim iTECHID As Integer = 0
        Select Case Proecess.Text
            Case Cst_ProeTxt_新增
                iTECHID = DbAccess.GetNewId(objconn, "TEACH_TEACHERINFO_TECHID_SEQ,TEACH_TEACHERINFO,TECHID")
                sql = "SELECT * FROM TEACH_TEACHERINFO WHERE 1<>1"
            Case Else 'Cst_修改
                iTECHID = Val(rqTechID)
                sql = "SELECT * FROM TEACH_TEACHERINFO WHERE TechID='" & rqTechID & "'"
        End Select
        If iTECHID = 0 Then Exit Sub
        '2006/03/28 add conn by matt

        Dim s_TransType As String = TIMS.cst_TRANS_LOG_Update
        Dim s_TargetTable As String = "TEACH_TEACHERINFO"
        Dim s_FuncPath As String = "/TC/01/TC_01_007"
        Const cst_fWHERE As String = "TECHID={0}"
        Dim s_WHERE As String = ""
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then
            s_TransType = TIMS.cst_TRANS_LOG_Insert
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("TECHID") = iTECHID 'TEACH_TEACHERINFO_TECHID_SEQ
            dr("RID") = RIDValue.Value
        Else
            dr = dt.Rows(0)
            iTECHID = dr("TECHID")
        End If
        s_WHERE = String.Format(cst_fWHERE, iTECHID)
        dr("TeacherID") = TeacherID.Text
        dr("TeachCName") = TeachCName.Text
        dr("TeachEName") = If(TeachEName.Text = "", Convert.DBNull, TeachEName.Text)
        dr("Birthday") = If(birthday.Text = "", Convert.DBNull, birthday.Text)
        dr("IDNO") = TIMS.ChangeIDNO(IDNO.Text)

        Dim v_TMID As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            v_TMID = If(jobValue.Value <> "", jobValue.Value, trainValue.Value)
        Else
            v_TMID = If(trainValue.Value <> "", trainValue.Value, "")
        End If
        dr("TMID") = If(v_TMID <> "", v_TMID, Convert.DBNull)
        dr("DegreeID") = DegreeID.SelectedValue
        dr("SchoolName") = If(SchoolName.Text = "", Convert.DBNull, SchoolName.Text)
        dr("Department") = If(Department.Text = "", Convert.DBNull, Department.Text)
        Dim v_GraduateStatus As String = TIMS.GetListValue(GraduateStatus)
        dr("GraduateStatus") = v_GraduateStatus 'GraduateStatus.SelectedValue
        dr("Phone") = Phone.Text

        city_code.Value = TIMS.ClearSQM(city_code.Value)
        AddressZIPB3.Value = TIMS.ClearSQM(AddressZIPB3.Value)
        hidAddressZIP6W.Value = TIMS.GetZIPCODE6W(city_code.Value, AddressZIPB3.Value)
        Address.Text = TIMS.ClearSQM(Address.Text)
        dr("AddressZip") = If(city_code.Value <> "", city_code.Value, Convert.DBNull)
        dr("AddressZip6W") = If(hidAddressZIP6W.Value <> "", hidAddressZIP6W.Value, Convert.DBNull)
        dr("Address") = Address.Text

        dr("WorkOrg") = WorkOrg.Text

        Invest2.Text = TIMS.ClearSQM(Invest2.Text)
        Invest1.Text = TIMS.ClearSQM(Invest1.Text)
        If LID.Value = 2 Then   '委訓單位
            dr("IVID") = Convert.DBNull
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '企訓專用
                dr("Invest") = Invest2.Text
            Else
                dr("Invest") = Invest1.Text
            End If
        Else        '中心
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 _
                OrElse KindEngage.SelectedValue <> 2 Then '假如是產業人才投資方案或是內聘

                dr("IVID") = IVID.SelectedValue
                dr("Invest") = Convert.DBNull
            Else             '外聘
                'If IVID.SelectedValue <> -1 Or IVID.SelectedValue <> "" Then '職稱的選單選項是否為空值,不是的話就存選單選項
                If IVID.SelectedIndex <> 0 Then
                    dr("IVID") = IVID.SelectedValue
                    dr("Invest") = Convert.DBNull
                ElseIf Invest1.Text <> "" Then '職稱的textbox是否為空值,不是的話就存textbox
                    dr("IVID") = Convert.DBNull
                    dr("Invest") = Invest1.Text
                End If
            End If
        End If

        If tr_techtype12.Visible Then
            '產投(顯示) 才存取此功能
            dr("TechType1") = If(cb_techtype1.Checked, "Y", Convert.DBNull) '講師(類別)
            dr("TechType2") = If(cb_techtype2.Checked, "Y", Convert.DBNull) '助教(類別)
        End If

        If tr_techtype34.Visible Then
            '2018 add 自辦在職師資類別(顯示) 才能存取此功能
            dr("TechType3") = If(cb_techtype3.Checked, "Y", Convert.DBNull)
            dr("TechType4") = If(cb_techtype4.Checked, "Y", Convert.DBNull)
        End If

        dr("ExpYears") = If(ExpYears.Text = "", Convert.DBNull, ExpYears.Text)

        city_code1.Value = TIMS.ClearSQM(city_code1.Value)
        WorkZIPB3.Value = TIMS.ClearSQM(WorkZIPB3.Value)
        hidWorkZIP6W.Value = TIMS.GetZIPCODE6W(city_code1.Value, WorkZIPB3.Value)
        WorkAddr.Text = TIMS.ClearSQM(WorkAddr.Text)
        dr("WorkZip") = If(city_code1.Value <> "", city_code1.Value, Convert.DBNull)
        dr("WorkZIP6W") = If(hidWorkZIP6W.Value <> "", hidWorkZIP6W.Value, Convert.DBNull)
        dr("WorkAddr") = If(WorkAddr.Text <> "", WorkAddr.Text, Convert.DBNull)

        dr("WorkPhone") = If(WorkPhone.Text = "", Convert.DBNull, WorkPhone.Text)
        dr("ExpUnit1") = If(ExpUnit1.Text = "", Convert.DBNull, ExpUnit1.Text)
        dr("ExpSDate1") = If(ExpSDate1.Text = "", Convert.DBNull, ExpSDate1.Text)
        dr("ExpEDate1") = If(ExpEDate1.Text = "", Convert.DBNull, ExpEDate1.Text)
        dr("ExpYears1") = If(ExpYears1.Text = "", Convert.DBNull, ExpYears1.Text)

        dr("ExpUnit2") = If(ExpUnit2.Text = "", Convert.DBNull, ExpUnit2.Text)
        dr("ExpSDate2") = If(ExpSDate2.Text = "", Convert.DBNull, ExpSDate2.Text)
        dr("ExpEDate2") = If(ExpEDate2.Text = "", Convert.DBNull, ExpEDate2.Text)
        dr("ExpYears2") = If(ExpYears2.Text = "", Convert.DBNull, ExpYears2.Text)

        dr("ExpUnit3") = If(ExpUnit3.Text = "", Convert.DBNull, ExpUnit3.Text)
        dr("ExpSDate3") = If(ExpSDate3.Text = "", Convert.DBNull, ExpSDate3.Text)
        dr("ExpEDate3") = If(ExpEDate3.Text = "", Convert.DBNull, ExpEDate3.Text)
        dr("ExpYears3") = If(ExpYears3.Text = "", Convert.DBNull, ExpYears3.Text)

        dr("ExpMonths") = If(ExpMonths.SelectedIndex = 0, Convert.DBNull, ExpMonths.SelectedValue)
        dr("ExpMonths1") = If(ExpMonths1.SelectedIndex = 0, Convert.DBNull, ExpMonths1.SelectedValue)
        dr("ExpMonths2") = If(ExpMonths2.SelectedIndex = 0, Convert.DBNull, ExpMonths2.SelectedValue)
        dr("ExpMonths3") = If(ExpMonths3.SelectedIndex = 0, Convert.DBNull, ExpMonths3.SelectedValue)

        dr("Specialty1") = If(Specialty1.Text = "", Convert.DBNull, Specialty1.Text)
        dr("Specialty2") = If(Specialty2.Text = "", Convert.DBNull, Specialty2.Text)
        dr("Specialty3") = If(Specialty3.Text = "", Convert.DBNull, Specialty3.Text)
        dr("Specialty4") = If(Specialty4.Text = "", Convert.DBNull, Specialty4.Text)
        dr("Specialty5") = If(Specialty5.Text = "", Convert.DBNull, Specialty5.Text)
        If LID.Value = 2 Then
            'dr("KindID") = Convert.DBNull 
            '130是講師 2010/12/13 號改成委訓單位都可以看到師資別,但是不可以選,
            dr("KindID") = 130
        Else
            dr("KindID") = KindID.SelectedValue
        End If
        dr("KindEngage") = KindEngage.SelectedValue
        dr("WorkStatus") = WorkStatus.SelectedValue '排課使用
        dr("Sex") = Sex.SelectedValue
        dr("Mobile") = If(Mobile.Text = "", Convert.DBNull, Mobile.Text)
        dr("Email") = If(Email.Text = "", Convert.DBNull, Email.Text)
        dr("ServDept") = If(ServDept.Text = "", Convert.DBNull, ServDept.Text)
        dr("Fax") = If(Fax.Text = "", Convert.DBNull, Fax.Text)
        dr("TransBook") = If(TransBook.Text = "", Convert.DBNull, TransBook.Text)

        If ProLicense.Text = hid_PLMsgX1.Value Then
            '未輸入有效資訊!!
            ProLicense.Text = ""
        End If
        dr("ProLicense") = If(ProLicense.Text = "", Convert.DBNull, ProLicense.Text)

        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        Select Case PassPortNO.SelectedValue
            Case "1", "2"
                dr("PassPortNO") = PassPortNO.SelectedValue
            Case Else
                dr("PassPortNO") = "2"
        End Select

        Dim htPP As New Hashtable
        htPP.Clear()
        htPP.Add("TransType", s_TransType)
        htPP.Add("TargetTable", s_TargetTable)
        htPP.Add("FuncPath", s_FuncPath)
        htPP.Add("s_WHERE", s_WHERE)
        TIMS.SaveTRANSLOG(sm, objconn, dr, htPP)

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

        If Session("MySearchStr") Is Nothing Then
            Session("MySearchStr") = ViewState("MySearchStr")
        End If

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
        Common.RespWrite(Me, "window.location.href='TC_01_007.aspx?ID=" & MRqID & "';")
        Common.RespWrite(Me, "</script>")

        'Try
        'Catch ex As Exception
        '    Common.RespWrite(Me, "<script language=javascript>window.alert('" & Cst_msg7 & "');")
        '    Common.MessageBox(Me, ex.ToString)
        'End Try
    End Sub

    '存檔。
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim rqTechID As String = Request("serial")
        rqTechID = TIMS.ClearSQM(rqTechID)
        Dim sql As String = ""
        'Dim dt As DataTable = Nothing
        'Dim dr As DataRow = Nothing
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
                'sql = "SELECT COUNT(1) CNT FROM TEACH_TEACHERINFO WHERE RID='" & IIf(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value) & "' and IDNO='" & IDNO.Text & "' "
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
                sql = ""
                sql &= " SELECT * FROM TEACH_TEACHERINFO"
                sql &= " WHERE RID='" & IIf(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value) & "'"
                sql &= " AND TeacherID='" & TeacherID.Text & "'"
                sql &= " AND TechID<>'" & rqTechID & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    double_flag = True '講師代碼重複
                End If
        End Select

        Call SaveData1(double_flag)
    End Sub

    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
        'Session("MySearchStr") = ViewState("MySearchStr")
        If Session("MySearchStr") Is Nothing Then
            Session("MySearchStr") = ViewState("MySearchStr")
        End If
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect1(Me, "TC_01_007.aspx?ID=" & MRqID & "")
        'Common.RespWrite(Me, "window.location.href='TC_01_007.aspx?ID=" & MRqID & "';</script>")
    End Sub

    Private Sub KindEngage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KindEngage.SelectedIndexChanged
        If KindEngage.SelectedIndex <> 0 Then
            'Dim sql As String
            'Dim dt As DataTable
            ''動態產生師資別
            Dim v_KindEngage As String = TIMS.GetListValue(KindEngage)
            KindID = TIMS.Get_KindOfTeacher(KindID, v_KindEngage, "", objconn)
        Else
            KindID.Items.Clear()
            KindID.Items.Add(New ListItem("請選擇內外聘", 0))
        End If

        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 _
            AndAlso LID.Value <= 1 Then '如果不是產學訓而且是分署(中心)或署(局) 【TIMS計畫】

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
        If TechID.Value <> "" Then
            Call sCreate1(TechID.Value)
        End If
    End Sub

End Class
