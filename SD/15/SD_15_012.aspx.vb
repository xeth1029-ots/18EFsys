Partial Class SD_15_012
    Inherits AuthBasePage

    '【綜合查詢統計表】
    '實際開訓人次： 班級開訓，學員開訓後14天實際錄訓人數,且有選擇預算別(公務/就安/就保或公務(ECFA))
    '結訓人次：班級開訓，學員補助符合補助者沒有離退訓，只剩開結訓人數，且結訓日期已過今天,且有選擇預算別(公務/就安/就保或公務(ECFA))
    '撥款人次：班級開訓，學員補助符合補助者 沒有離退訓，只剩開結訓人數，且結訓日期已過今天,且有選擇預算別(公務/就安/就保或公務(ECFA))-學員經費撥款狀態：已撥款之人數

    '特殊身分
    'Const Cst_SPEIdentity As String = "'01','04','05','06','07','10','26','28','33','37','40'"
    'TIMS.cst_Identity28
    '取得搜尋範圍 (班級) (SQL WHERE)-MV_CLASS_1-'Batch\Dbt_20190313

    Dim ff3 As String = ""
    Dim sCJOB_UNKEY As String = ""
    Dim dtSHARECJOB As DataTable
    Dim dtIdentity As DataTable 'key_identity
    Dim dtZip As DataTable
    Dim flag_use_上課地址及教室 As Boolean = False

    Dim str_TIME_MSG_A As String = ""
    Dim DateSec1 As DateTime
    Dim DateSec2 As DateTime

    '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
    Dim fg_Work2026x02 As Boolean = False 'TIMS.SHOW_W2026x02(sm)

    Const cst_BUDGET97_N As String = "公務ECFA" '公務ECFA // 協助

    'Dim gGovClassT As String = "" '1/2/3
    Const cst_GovClassT1 As String = "1"
    Const cst_GovClassT2 As String = "2"
    Const cst_GovClassT3 As String = "3"

    'test 測試環境測試
    Dim flag_chktest As Boolean = False

    '2018 (政府政策性產業)
    'Dim gflag_SHOW_2019_1 As Boolean = False 'TIMS.SHOW_2019_1(sm)
    Dim gfg_SHOW_2025_1 As Boolean = False 'TIMS.SHOW_2025_1
    Dim gflag_SHOW_2019_2 As Boolean = False '課程分類
    Dim gflag_SHOW_APPSTAGE As Boolean = False 'APPSTAGE-產投才顯示
    Dim flagExp1 As Boolean = False
    'Dim ipSearchStr As String = ""
    Dim g_ErrSql As String = ""

    Const Cst_全部 As Integer = 0
    'Const Cst_政府政策性產業_108NOUSE As Integer = 1 '(108年之後不使用此欄)
    Const Cst_政府政策性產業_114NOUSE As Integer = 1 '(114年後不使用) (114年後(含))
    Const Cst_新南向政策 As Integer = 2 '--dd.KID22,dd.KNAME22 進階政策性產業類別

    Const Cst_轄區重點產業 As Integer = 3
    Const Cst_生產力40 As Integer = 4 '生產力4.0
    Const Cst_申請人次 As Integer = 5

    Const Cst_申請補助費 As Integer = 6
    Const Cst_核定人次 As Integer = 7
    Const Cst_核定補助費 As Integer = 8

    Const Cst_實際開訓人次 As Integer = 9
    Const Cst_實際開訓人次加總 As Integer = 10
    Const Cst_預估補助費 As Integer = 11

    Const Cst_預估補助費加總 As Integer = 12
    Const Cst_結訓人次 As Integer = 13
    Const Cst_撥款人次 As Integer = 14

    Const Cst_撥款補助費 As Integer = 15
    Const Cst_不預告訪視次數_實地抽訪 As Integer = 16 '訪視日期
    Const Cst_不預告訪視次數_電話抽訪 As Integer = 17 '訪視日期

    Const Cst_不預告訪視次數_視訊訪查 As Integer = 18 '【累計不預告視訊抽訪次數】、【視訊訪視日期】
    Const Cst_累積訪視異常次數 As Integer = 19
    Const Cst_累計訪視異常原因 As Integer = 20

    Const Cst_離訓人次 As Integer = 21
    Const Cst_退訓人次 As Integer = 22
    Const Cst_訓練時數 As Integer = 23

    'https://jira.turbotech.com.tw/browse/TIMSC-301
    Const Cst_固定費用總額 As Integer = 24
    Const Cst_固定費用單一人時成本 As Integer = 25
    Const Cst_人時成本超出原因說明 As Integer = 26

    Const Cst_材料費總額 As Integer = 27
    Const Cst_材料費占比 As Integer = 28
    Const Cst_超出材料費比率上限原因說明 As Integer = 29 'Const Cst_人時成本 As Integer = 30

    Const Cst_上課時間 As Integer = 30
    Const Cst_撥款日期 As Integer = 31
    Const Cst_統一編號 As Integer = 32

    Const Cst_立案縣市 As Integer = 33
    Const Cst_包班事業單位 As Integer = 34
    Const Cst_師資名單 As Integer = 35

    Const Cst_上課地址及教室 As Integer = 36
    Const Cst_包班事業單位保險證號 As Integer = 37
    Const Cst_包班事業單位統一編號 As Integer = 38

    Const Cst_公務ECFA性別人數 As Integer = 39 '公務ECFA // 協助
    Const cst_課程申請流水號 As Integer = 40
    Const cst_上架日期 As Integer = 41

    Const cst_開放報名結束日期 As Integer = 42
    'https://jira.turbotech.com.tw/browse/TIMSC-218
    Const cst_課程備註 As Integer = 43
    'https://jira.turbotech.com.tw/browse/TIMSC-259
    '術科時數"、"聯絡人"、"聯絡電話"、"是否停辦
    Const cst_術科時數 As Integer = 44

    Const cst_聯絡人 As Integer = 45
    Const cst_聯絡電話 As Integer = 46
    Const cst_是否停辦 As Integer = 47

    'Const cst_政策性產業課程可辦理班數 As Integer = 48
    Const cst_iCAP標章證號及效期 As Integer = 48
    Const Cst_各身分別撥款人次 As Integer = 49
    Const Cst_各身分別撥款補助費 As Integer = 50

    Const Cst_辦理方式 As Integer = 51
    Const Cst_實際開訓性別人數 As Integer = 52
    Const Cst_行政管理疏失重大異常狀況 As Integer = 53

    Const Cst_報名繳費方式 As Integer = 54 'ENTERSUPPLYSTYLE

#Region "Function"
    'GET_ExitCell(ChbExit)'匯出欄位 
    Sub GET_ExitCell(ByRef obj As ListControl) 'As ListControl
        obj.Items.Clear()

        Dim listStr1 As String = ""
        listStr1 = ""
        listStr1 &= "全部,政府政策性產業(114年後不使用),新南向政策"
        listStr1 &= ",轄區重點產業,生產力4.0,申請人次"
        listStr1 &= ",申請補助費,核定人次,核定補助費"
        listStr1 &= ",實際開訓人次,實際開訓人次加總,預估補助費"
        listStr1 &= ",預估補助費加總,結訓人次,撥款人次"
        listStr1 &= ",撥款補助費,不預告訪視次數-實地抽訪,不預告訪視次數-電話抽訪"
        listStr1 &= ",不預告訪視次數-視訊訪查,累積訪視異常次數,累計訪視異常原因"
        listStr1 &= ",離訓人次,退訓人次,訓練時數"
        'https://jira.turbotech.com.tw/browse/TIMSC-301
        listStr1 &= ",固定費用總額,固定費用單一人時成本,人時成本超出原因說明"
        listStr1 &= ",材料費總額,材料費占比,超出材料費比率上限原因說明"
        'listStr1 &= ",人時成本,上課時間,撥款日期"
        listStr1 &= ",上課時間,撥款日期,統一編號"
        listStr1 &= ",立案縣市,包班事業單位,師資名單"
        listStr1 &= ",上課地址及教室,包班事業單位保險證號,包班事業單位統一編號"

        listStr1 &= String.Concat(",", cst_BUDGET97_N, "性別人數", ",課程申請流水號,上架日期")
        'https://jira.turbotech.com.tw/browse/TIMSC-259
        listStr1 &= ",開放報名/結束日期,課程備註,術科時數"
        listStr1 &= ",聯絡人,聯絡電話,是否停辦"
        '政策性產業課程可辦理班數
        listStr1 &= ",iCAP標章證號及效期,各身分別撥款人次,各身分別撥款補助費"
        listStr1 &= ",辦理方式,實際開訓性別人數,行政管理疏失/重大異常狀況"
        listStr1 &= ",報名繳費方式"
        'listStr1 &= ""
        Dim aStr1 As String() = Split(listStr1, ",")

        With obj.Items
            For i As Integer = 0 To aStr1.Length - 1
                .Insert(i, New ListItem(aStr1(i), i))
            Next
        End With

    End Sub

    '是否鍵詞
    Function AddList(ByVal obj As ListControl) As ListControl
        With obj.Items
            .Clear()
            .Insert(0, New ListItem("不區分", "A"))
            .Insert(1, New ListItem("是", "Y"))
            .Insert(2, New ListItem("否", "N"))
        End With
        Return obj
    End Function

#End Region

#Region "NO USE"
    'Cst_人時成本
    'select ip.years,ip.distname
    ', pp.DefGovCost
    ',pp.DefStdCost
    ',pp.DefUnitCost
    ',pp.DefCenterCost
    ',pp.DefMainCost
    ',pp.TotalHours
    ',pp.TNum
    ',pp.THours,pp.TotalCost
    'from plan_planinfo pp
    'join view_plan ip on ip.planid =pp.planid 
    'where 1=1
    'and ip.tplanid ='28'
    'and ip.years>=2011

    '訓練機構選擇
    'Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim DistID1 As String
    '    Dim N As Integer
    '    Dim i As Integer
    '    Dim msg As String = ""
    '    If sm.UserInfo.DistID = "000" Then
    '        DistID1 = ""
    '        N = 0   '預設 N =0 表示沒有勾選轄區選項
    '        For i = 1 To Me.Distid.Items.Count - 1
    '            If Me.Distid.Items(i).Selected Then '假如有勾選
    '                N = N + 1  '計算轄區勾選選項的數目
    '                If N = 1 Then '如果是勾選一個選項
    '                    DistID1 = Convert.ToString(Me.Distid.Items(i).Value) '取得選項的值
    '                End If
    '                'If N = 2 Then '如果轄區勾選選項的數目=2
    '                '    msg += "只能選擇一個轄區!" & vbCrLf
    '                '    DistID1 = ""
    '                '    Exit For
    '                'End If
    '            End If
    '        Next
    '        If N = 0 Then '如果轄區選項沒有選
    '            msg += "請選擇轄區!" & vbCrLf
    '        End If
    '        If msg <> "" Then
    '            Common.MessageBox(Me, msg)
    '        End If
    '    End If
    'End Sub
#End Region

    Dim cblX1 As CheckBoxList = Nothing

    Dim sCaseYears As String = "" '(FLAG)--(停)登入年度有關--/與選擇年度有關 hid_ssYears.Value
    'Const cst_y2012 As String = "2012"
    'Const cst_y2013 As String = "2013"
    'Const cst_y2015 As String = "2015"
    'Const cst_y2017 As String = "2017"
    Const cst_y2018 As String = "2018"
    Const cst_y2019 As String = "2019"

    'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me)  '1:2017前 2:2017 3:2018 4:2019
    'Dim au As New cAUTH
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objConn) '開啟連線
        'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018 4:2019
        'iPYNum = TIMS.sUtl_GetPYNum(Me)
        DateSec1 = Now

        'test 測試環境測試
        flag_chktest = If(TIMS.sUtl_ChkTest(), True, False) '(測試環境中)

        '2018 (政府政策性產業) 'Dim gflag_SHOW_2019_1 As Boolean = TIMS.SHOW_2019_1(sm)
        'gflag_SHOW_2019_1 = TIMS.SHOW_2019_1(sm)
        '產投顯示 申請階段 -APPSTAGE
        If sm.UserInfo.TPlanID = "28" Then gflag_SHOW_APPSTAGE = True
        'Dim gflag_SHOW_2019_2 As Boolean = False '課程分類
        gflag_SHOW_2019_2 = (sm.UserInfo.TPlanID = "28")
        'gflag_SHOW_2019_28_1 = TIMS.SHOW_2019_28_1(sm)
        gfg_SHOW_2025_1 = TIMS.SHOW_2025_1(sm)
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        fg_Work2026x02 = TIMS.SHOW_W2026x02(sm)

        Dim sql As String = " SELECT * FROM dbo.VIEW_ZIPNAME WITH(NOLOCK) ORDER BY ZIPCODE"
        dtZip = DbAccess.GetDataTable(sql, objConn)

        Dim s_Identity1 As String = If(TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1, TIMS.Cst_Identity06_2019_11, TIMS.Cst_Identity28_2019_11)
        Dim sql2 As String = $" SELECT IDENTITYID,NAME FROM dbo.KEY_IDENTITY WITH(NOLOCK) WHERE IDENTITYID IN ({s_Identity1}) ORDER BY SORT28"
        dtIdentity = DbAccess.GetDataTable(sql2, objConn)

        '通俗職類-table-含命名
        dtSHARECJOB = TIMS.Get_SHARECJOBdtV(objConn)

        If Not IsPostBack Then
            'BtnOpenOrg1.Attributes("onclick") = "OpenOrg('" & sm.UserInfo.TPlanID & "');"
            Call sCreate1()
        End If

    End Sub

    '設定年度(某些選項與年度有關)
    Sub setSelYears1(ByVal rYear As String)
        Dim int_Years As Integer = Val(rYear) '2012 'int_Years
        'int_Years = CInt(rYear) 'CInt(sm.UserInfo.Years)
        'Dim sCaseYears As String= cst_y2012 '"2012" '2012:舊年度資訊 2013:新年度資訊
        'sCaseYears = cst_y2012 '"2012" '2012:舊年度資訊 2013:新年度資訊
        'If int_Years >= 2013 AndAlso int_Years <= 2014 Then sCaseYears = cst_y2013
        'If int_Years >= 2015 AndAlso int_Years <= 2016 Then sCaseYears = cst_y2015
        'If int_Years >= 2017 Then sCaseYears = cst_y2017
        'If int_Years >= 2018 Then sCaseYears = cst_y2018
        'If int_Years >= 2019 Then sCaseYears = cst_y2019
        'If iPYNum >= 4 Then sCaseYears = cst_y2019

        sCaseYears = cst_y2018
        If int_Years >= 2018 Then sCaseYears = cst_y2018
        If int_Years >= 2019 Then sCaseYears = cst_y2019
        hid_ssYears.Value = sCaseYears

        KID_4_TR.Visible = True
        'KID_17.Items.Clear()
        'KID_17_tr.Visible = False
        'KID_17.Visible = True
        'KID_19.Visible = False
        Hid_GovClassT.Value = cst_GovClassT1
        'Select Case hid_ssYears.Value 'sCaseYears
        '    Case cst_y2012
        '        TIMS.Get_KeyBusiness(KID_6, "05", objConn)
        '        TIMS.Get_KeyBusiness(KID_10, "02", objConn)
        '        TIMS.Get_KeyBusiness(KID_4, "03", objConn)
        '        GovClassName = TIMS.Get_GovClass(GovClassName, 1, objConn) '訓練業別
        '        'Get_GovClass(GovClassName, 1) '訓練業別
        '        'get_Key_BusID(KID_7, "04") '舊的不管
        '    Case cst_y2013 '新年度2013
        '        TIMS.Get_KeyBusiness(KID_6, "07", objConn)
        '        TIMS.Get_KeyBusiness(KID_10, "08", objConn)
        '        TIMS.Get_KeyBusiness(KID_4, "06", objConn)
        '        GovClassName = TIMS.Get_GovClass(GovClassName, 1, objConn) '訓練業別
        '        'Get_GovClass(GovClassName, 1) '訓練業別
        '        'get_Key_BusID(KID_7, "04") '舊的不管
        '    Case cst_y2015
        '        Hid_GovClassT.Value = cst_GovClassT2
        '        TIMS.Get_KeyBusiness(KID_6, "10", objConn) '新興產業
        '        TIMS.Get_KeyBusiness(KID_10, "11", objConn) '重點服務業
        '        TIMS.Get_KeyBusiness(KID_4, "09", objConn) '新興智慧型產業
        '        GovClassName = TIMS.Get_GovClass(GovClassName, 2, objConn) '訓練業別
        '    Case cst_y2017
        '        Hid_GovClassT.Value = cst_GovClassT2
        '        KID_4.Items.Clear()
        '        KID_4_TR.Visible = False
        '        KID_17_tr.Visible = True
        '        TIMS.Get_KeyBusiness(KID_6, "10", objConn) '6大新興產業
        '        TIMS.Get_KeyBusiness(KID_10, "16", objConn) '10大重點服務業(9項)
        '        TIMS.Get_KeyBusiness(KID_17, "17", objConn)
        '        GovClassName = TIMS.Get_GovClass(GovClassName, 2, objConn) '訓練業別
        '    Case cst_y2018
        '        Hid_GovClassT.Value = cst_GovClassT3
        '        KID_4.Items.Clear()
        '        KID_4_TR.Visible = False
        '        KID_17_tr.Visible = True
        '        KID_17.Visible = False
        '        KID_19.Visible = True
        '        TIMS.Get_KeyBusiness(KID_6, "10", objConn) '6大新興產業
        '        TIMS.Get_KeyBusiness(KID_10, "16", objConn) '10大重點服務業(9項)
        '        'TIMS.Get_KeyBusiness(KID_17, "17", objConn)
        '        TIMS.Get_KeyBusiness(KID_19, "19", objConn)
        '        GovClassName = TIMS.Get_GovClass(GovClassName, 3, objConn) '訓練業別
        'End Select

        Hid_GovClassT.Value = cst_GovClassT3
        KID_4.Items.Clear()
        KID_4_TR.Visible = False
        'KID_17_tr.Visible = True
        'KID_17.Visible = False
        'KID_19.Visible = True
        TIMS.Get_KeyBusiness(KID_6, "10", objConn) '6大新興產業
        TIMS.Get_KeyBusiness(KID_10, "16", objConn) '10大重點服務業(9項)
        'TIMS.Get_KeyBusiness(KID_17, "17", objConn)
        'TIMS.Get_KeyBusiness(KID_19, "19", objConn)
        TIMS.Get_KeyBusiness(KID_20, "20", objConn)
        TIMS.Get_KeyBusiness(KID_25, "25", objConn)
        GovClassName = TIMS.Get_GovClass(GovClassName, 3, objConn) '訓練業別

        '產業別鍵詞
        KID_6_hid.Value = "0"
        KID_10_hid.Value = "0"
        KID_4_hid.Value = "0"
        'KID_17_hid.Value = "0" 'KID_19_hid.Value = "0"
        KID_20_hid.Value = "0"
        KID_25_hid.Value = "0"
        KID_6.Attributes("onclick") = "SelectAll('KID_6','KID_6_hid');"
        KID_10.Attributes("onclick") = "SelectAll('KID_10','KID_10_hid');"
        KID_4.Attributes("onclick") = "SelectAll('KID_4','KID_4_hid');"
        'KID_17.Attributes("onclick") = "SelectAll('KID_17','KID_17_hid');" 'KID_19.Attributes("onclick") = "SelectAll('KID_19','KID_19_hid');"
        KID_20.Attributes("onclick") = "SelectAll('KID_20','KID_20_hid');" '
        KID_25.Attributes("onclick") = "SelectAll('KID_25','KID_25_hid');" '
    End Sub

    Sub sCreate1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        '申請階段
        cbl_AppStage = TIMS.Get_APPSTAGE2(cbl_AppStage)

        cblDistid = TIMS.Get_DistID(cblDistid)
        cblDistid.Items.Insert(0, New ListItem("全部", 0))

        cblDistid.Enabled = True
        If sm.UserInfo.DistID <> "000" Then
            'Distid.Attributes("onclick") = "alert('x');"
            'Distid.SelectedValue = sm.UserInfo.DistID
            Common.SetListItem(cblDistid, sm.UserInfo.DistID)
            cblDistid.Enabled = False
        End If

        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)
        Call setSelYears1(sm.UserInfo.Years) '設定年度

        '產投才顯示
        trPlanKind.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        '產投不顯示
        trPackageType.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, False, True)

        '沒有不區分
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            OrgPlanKind = TIMS.Get_RblOrgPlanKind(OrgPlanKind, objConn)
            OrgPlanKind.Items.Insert(0, New ListItem("全部", "A")) 'A/G/W
            Common.SetListItem(OrgPlanKind, "A")
        End If

        Dim dtCity As DataTable = TIMS.Get_dtCity(objConn)
        PackageType = TIMS.GetPackageType(PackageType, sm.UserInfo.TPlanID)
        Tcitycode = TIMS.Get_CityName(Tcitycode, dtCity) 'TIMS.dtNothing)
        Ocitycode = TIMS.Get_CityName(Ocitycode, dtCity) 'TIMS.dtNothing)

        Call TIMS.Get_ClassCatelog(CCID, objConn)    '課程職能
        CCID.Items.Insert(0, New ListItem("全部", 0))
        PointYN = AddList(PointYN)
        Apppass = AddList(Apppass)
        Endclass = AddList(Endclass)
        'Appmoney = AddList(Appmoney)
        Stopclass = AddList(Stopclass)
        Call GET_ExitCell(ChbExit) '匯出欄位 

        'onchange
        'yearlist.Attributes("onchange") = "return SelectedNotEmpty('yearlist');"
        DistHidden.Value = "0"
        PackageHidden.Value = "0"
        TcityHidden.Value = "0"
        OcityHidden.Value = "0"
        GovClassHidden.Value = "0"
        CCIDHidden.Value = "0"
        ChbExitHidden.Value = "0"
        cblDistid.Attributes("onclick") = "SelectAll('cblDistid','DistHidden');"
        Org.Attributes("onclick") = If((sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1), "javascript:openOrg('../../Common/LevOrg.aspx');", "javascript:openOrg('../../Common/LevOrg1.aspx');")
        PackageType.Attributes("onclick") = "SelectAll('PackageType','PackageHidden');"
        Tcitycode.Attributes("onclick") = "SelectAll('Tcitycode','TcityHidden');"
        Ocitycode.Attributes("onclick") = "SelectAll('Ocitycode','OcityHidden');"
        GovClassName.Attributes("onclick") = "SelectAll('GovClassName','GovClassHidden');"
        CCID.Attributes("onclick") = "SelectAll('CCID','CCIDHidden');"
        ChbExit.Attributes("onclick") = "SelectAll('ChbExit','ChbExitHidden');"

        'Plankind.SelectedIndex = 0
        PointYN.SelectedIndex = 0
        Apppass.SelectedIndex = 0
        Endclass.SelectedIndex = 0
        Appmoney.SelectedIndex = 0
        Stopclass.SelectedIndex = 0

        '(V_DEPOT12)
        '課程分類 'KID12 'SELECT * FROM KEY_BUSINESS A WHERE 1=1 AND A.DEPID='12'
        HidcblDepot12.Value = "0"
        cblDepot12 = TIMS.Get_KeyBusiness(cblDepot12, "12", objConn) '課程分類
        cblDepot12.Attributes("onclick") = "SelectAll('cblDepot12','HidcblDepot12');"
    End Sub

    ''' <summary> 取得搜尋範圍 (班級) (SQL WHERE) </summary>
    ''' <returns></returns>
    Function Get_SearchStr_CC() As String
        Dim strSearch As String = "" '回傳值
        'cc 'pp  'ip,iz,iz3,iz4, iz2,ig ,vd06,vd10,vd04
        Dim DistID2 As String = ""
        Dim TCityCode2 As String = ""
        Dim OCityCode2 As String = ""
        Dim GovClassName2 As String = ""
        Dim PackageType2 As String = ""
        'Dim ExportStr As String = ""
        Dim CCID2 As String = ""

        '含未檢送研提資料
        '【提案匯總表】匯出時請排除，不匯出有勾選【未檢送資料】之班級。(預設)
        If Not CB_DataNotSent_SCH.Checked Then strSearch &= " and pp.DataNotSent IS NULL" & vbCrLf

        'Dim SeqNostr1 As String = "" '改為KID
        'Dim SeqNostr2 As String = "" '改為KID
        'Dim SeqNostr3 As String = "" '改為KID
        'Dim TMID As String = ""
        ' dt = New DataTable
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim fg_NG_ALL_NO_SEL1 As Boolean = (v_yearlist = "" AndAlso SDate1.Text <> "" AndAlso SDate2.Text <> "" AndAlso EDate1.Text <> "" AndAlso EDate2.Text <> "")
        If v_yearlist = "" AndAlso fg_NG_ALL_NO_SEL1 Then
            v_yearlist = sm.UserInfo.Years '(強制選擇登入年度)
            Common.SetListItem(yearlist, sm.UserInfo.Years)
        End If
        If v_yearlist <> "" Then strSearch &= " AND ip.YEARS='" & v_yearlist & "'" & vbCrLf

        '轄區
        DistID2 = ""
        If sm.UserInfo.DistID = "000" Then
            For i As Integer = 0 To cblDistid.Items.Count - 1
                If cblDistid.Items.Item(i).Selected = True AndAlso cblDistid.Items.Item(i).Value <> "" Then
                    If cblDistid.Items.Item(i).Text <> "全部" Then
                        DistID2 &= String.Concat(If(DistID2 <> "", ",", ""), "'", cblDistid.Items.Item(i).Value, "'")
                    End If
                End If
            Next
            If DistID2 <> "" Then
                'DistID2 = Left(DistID2, Len(DistID2) - 1)
                strSearch &= " and ip.DISTID IN (" & DistID2 & ")" & vbCrLf
            End If
        Else
            '依登入轄區
            strSearch &= " and ip.DISTID = '" & sm.UserInfo.DistID & "'" & vbCrLf
        End If

        Dim v_cbl_AppStage As String = TIMS.CombiSQM2IN(TIMS.GetCblValue(cbl_AppStage))
        If v_cbl_AppStage <> "" Then strSearch &= " and pp.APPSTAGE IN (" & v_cbl_AppStage & ")" & vbCrLf

        'If RIDValue.Value <> "" Then strSearch &= " and ar.RID = '" & RIDValue.Value & "'"
        'If OCIDValue1.Value <> "" Then strSearch &= " and cc.ocid = '" & OCIDValue1.Value & "'"
        'If Plankind.SelectedIndex <> 0 Then strSearch &= " and oo.OrgKind2 = '" & Plankind.SelectedValue & "'"

        '是產投計畫而且有打開計畫別功能
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso trPlanKind.Visible Then
            '有顯示才處理條件 
            Dim v_OrgPlanKind As String = TIMS.GetListValue(OrgPlanKind)
            Select Case v_OrgPlanKind'OrgPlanKind.SelectedValue 'A/G/W
                Case "A"
                Case "G", "W"
                    strSearch &= " and vr.OrgKind2= '" & v_OrgPlanKind & "'" & vbCrLf
                Case Else
                    strSearch &= " and 1<>1" & vbCrLf
            End Select
        End If

        '辦訓地縣市
        TCityCode2 = ""
        For i As Integer = 0 To Tcitycode.Items.Count - 1
            If Tcitycode.Items.Item(i).Selected = True AndAlso Tcitycode.Items.Item(i).Value <> "" Then
                If Tcitycode.Items.Item(i).Text <> "全部" Then
                    TCityCode2 &= String.Concat(If(TCityCode2 <> "", ",", ""), Tcitycode.Items.Item(i).Value)
                End If
            End If
        Next

        'Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        'If v_yearlist = "" Then
        '    v_yearlist = sm.UserInfo.Years
        '    Common.SetListItem(yearlist, sm.UserInfo.Years)
        'End If
        Dim intYears As Integer = Val(v_yearlist) 'yearlist.SelectedValue 'sm.UserInfo.Years
        If intYears <= 2010 Then
            If TCityCode2 <> "" Then
                strSearch &= " and iz.CTID IN (" & TCityCode2 & ")" & vbCrLf
            End If
        ElseIf intYears >= 2011 Then
            If TCityCode2 <> "" Then
                'TCityCode2 = Left(TCityCode2, Len(TCityCode2) - 1)
                strSearch &= " and (vtp.SPCTID IN (" & TCityCode2 & ") or vtp.TPCTID IN (" & TCityCode2 & "))" & vbCrLf
            End If
        End If
        '增加「訓練機構」搜尋條件，因分署會僅需匯出特定訓練單位之相關課程資料
        Dim s_ORGNAME As String = center.Text
        s_ORGNAME = TIMS.ClearSQM(s_ORGNAME)
        If s_ORGNAME <> "" Then strSearch &= " and oo.ORGNAME='" & s_ORGNAME & "'" & vbCrLf

        '包班種類 包班總類
        If trPackageType.Visible Then
            For i As Integer = 0 To PackageType.Items.Count - 1
                If PackageType.Items.Item(i).Selected = True AndAlso PackageType.Items.Item(i).Value <> "" Then
                    If PackageType.Items.Item(i).Text <> "全部" Then
                        If PackageType2 <> "" Then PackageType2 &= ","
                        PackageType2 &= "'" & PackageType.Items.Item(i).Value & "'"
                    End If
                End If
            Next
            If PackageType2 <> "" Then
                'PackageType2 = Left(PackageType2, Len(PackageType2) - 1)
                strSearch &= " and pp.PackageType IN (" & PackageType2 & ")" & vbCrLf
            End If
        End If

        '立案縣市 orgZipCode 立案地縣市
        OCityCode2 = ""
        For i As Integer = 0 To Ocitycode.Items.Count - 1
            If Ocitycode.Items.Item(i).Selected = True AndAlso Ocitycode.Items.Item(i).Value <> "" Then
                If Ocitycode.Items.Item(i).Text <> "全部" Then
                    If OCityCode2 <> "" Then OCityCode2 &= ","
                    OCityCode2 += Ocitycode.Items.Item(i).Value
                End If
            End If
        Next
        If OCityCode2 <> "" Then 'CTID2
            'iz2.CTID CTID2
            strSearch &= " and iz2.CTID IN (" & OCityCode2 & ")" & vbCrLf
        End If

        '課程分類 '訓練課程分類 KID12
        Dim sCblDepot12 As String = ""
        For i As Integer = 0 To cblDepot12.Items.Count - 1
            If cblDepot12.Items.Item(i).Selected _
                AndAlso cblDepot12.Items.Item(i).Value <> "" Then

                If cblDepot12.Items.Item(i).Text <> "全部" Then
                    If sCblDepot12 <> "" Then sCblDepot12 &= ","
                    sCblDepot12 += "'" & TIMS.ClearSQM(cblDepot12.Items.Item(i).Value) & "'"
                End If
            End If
        Next
        If sCblDepot12 <> "" Then 'KID12
            strSearch &= " and dd.KID12 IN (" & sCblDepot12 & ")" & vbCrLf
        End If

        '訓練業別
        For i As Integer = 0 To GovClassName.Items.Count - 1
            If GovClassName.Items.Item(i).Selected = True AndAlso GovClassName.Items.Item(i).Value <> "" Then
                If GovClassName.Items.Item(i).Text <> "全部" Then
                    If GovClassName2 <> "" Then GovClassName2 &= ","
                    GovClassName2 += "'" & GovClassName.Items.Item(i).Value & "'"
                End If
            End If
        Next
        If GovClassName2 <> "" Then
            'GovClassName2 = Left(GovClassName2, Len(GovClassName2) - 1)
            strSearch &= " and (1!=1" & vbCrLf
            'Hid_GovClassT.Value = cst_GovClassT2
            'Dim GovClassT As String = Hid_GovClassT.Value '1/2/3
            'Select Case Hid_GovClassT.Value 'GovClassT
            '    Case cst_GovClassT1 'GCODET1
            '        strSearch &= " OR convert(varchar,ig.GOVCLASS)+','+convert(varchar,ig.GCODE1) IN (" & GovClassName2 & ")" & vbCrLf
            '    Case cst_GovClassT2 'GCODET2
            '        strSearch &= " OR ig2.GCODE1 IN (" & GovClassName2 & ")" & vbCrLf
            '    Case cst_GovClassT3 'GCODET3
            '        strSearch &= " OR ig3.GCODE31 IN (" & GovClassName2 & ")" & vbCrLf
            'End Select

            Select Case Hid_GovClassT.Value 'GovClassT
                Case cst_GovClassT1 'GCODET1
                    'convert(varchar,ig.GOVCLASS)+','+convert(varchar,ig.GCODE1) GCODET1
                    strSearch &= " OR convert(varchar,ig.GOVCLASS)+','+convert(varchar,ig.GCODE1) IN (" & GovClassName2 & ")" & vbCrLf
                Case cst_GovClassT2 'GCODET2
                    'ig2.GCODE1 GCODET2
                    strSearch &= " OR ig2.GCODE1 IN (" & GovClassName2 & ")" & vbCrLf
                Case cst_GovClassT3 'GCODET3
                    'ig3.GCODE31 GCODET3
                    strSearch &= " OR ig3.GCODE31 IN (" & GovClassName2 & ")" & vbCrLf
            End Select
            strSearch &= " )" & vbCrLf
        End If

        '訓練職能
        For i As Integer = 0 To CCID.Items.Count - 1
            If CCID.Items.Item(i).Selected = True AndAlso CCID.Items.Item(i).Value <> "" Then
                If CCID.Items.Item(i).Text <> "全部" Then
                    If CCID2 <> "" Then CCID2 &= ","
                    CCID2 += CCID.Items.Item(i).Value
                End If
            End If
        Next
        If CCID2 <> "" Then
            strSearch &= " and pp.ClassCate IN (" & CCID2 & ")" & vbCrLf
        End If

        Dim tmpVal1 As String = ""
        '六大新興產業
        tmpVal1 = TIMS.GetCblValueIn(KID_6)
        If tmpVal1 <> "" Then strSearch &= " and dd.KID06 IN (" & tmpVal1 & ")" & vbCrLf
        '十大重點服務業
        tmpVal1 = TIMS.GetCblValueIn(KID_10)
        If tmpVal1 <> "" Then strSearch &= " and dd.KID10 IN (" & tmpVal1 & ")" & vbCrLf
        If KID_4_TR.Visible Then
            '四大新興智慧型產業
            tmpVal1 = TIMS.GetCblValueIn(KID_4)
            If tmpVal1 <> "" Then strSearch &= " and dd.KID04 IN (" & tmpVal1 & ")" & vbCrLf
        End If
        'If KID_17_tr.Visible Then
        '    If KID_17.Visible Then
        '        '政府政策性產業(108年之後不使用此欄)
        '        tmpVal1 = TIMS.GetCblValueIn(KID_17)
        '        If tmpVal1 <> "" Then strSearch &= " and dd.KID17 IN (" & tmpVal1 & ")" & vbCrLf
        '    End If
        '    If KID_19.Visible Then
        '        '政府政策性產業(108年之後不使用此欄)
        '        tmpVal1 = TIMS.GetCblValueIn(KID_19)
        '        If tmpVal1 <> "" Then strSearch &= " and dd.KID19 IN (" & tmpVal1 & ")" & vbCrLf
        '    End If
        'End If
        If KID_20.Visible Then
            '政府政策性產業(114年前(不含))
            tmpVal1 = TIMS.GetCblValueIn(KID_20)
            'If tmpVal1 <> "" Then strSearch &= String.Concat(" and dd.KID20 IN (", tmpVal1, ")", vbCrLf)
            If tmpVal1 <> "" Then strSearch &= TIMS.CombiSqlInlikeSta(tmpVal1, "dd.KID20")
        End If
        If KID_25.Visible Then
            '政府政策性產業(114年後(含))
            tmpVal1 = TIMS.GetCblValueIn(KID_25)
            'If tmpVal1 <> "" Then strSearch &= String.Concat(" and dd.KID25 IN (", tmpVal1, ")", vbCrLf)
            If tmpVal1 <> "" Then strSearch &= TIMS.CombiSqlInlikeSta(tmpVal1, "dd.KID25")
        End If
        Dim v_PointYN As String = TIMS.GetListValue(PointYN)
        If PointYN.SelectedIndex <> 0 AndAlso v_PointYN <> "" Then
            strSearch &= " and pp.PointYN = '" & v_PointYN & "'" & vbCrLf
        End If
        '是否核定(是)
        Dim v_Apppass As String = TIMS.GetListValue(Apppass)
        If Apppass.SelectedIndex <> 0 AndAlso v_Apppass <> "" Then
            strSearch &= " and pp.AppliedResult = '" & v_Apppass & "'" & vbCrLf
        End If

        Dim v_Endclass As String = TIMS.GetListValue(Endclass)
        If Endclass.SelectedIndex <> 0 AndAlso v_Endclass <> "" Then   '是否結訓
            If v_Endclass = "Y" Then
                strSearch &= " and cc.FTDate<GETDATE()"
            Else
                strSearch &= " and cc.FTDATE>=GETDATE()" & vbCrLf
            End If
        End If

        '是否停辦(否)
        Dim v_Stopclass As String = TIMS.GetListValue(Stopclass)
        If Stopclass.SelectedIndex <> 0 AndAlso v_Stopclass <> "" Then
            If v_Stopclass = "Y" Then
                strSearch &= " and cc.NOTOPEN='Y'" & vbCrLf '不開班
            Else
                strSearch &= " and cc.NOTOPEN='N'" & vbCrLf '開班
            End If
        End If

        If SDate1.Text <> "" Then
            'pp.STDate pSTDate
            strSearch &= " and pp.STDate >= " & TIMS.To_date(SDate1.Text) & vbCrLf '"','yyyy/MM/dd')" & vbCrLf
        End If
        If SDate2.Text <> "" Then
            'pp.STDate pSTDate
            strSearch &= " and pp.STDate <= " & TIMS.To_date(SDate2.Text) & vbCrLf 'SDate2.Text & "','yyyy/MM/dd')" & vbCrLf
        End If
        If EDate1.Text <> "" Then
            'pp.FDDate pFTDATE
            strSearch &= " and pp.FDDate >= " & TIMS.To_date(EDate1.Text) & vbCrLf 'EDate1.Text & "','yyyy/MM/dd')" & vbCrLf
        End If
        If EDate2.Text <> "" Then
            'pp.FDDate pFTDATE
            strSearch &= " and pp.FDDate <= " & TIMS.To_date(EDate2.Text) & vbCrLf 'EDate2.Text & "','yyyy/MM/dd')" & vbCrLf
        End If

        Return strSearch
    End Function

    ''' <summary> 取得搜尋範圍 (班級) (SQL WHERE)-CLASS_CLASSINFO-MV_CLASS_1-'Batch\Dbt_20190313 </summary>
    ''' <returns></returns>
    Function Get_SearchStr_MV() As String
        Dim strSearch As String = "" '回傳值
        'cc 'pp  'ip,iz,iz3,iz4, iz2,ig ,vd06,vd10,vd04
        Dim DistID2 As String = ""
        Dim TCityCode2 As String = ""
        Dim OCityCode2 As String = ""
        Dim GovClassName2 As String = ""
        Dim PackageType2 As String = ""
        'Dim ExportStr As String = ""
        Dim CCID2 As String = ""

        '含未檢送研提資料 '【提案匯總表】匯出時請排除，不匯出有勾選【未檢送資料】之班級。(預設)
        If Not CB_DataNotSent_SCH.Checked Then strSearch &= " and pp.DataNotSent IS NULL" & vbCrLf

        'Dim SeqNostr1 As String = "" '改為KID'Dim SeqNostr2 As String = "" '改為KID'Dim SeqNostr3 As String = "" '改為KID'Dim TMID As String = ""' dt = New DataTable
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim fg_NG_ALL_NO_SEL1 As Boolean = (v_yearlist = "" AndAlso SDate1.Text <> "" AndAlso SDate2.Text <> "" AndAlso EDate1.Text <> "" AndAlso EDate2.Text <> "")
        If v_yearlist = "" AndAlso fg_NG_ALL_NO_SEL1 Then
            v_yearlist = sm.UserInfo.Years '(強制選擇登入年度)
            Common.SetListItem(yearlist, sm.UserInfo.Years)
        End If
        If v_yearlist <> "" Then strSearch &= " AND pp.YEARS='" & v_yearlist & "'" & vbCrLf

        '轄區
        DistID2 = ""
        If sm.UserInfo.DistID = "000" Then
            For i As Integer = 0 To cblDistid.Items.Count - 1
                If cblDistid.Items.Item(i).Selected = True AndAlso cblDistid.Items.Item(i).Value <> "" Then
                    If cblDistid.Items.Item(i).Text <> "全部" Then
                        If DistID2 <> "" Then DistID2 += ","
                        DistID2 += "'" & cblDistid.Items.Item(i).Value & "'"
                    End If
                End If
            Next
            If DistID2 <> "" Then
                'DistID2 = Left(DistID2, Len(DistID2) - 1)
                strSearch &= " and pp.DISTID IN (" & DistID2 & ")" & vbCrLf
            End If
        Else
            '依登入轄區
            strSearch &= " and pp.DISTID = '" & sm.UserInfo.DistID & "'" & vbCrLf
        End If

        'If RIDValue.Value <> "" Then
        '    strSearch &= " and ar.RID = '" & RIDValue.Value & "'"
        'End If
        'If OCIDValue1.Value <> "" Then
        '    strSearch &= " and cc.ocid = '" & OCIDValue1.Value & "'"
        'End If
        'If Plankind.SelectedIndex <> 0 Then
        '    strSearch &= " and oo.OrgKind2 = '" & Plankind.SelectedValue & "'"
        'End If

        '是產投計畫而且有打開計畫別功能
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso trPlanKind.Visible Then
            '有顯示才處理條件 
            Select Case OrgPlanKind.SelectedValue 'A/G/W
                Case "A"
                Case "G", "W"
                    strSearch &= " and pp.OrgKind2= '" & OrgPlanKind.SelectedValue & "'" & vbCrLf
                Case Else
                    strSearch &= " and 1<>1" & vbCrLf
            End Select
        End If

        '辦訓地縣市
        TCityCode2 = ""
        For i As Integer = 0 To Tcitycode.Items.Count - 1
            If Tcitycode.Items.Item(i).Selected = True AndAlso Tcitycode.Items.Item(i).Value <> "" Then
                If Tcitycode.Items.Item(i).Text <> "全部" Then
                    If TCityCode2 <> "" Then TCityCode2 &= ","
                    TCityCode2 += Tcitycode.Items.Item(i).Value
                End If
            End If
        Next

        'Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        'If v_yearlist = "" Then
        '    v_yearlist = sm.UserInfo.Years
        '    Common.SetListItem(yearlist, sm.UserInfo.Years)
        'End If
        Dim intYears As Integer = Val(v_yearlist) 'yearlist.SelectedValue 'sm.UserInfo.Years
        If intYears <= 2010 Then
            If TCityCode2 <> "" Then
                'TCityCode2 = Left(TCityCode2, Len(TCityCode2) - 1)
                strSearch &= " and pp.CTID IN (" & TCityCode2 & ")" & vbCrLf
            End If
        ElseIf intYears >= 2011 Then
            If TCityCode2 <> "" Then
                'TCityCode2 = Left(TCityCode2, Len(TCityCode2) - 1)
                'strSearch &= " and (iz3.CTID IN (" & TCityCode2 & ") or iz4.CTID IN (" & TCityCode2 & "))" & vbCrLf
                strSearch &= " and (pp.SPCTID IN (" & TCityCode2 & ") or pp.TPCTID IN (" & TCityCode2 & "))" & vbCrLf
            End If
        End If

        '增加「訓練機構」搜尋條件，因分署會僅需匯出特定訓練單位之相關課程資料
        Dim s_ORGNAME As String = TIMS.ClearSQM(center.Text)
        If s_ORGNAME <> "" Then strSearch &= " and pp.ORGNAME='" & s_ORGNAME & "'" & vbCrLf

        '包班種類 包班總類
        If trPackageType.Visible Then
            For i As Integer = 0 To PackageType.Items.Count - 1
                If PackageType.Items.Item(i).Selected = True AndAlso PackageType.Items.Item(i).Value <> "" Then
                    If PackageType.Items.Item(i).Text <> "全部" Then
                        If PackageType2 <> "" Then PackageType2 &= ","
                        PackageType2 &= "'" & PackageType.Items.Item(i).Value & "'"
                    End If
                End If
            Next
            If PackageType2 <> "" Then
                'PackageType2 = Left(PackageType2, Len(PackageType2) - 1)
                strSearch &= " and pp.PackageType IN (" & PackageType2 & ")" & vbCrLf
            End If
        End If

        '立案縣市 orgZipCode 立案地縣市
        OCityCode2 = ""
        For i As Integer = 0 To Ocitycode.Items.Count - 1
            If Ocitycode.Items.Item(i).Selected = True AndAlso Ocitycode.Items.Item(i).Value <> "" Then
                If Ocitycode.Items.Item(i).Text <> "全部" Then
                    If OCityCode2 <> "" Then OCityCode2 &= ","
                    OCityCode2 += Ocitycode.Items.Item(i).Value
                End If
            End If
        Next
        If OCityCode2 <> "" Then 'CTID2
            'OCityCode2 = Left(OCityCode2, Len(OCityCode2) - 1)
            strSearch &= " and pp.CTID2 IN (" & OCityCode2 & ")" & vbCrLf
        End If

        '課程分類 '訓練課程分類 KID12
        Dim sCblDepot12 As String = ""
        For i As Integer = 0 To cblDepot12.Items.Count - 1
            If cblDepot12.Items.Item(i).Selected _
                AndAlso cblDepot12.Items.Item(i).Value <> "" Then

                If cblDepot12.Items.Item(i).Text <> "全部" Then
                    If sCblDepot12 <> "" Then sCblDepot12 &= ","
                    sCblDepot12 += "'" & TIMS.ClearSQM(cblDepot12.Items.Item(i).Value) & "'"
                End If
            End If
        Next
        If sCblDepot12 <> "" Then 'KID12
            'strSearch &= " and vd12.KID IN (" & sCblDepot12 & ")" & vbCrLf
            strSearch &= " and pp.KID12 IN (" & sCblDepot12 & ")" & vbCrLf
        End If

        '訓練業別
        For i As Integer = 0 To GovClassName.Items.Count - 1
            If GovClassName.Items.Item(i).Selected = True AndAlso GovClassName.Items.Item(i).Value <> "" Then
                If GovClassName.Items.Item(i).Text <> "全部" Then
                    If GovClassName2 <> "" Then GovClassName2 &= ","
                    GovClassName2 += "'" & GovClassName.Items.Item(i).Value & "'"
                End If
            End If
        Next
        If GovClassName2 <> "" Then
            'GovClassName2 = Left(GovClassName2, Len(GovClassName2) - 1)
            strSearch &= " and (1!=1" & vbCrLf
            'Hid_GovClassT.Value = cst_GovClassT2
            'Dim GovClassT As String = Hid_GovClassT.Value '1/2/3
            'Select Case Hid_GovClassT.Value 'GovClassT
            '    Case cst_GovClassT1 'GCODET1
            '        strSearch &= " OR convert(varchar,ig.GOVCLASS)+','+convert(varchar,ig.GCODE1) IN (" & GovClassName2 & ")" & vbCrLf
            '    Case cst_GovClassT2 'GCODET2
            '        strSearch &= " OR ig2.GCODE1 IN (" & GovClassName2 & ")" & vbCrLf
            '    Case cst_GovClassT3 'GCODET3
            '        strSearch &= " OR ig3.GCODE31 IN (" & GovClassName2 & ")" & vbCrLf
            'End Select

            Select Case Hid_GovClassT.Value 'GovClassT
                Case cst_GovClassT1 'GCODET1
                    strSearch &= " OR pp.GCODET1 IN (" & GovClassName2 & ")" & vbCrLf
                Case cst_GovClassT2 'GCODET2
                    strSearch &= " OR pp.GCODET2 IN (" & GovClassName2 & ")" & vbCrLf
                Case cst_GovClassT3 'GCODET3
                    strSearch &= " OR pp.GCODET3 IN (" & GovClassName2 & ")" & vbCrLf
            End Select
            strSearch &= " )" & vbCrLf
        End If

        '訓練職能
        For i As Integer = 0 To CCID.Items.Count - 1
            If CCID.Items.Item(i).Selected = True AndAlso CCID.Items.Item(i).Value <> "" Then
                If CCID.Items.Item(i).Text <> "全部" Then
                    If CCID2 <> "" Then CCID2 &= ","
                    CCID2 += CCID.Items.Item(i).Value
                End If
            End If
        Next
        If CCID2 <> "" Then
            'CCID2 = Left(CCID2, Len(CCID2) - 1)
            strSearch &= " and pp.ClassCate IN (" & CCID2 & ")" & vbCrLf
        End If

        Dim tmpVal1 As String = ""
        '六大新興產業
        tmpVal1 = TIMS.GetCblValueIn(KID_6)
        If tmpVal1 <> "" Then strSearch &= " and pp.KID06 IN (" & tmpVal1 & ")" & vbCrLf
        '十大重點服務業
        tmpVal1 = TIMS.GetCblValueIn(KID_10)
        If tmpVal1 <> "" Then strSearch &= " and pp.KID10 IN (" & tmpVal1 & ")" & vbCrLf
        If KID_4_TR.Visible Then
            '四大新興智慧型產業
            tmpVal1 = TIMS.GetCblValueIn(KID_4)
            If tmpVal1 <> "" Then strSearch &= " and pp.KID04 IN (" & tmpVal1 & ")" & vbCrLf
        End If
        'If KID_17_tr.Visible Then
        '    If KID_17.Visible Then
        '        '政府政策性產業(108年之後不使用此欄)
        '        tmpVal1 = TIMS.GetCblValueIn(KID_17)
        '        If tmpVal1 <> "" Then strSearch &= " and pp.KID17 IN (" & tmpVal1 & ")" & vbCrLf
        '    End If
        '    If KID_19.Visible Then
        '        '政府政策性產業(108年之後不使用此欄)
        '        tmpVal1 = TIMS.GetCblValueIn(KID_19)
        '        If tmpVal1 <> "" Then strSearch &= " and pp.KID19 IN (" & tmpVal1 & ")" & vbCrLf
        '    End If
        'End If
        If KID_20.Visible Then
            '政府政策性產業(114年前(不含))
            tmpVal1 = TIMS.GetCblValueIn(KID_20)
            'If tmpVal1 <> "" Then strSearch &= String.Concat(" and pp.KID20 IN (", tmpVal1, ")", vbCrLf)
            If tmpVal1 <> "" Then strSearch &= TIMS.CombiSqlInlikeSta(tmpVal1, "pp.KID20")
        End If
        If KID_25.Visible Then
            '政府政策性產業(114年後(含))
            tmpVal1 = TIMS.GetCblValueIn(KID_25)
            'If tmpVal1 <> "" Then strSearch &= String.Concat(" and pp.KID25 IN (", tmpVal1, ")", vbCrLf)
            If tmpVal1 <> "" Then strSearch &= TIMS.CombiSqlInlikeSta(tmpVal1, "pp.KID25")
        End If
        Dim v_PointYN As String = TIMS.GetListValue(PointYN)
        If PointYN.SelectedIndex <> 0 AndAlso v_PointYN <> "" Then
            strSearch &= " and pp.PointYN = '" & v_PointYN & "'" & vbCrLf
        End If
        '是否核定(是)
        Dim v_Apppass As String = TIMS.GetListValue(Apppass)
        If Apppass.SelectedIndex <> 0 AndAlso v_Apppass <> "" Then
            strSearch &= " and pp.AppliedResult = '" & v_Apppass & "'" & vbCrLf
        End If
        Dim v_Endclass As String = TIMS.GetListValue(Endclass)
        If Endclass.SelectedIndex <> 0 AndAlso v_Endclass <> "" Then   '是否結訓
            If v_Endclass = "Y" Then
                strSearch &= " and pp.FTDate<GETDATE()"
            Else
                strSearch &= " and pp.FTDate>=GETDATE()" & vbCrLf
            End If
        End If
        Dim v_Stopclass As String = TIMS.GetListValue(Stopclass)
        If Stopclass.SelectedIndex <> 0 AndAlso v_Stopclass <> "" Then
            If v_Stopclass = "Y" Then
                strSearch &= " and pp.NOTOPEN='Y'" & vbCrLf '不開班
            Else
                strSearch &= " and pp.NOTOPEN='N'" & vbCrLf '開班
            End If
        End If

        If SDate1.Text <> "" Then
            strSearch &= " and pp.pSTDate >= " & TIMS.To_date(SDate1.Text) & vbCrLf '"','yyyy/MM/dd')" & vbCrLf
        End If
        If SDate2.Text <> "" Then
            strSearch &= " and pp.pSTDate <= " & TIMS.To_date(SDate2.Text) & vbCrLf 'SDate2.Text & "','yyyy/MM/dd')" & vbCrLf
        End If
        If EDate1.Text <> "" Then
            strSearch &= " and pp.pFTDate >= " & TIMS.To_date(EDate1.Text) & vbCrLf 'EDate1.Text & "','yyyy/MM/dd')" & vbCrLf
        End If
        If EDate2.Text <> "" Then
            strSearch &= " and pp.pFTDate <= " & TIMS.To_date(EDate2.Text) & vbCrLf 'EDate2.Text & "','yyyy/MM/dd')" & vbCrLf
        End If

        Return strSearch
    End Function

    ''' <summary>  取得搜尋範圍 (班級)1 PLAN_PLANINFO -MV_CLASS_1-'Batch\Dbt_20190313 </summary>
    ''' <returns></returns>
    Function Get_SqlSchC1_CC() As String
        'Dim ssYears As String = yearlist.SelectedValue 'sm.UserInfo.Years
        Dim strSearchCls As String = Get_SearchStr_CC() '取得搜尋範圍  (班級)
        'Batch\Dbt_20190313-Get_MV_CLASS_dt

        '(YEARS)
        Dim sql As String = ""
        sql &= " SELECT pp.TMID ,concat(o1.typeid2,'-',o1.typeid2name) OrgTypeName" & vbCrLf
        sql &= " ,ip.DistName" & vbCrLf
        sql &= " ,vr.OrgPlanName2" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,cc.NOTOPEN" & vbCrLf
        sql &= " ,cc.STDATE ,cc.FTDATE" & vbCrLf
        sql &= " ,oo.ORGNAME ORGNAME" & vbCrLf
        sql &= " ,pp.CLASSNAME CLASSNAME" & vbCrLf
        sql &= " ,pp.FIRSTSORT FIRSTSORT" & vbCrLf
        sql &= " ,pp.CyclType" & vbCrLf
        sql &= " ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE" & vbCrLf
        sql &= " ,cc.ocid ClassID" & vbCrLf
        sql &= " ,pp.STDate pSTDate" & vbCrLf
        sql &= " ,pp.FDDate pFTDATE" & vbCrLf
        sql &= " ,pp.PSNO28" & vbCrLf
        sql &= " ,pp.ProTechHours" & vbCrLf
        sql &= " ,pp.ContactName" & vbCrLf
        sql &= " ,pp.ContactPhone" & vbCrLf
        sql &= " ,pp.ContactMobile" & vbCrLf
        sql &= " ,dbo.DECODE2(cc.NotOpen,'Y','是','否') NotOpenN" & vbCrLf
        'sql &= " ,format(cc.ONSHELLDATE,'yyyy/MM/dd HH:mm') ONSHELLDATE" & vbCrLf
        sql &= " ,cc.ONSHELLDATE" & vbCrLf
        sql &= " ,cc.SENTERDATE,cc.FENTERDATE" & vbCrLf
        sql &= " ,pv.memo8" & vbCrLf
        sql &= " ,pv.memo82" & vbCrLf
        sql &= " ,pp.FIXSUMCOST" & vbCrLf
        sql &= " ,pp.ACTHUMCOST" & vbCrLf
        sql &= " ,pp.FIXExceeDesc" & vbCrLf
        sql &= " ,pp.METSUMCOST" & vbCrLf
        sql &= " ,concat(ISNULL(pp.METCOSTPER,0),'%') METCOSTPER" & vbCrLf
        sql &= " ,pp.METExceeDesc" & vbCrLf
        sql &= " ,pp.PLANID" & vbCrLf
        sql &= " ,pp.SEQNO" & vbCrLf
        sql &= " ,oo.COMIDNO" & vbCrLf
        'oo.orgZipCode=iz2.ZipCode
        sql &= " ,iz2.CTName CTName2" & vbCrLf
        sql &= " ,case pp.PackageType when '1' then '非包班' when '2' then '企業包班' when '3' then '聯合企業包班' end PackageTypeN" & vbCrLf
        '課程分類
        sql &= " ,ig3.PNAME PKNAME12" & vbCrLf
        '轄區重點產業
        sql &= " ,dd.kname13" & vbCrLf
        sql &= " ,dd.kname15" & vbCrLf
        '生產力4.0
        sql &= " ,dd.kname14" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.KNAME06 ,ISNULL(dd.KNAME06 ,vd06.kname)) kname1" & vbCrLf '新興產業
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.KNAME10 ,ISNULL(dd.KNAME10 ,vd10.kname)) kname2" & vbCrLf '重點服務業
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.KNAME04 ,ISNULL(dd.KNAME04 ,vd04.kname)) kname3" & vbCrLf '新興智慧型產業
        'sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.KNAME17 ,NULL) KNAME17" & vbCrLf '政府政策性產業(108年之後不使用此欄)
        'sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.KNAME19 ,NULL) KNAME19" & vbCrLf '政府政策性產業(108年之後不使用此欄)
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.KNAME18 ,NULL) KNAME18" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.KID20 ,NULL) KID20" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.KID25 ,NULL) KID25" & vbCrLf
        '2019年啟用 work2019x01:2019 政府政策性產業
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME1 ,'無') D20KNAME1" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME2 ,'無') D20KNAME2" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME3 ,'無') D20KNAME3" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME4 ,'無') D20KNAME4" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME5 ,'無') D20KNAME5" & vbCrLf
        sql &= " ,dbo.DECODE2(dd.AppResult,'Y',dd.D20KNAME6 ,'無') D20KNAME6" & vbCrLf
        'dd.KID22,dd.KNAME22 進階政策性產業類別
        sql &= " ,dd.KID22,dd.KNAME22" & vbCrLf
        '2025 政府政策性產業 (產投)
        sql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8" & vbCrLf
        '2026 政府政策性產業 (產投)
        sql &= " ,dd.D26KNAME1,dd.D26KNAME2,dd.D26KNAME7,dd.D26KNAME3,dd.D26KNAME5,dd.D26KNAME4,dd.D26KNAME6,dd.D26KNAME8,dd.D26KNAME9"
        sql &= " ,ISNULL(ig.GOVCLASSN,ISNULL(ig2.GCODE2,ig3.GCODE2)) GCodeName" & vbCrLf 'GCodeName: 訓練業別編碼
        sql &= " ,ISNULL(ig.CNAME,ISNULL(ig2.CNAME,ig3.CNAME)) GCNAME" & vbCrLf
        sql &= " ,pp.CJOB_UNKEY" & vbCrLf
        sql &= " ,kc.CCName" & vbCrLf
        sql &= " ,case when CONVERT(INT, pp.PlanYear) <= 2010 then iz.CTName else vtp.SPCTName end AddressSciPTID" & vbCrLf
        sql &= " ,case when CONVERT(INT, pp.PlanYear) <= 2010 then iz.CTName else vtp.TPCTName end AddressTechPTID" & vbCrLf
        sql &= " ,ISNULL(pp.DefGovCost,0) ADefGovCost" & vbCrLf
        sql &= " ,ISNULL(pp.TNum,0) ATNum" & vbCrLf
        sql &= " ,ISNULL(pp.THours,0) THours" & vbCrLf
        sql &= " ,case when pp.AppliedResult='Y' then ISNULL(pp.DefGovCost,0) else 0 end DefGovCost" & vbCrLf
        sql &= " ,case when pp.AppliedResult='Y' then ISNULL(pp.TotalCost,0) else 0 end TotalCost" & vbCrLf
        sql &= " ,case when pp.AppliedResult='Y' then ISNULL(pp.TNum,0) else 0 end TNum" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_ONCLASS2(pp.PlanID,pp.ComIDNO,pp.SeqNo,'WEEKTIME') WEEKSTIME" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_BUSPACKAGE2(pp.PlanID,pp.ComIDNO,pp.SeqNo,'1') BusPackage" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_BUSPACKAGE2(pp.PlanID,pp.ComIDNO,pp.SeqNo,'2') BusPackage2" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_BUSPACKAGE2(pp.PlanID,pp.ComIDNO,pp.SeqNo,'3') BusPackage3" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_TEACHER2(pp.PlanID,pp.ComIDNO,pp.SeqNo,'1') PlanTeacher" & vbCrLf
        sql &= " ,VTP.S1PLACEID ,VTP.S2PLACEID ,VTP.T1PLACEID ,VTP.T2PLACEID ,VTP.SPPLACEID ,VTP.TPPLACEID
 ,VTP.S1PTID ,VTP.S2PTID ,VTP.T1PTID ,VTP.T2PTID ,VTP.SPPTID ,VTP.TPPTID
 ,VTP.S1PLACENAME ,VTP.S2PLACENAME ,VTP.T1PLACENAME ,VTP.T2PLACENAME ,VTP.SPPLACENAME ,VTP.TPPLACENAME
 ,VTP.S1ZIPCODE ,VTP.S1ADDRESS ,VTP.S1ZIP6W
 ,VTP.S2ZIPCODE ,VTP.S2ADDRESS ,VTP.S2ZIP6W
 ,VTP.T1ZIPCODE ,VTP.T1ADDRESS ,VTP.T1ZIP6W
 ,VTP.T2ZIPCODE ,VTP.T2ADDRESS ,VTP.T2ZIP6W
 ,VTP.SPZIPCODE ,VTP.SPADDRESS ,VTP.SPZIP6W
 ,VTP.TPZIPCODE ,VTP.TPADDRESS ,VTP.TPZIP6W
 "
        sql &= " ,pp.IsApprPaper" & vbCrLf
        sql &= " ,ip.TPLANID" & vbCrLf
        sql &= " ,ip.YEARS" & vbCrLf
        sql &= " ,ip.DISTID" & vbCrLf
        sql &= " ,vr.OrgKind2" & vbCrLf
        sql &= " ,iz.CTID" & vbCrLf
        sql &= " ,vtp.SPCTID" & vbCrLf
        sql &= " ,vtp.TPCTID" & vbCrLf
        sql &= " ,pp.PackageType" & vbCrLf
        sql &= " ,iz2.CTID CTID2" & vbCrLf
        sql &= " ,ig.PGcid GCODET1" & vbCrLf
        sql &= " ,ig2.GCODE1 GCODET2" & vbCrLf
        sql &= " ,ig3.GCODE31 GCODET3" & vbCrLf
        sql &= " ,dd.KID12" & vbCrLf
        sql &= " ,pp.ClassCate" & vbCrLf
        sql &= " ,dd.KID06" & vbCrLf
        sql &= " ,dd.KID10" & vbCrLf
        sql &= " ,dd.KID04" & vbCrLf
        'sql &= " ,dd.KID17" & vbCrLf
        'sql &= " ,dd.KID19" & vbCrLf
        sql &= " ,pp.PointYN" & vbCrLf
        sql &= " ,pp.AppliedResult" & vbCrLf
        'ICAPNUM-iCAP標章證號 -cst_iCAP標章證號及效期
        sql &= " ,pp.ICAPNUM" & vbCrLf
        sql &= " ,pp.iCAPMARKDATE" & vbCrLf 'sql &= " ,format(pp.iCAPMARKDATE,'yyyy/MM/dd') iCAPMARKDATE" & vbCrLf
        '辦理方式-遠距教學 'DISTANCE,dbo.FN_GET_DISTANCE(DISTANCE) DISTANCE_N
        sql &= " ,pp.DISTANCE" & vbCrLf

        sql &= " ,pp.OUTDOOR" & vbCrLf '課程內容有室外教學 室外教學課程
        sql &= " ,pv.REPORTE" & vbCrLf '報請主管機關核備
        sql &= " ,pp.DataNotSent" & vbCrLf '未檢送資料
        sql &= " ,pp.ENTERSUPPLYSTYLE" & vbCrLf '報名繳費方式
        sql &= " ,(select max(m.VERIFYDATE) VERIFYDATE from CLASS_MAJOR m WHERE m.OCID=cc.OCID) VERIFYDATE" & vbCrLf '為分署於該功能登打之經查核確認日期。
        'sql &= " ,dbo.FN_GET_ROC_YEAR(CONVERT(int,ip.YEARS)+0) YR1" & vbCrLf
        'sql &= " ,dbo.FN_GET_ROC_YEAR(CONVERT(int,ip.YEARS)+1) YR2" & vbCrLf
        'sql &= " ,dbo.FN_GET_ROC_YEAR(CONVERT(int,ip.YEARS)+2) YR3" & vbCrLf
        'sql &= " ,dbo.FN_GET_PRECLASS_PCNT1(pp.PLANID,pp.COMIDNO,pp.SEQNO,ip.YEARS,1) PCNT11" & vbCrLf
        'sql &= " ,dbo.FN_GET_PRECLASS_PCNT1(pp.PLANID,pp.COMIDNO,pp.SEQNO,ip.YEARS,2) PCNT12" & vbCrLf
        'sql &= " ,dbo.FN_GET_PRECLASS_PCNT1(pp.PLANID,pp.COMIDNO,pp.SEQNO,ip.YEARS,3) PCNT13" & vbCrLf
        sql &= " ,GETDATE() SCHUMDATE" & vbCrLf
        'sql &= " INTO MV_CLASS_1" & vbCrLf --Dbt_20190313
        sql &= " FROM dbo.KEY_ORGTYPE ky WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) on oo.orgkind=ky.orgTypeid" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO pp WITH(NOLOCK) on pp.comidno=oo.comidno" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc WITH(NOLOCK) on pp.ClassCate=kc.CCID" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip WITH(NOLOCK) on ip.planid=pp.planid" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME vr WITH(NOLOCK) on vr.RID=pp.RID" & vbCrLf
        sql &= " LEFT JOIN dbo.KEY_ORGTYPE1 o1 WITH(NOLOCK) on oo.OrgKind1=o1.OrgTypeID1" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_GOVCLASSCAST ig WITH(NOLOCK) on pp.GCID=ig.GCID" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST2 ig2 WITH(NOLOCK) on pp.GCID2=ig2.GCID2" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 ig3 WITH(NOLOCK) on pp.GCID3=ig3.GCID3" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_TRAINTYPE tt WITH(NOLOCK) on tt.TMID=pp.TMID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_DEPOT12 vd12 WITH(NOLOCK) on vd12.GCID2=pp.GCID2" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd WITH(NOLOCK) on dd.planid=pp.planid and dd.comidno=pp.comidno and dd.seqno=pp.seqno" & vbCrLf

        sql &= " LEFT JOIN dbo.VIEW_DEPOT10 vd06 on vd06.GCID2=pp.GCID2" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_DEPOT11 vd10 on vd10.GCID2=pp.GCID2" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_DEPOT09 vd04 on vd04.GCID2=pp.GCID2" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_VERREPORT pv WITH(NOLOCK) on pp.planid=pv.planid and pp.comidno=pv.comidno and pp.seqno=pv.seqno" & vbCrLf
        sql &= " LEFT JOIN dbo.CLASS_CLASSINFO cc WITH(NOLOCK) on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno" & vbCrLf
        sql &= " LEFT JOIN dbo.ID_CLASS f WITH(NOLOCK) on cc.CLSID=f.CLSID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz on pp.TaddressZip=iz.ZipCode" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz2 on oo.orgZipCode=iz2.ZipCode" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_TRAINPLACE vtp WITH(NOLOCK) on vtp.planid=pp.planid and vtp.comidno=pp.comidno and vtp.seqno=pp.seqno" & vbCrLf
        sql &= " WHERE pp.IsApprPaper ='Y'" & vbCrLf 'sql &= " and ip.TPLANID =@TPLANID" & vbCrLf 'sql &= " and ip.YEARS =@YEARS" & vbCrLf
        Dim s_pTPlanID1 As String = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, sm.UserInfo.TPlanID, "28")
        sql &= String.Format(" and ip.TPLANID ='{0}'", s_pTPlanID1) & vbCrLf

        '取得搜尋範圍  (班級)
        sql &= strSearchCls

        'MV_CLASS_1  'test 測試環境測試
        If (flag_chktest) Then
            TIMS.WriteLog(Me, String.Concat("##SD_15_012.aspx,", vbCrLf, ",Get_SqlSchC1_CC - PLAN_PLANINFO -MV_CLASS_1-'Batch\Dbt_20190313-sql:", vbCrLf, sql))
        End If

        'sql &= " and pp.Tplanid ='28' and pp.IsApprPaper ='Y' and ip.Years='2016' and ip.Distid='005'" & vbCrLf
        Return sql
    End Function

    ''' <summary> 取得搜尋範圍 (班級)1 CLASS_CLASSINFO-MV_CLASS_1 --Dbt_20190313 </summary>
    ''' <returns></returns>
    Function Get_SqlSchC1_MV() As String
        'Batch\Dbt_20190313
        'Dbt_20190313.vbproj
        'Dim ssYears As String=yearlist.SelectedValue 'sm.UserInfo.Years
        Dim strSearchCls As String = Get_SearchStr_MV() '取得搜尋範圍  (班級)

        Dim s_pTPlanID1 As String = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, sm.UserInfo.TPlanID, "28")
        Dim sql As String = " SELECT pp.* FROM dbo.MV_CLASS_1 pp WHERE 1=1" & vbCrLf
        sql &= $" AND pp.TPLANID='{s_pTPlanID1}'{vbCrLf}"
        '取得搜尋範圍  (班級)
        sql &= strSearchCls
        'sql &= " and pp.Tplanid ='28' and pp.IsApprPaper ='Y' and ip.Years='2016' and ip.Distid = '005'" & vbCrLf
        Return sql
    End Function

    ''' <summary> 取得搜尋範圍 (班級)2(不預告實地抽訪) CLASS_UNEXPECTVISITOR </summary>
    ''' <returns></returns>
    Function Get_SqlSchC2() As String
        Dim sql As String = ""
        sql &= " SELECT U.OCID" & vbCrLf
        '/*累計不預告實地抽訪次數*/ 'sql &= " ,COUNT(1) cuall" & vbCrLf
        sql &= " ,COUNT(CASE WHEN ISNULL(u.VISITWAY,'1')='1' THEN 1 END) CUALL" & vbCrLf
        '累計不預告視訊抽訪次數 視訊訪查 VW2CNT
        sql &= " ,COUNT(CASE WHEN u.VISITWAY='2' THEN 1 END) VW2CNT" & vbCrLf
        '/*累計不預告實地抽訪異常次數*/
        sql &= " ,COUNT(case when U.LItem1='2' AND ISNULL(u.VISITWAY,'1')='1' then 1 end) vitN" & vbCrLf
        sql &= " ,COUNT(case when U.LItem1='2' AND u.VISITWAY='2' then 1 end) vitTVN" & vbCrLf
        '/*累計不預告實地抽訪異常次數 累計訪視異常原因*/
        sql &= " ,COUNT(case when U.LItem2_2b LIKE '%01%' then 1 end) It22b01N" & vbCrLf
        sql &= " ,COUNT(case when U.LItem2_2b LIKE '%02%' then 1 end) It22b02N" & vbCrLf
        sql &= " ,COUNT(case when U.LItem2_2b LIKE '%03%' then 1 end) It22b03N" & vbCrLf
        sql &= " ,COUNT(case when U.LItem2_2b LIKE '%06%' then 1 end) It22b06N" & vbCrLf
        sql &= " ,COUNT(case when U.LItem2_2b LIKE '%04%' then 1 end) It22b04N" & vbCrLf
        sql &= " ,COUNT(case when U.LItem2_2b LIKE '%05%' then 1 end) It22b05N" & vbCrLf
        'sql &= " ,COUNT(case when U.LItem2_2b LIKE '%99%' then 1 end) It22b99N" & vbCrLf

        sql &= " ,dbo.FN_GET_UNEXPECTVISITOR(U.OCID,1) ItAPPLYDATE" & vbCrLf 'ItAPPLYDATE 實地訪視日期
        sql &= " ,dbo.FN_GET_UNEXPECTVISITOR(U.OCID,21) VW2APPLYDATE" & vbCrLf '視訊訪視日期 視訊訪查 VW2APPLYDATE

        sql &= " ,dbo.FN_GET_UNEXPECTVISITOR(U.OCID,2) It22b99NOTE" & vbCrLf  'IT22B99NOTE 累計訪視異常原因 其他
        sql &= " ,dbo.FN_GET_UNEXPECTVISITOR(U.OCID,5) LITEM23NOTE" & vbCrLf 'LITEM23NOTE 其他補充說明

        sql &= " FROM dbo.CLASS_UNEXPECTVISITOR U WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN WC1 cc ON cc.OCID =U.OCID" & vbCrLf
        sql &= " GROUP BY U.OCID" & vbCrLf
        Return sql
    End Function

    ''' <summary> 取得搜尋範圍 (班級)2(不預告電話抽訪) CLASS_UNEXPECTTEL </summary>
    ''' <returns></returns>
    Function Get_SqlSchC3() As String
        Dim sql As String = ""
        sql &= " SELECT U.OCID" & vbCrLf
        '邏輯改為：僅計算電話抽訪原因為非「2:實地抽訪時未到」的件數//FN_GET_UNEXPECTVISITOR N,3
        sql &= " ,COUNT(CASE WHEN ISNULL(U.TELVISITREASON,'1')!='2' THEN 1 END) CTALL3" & vbCrLf
        '邏輯為：僅計算電話抽訪原因=「實地抽訪時未到」的件數
        sql &= " ,COUNT(CASE WHEN U.TELVISITREASON='2' THEN 1 END) CTALL4" & vbCrLf
        '/*累計不預告電話抽訪次數*/ 'sql &= " ,COUNT(1) CTALL" & vbCrLf
        '/*累計不預告電話抽訪異常次數*/
        sql &= " ,COUNT(CASE WHEN U.ITEM10='2' THEN 1 END) VitTelN" & vbCrLf
        'sql &= " ,WM_CONCAT(DISTINCT CONVERT(varchar, U.APPLYDATE, 111)) cuAPPLYDATE" & vbCrLf
        sql &= " ,dbo.[FN_GET_UNEXPECTVISITOR](U.OCID,3) cuAPPLYDATE3" & vbCrLf '僅顯示電話抽訪原因為非「實地抽訪時未到」的電話訪視日期。AND ISNULL(U.TELVISITREASON,1)='1'
        sql &= " ,dbo.[FN_GET_UNEXPECTVISITOR](U.OCID,4) cuAPPLYDATE4" & vbCrLf '僅顯示電話抽訪原因=「實地抽訪時未到」的電話訪視日期。AND U.TELVISITREASON='2'
        sql &= " FROM dbo.CLASS_UNEXPECTTEL U WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN WC1 cc ON cc.OCID =U.OCID" & vbCrLf
        sql &= " GROUP BY U.ocid" & vbCrLf
        Return sql
    End Function

    ''' <summary>取得搜尋範圍 (學員) (SQL WHERE)</summary>
    ''' <returns></returns>
    Function Get_SearchStr2() As String
        Dim strSearch As String = ""
        'ss
        Dim v_Appmoney As String = TIMS.GetListValue(Appmoney)

        If Appmoney.SelectedIndex <> 0 AndAlso v_Appmoney <> "" Then strSearch &= " and ss.AppliedStatus = '" & v_Appmoney & "'" & vbCrLf

        If AllotDate1.Text <> "" Then strSearch &= " and ss.AllotDate >= " & TIMS.To_date(AllotDate1.Text) & vbCrLf

        If AllotDate2.Text <> "" Then strSearch &= " and ss.AllotDate <= " & TIMS.To_date(AllotDate2.Text) & vbCrLf

        Return strSearch
    End Function

    ''' <summary> 取得搜尋範圍 (學員)1 CLASS_STUDENTSOFCLASS </summary>
    ''' <returns></returns>
    Function Get_SqlSchS1() As String
        Dim strSearchStd As String = Get_SearchStr2() '取得搜尋範圍 (學員)

        Const cst_BudgetID_All_in As String = "'02','03','97'" '01','02','03','97'

        Dim sql As String = ""
        sql &= " select cs.OCID," & vbCrLf
        '"/*撥款日期*/"
        sql &= " MIN(ss.AllotDate) AllotDate," & vbCrLf
        '/*開訓人次-就保*/
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID='03' then 1 else 0 end ) openstudcount1," & vbCrLf
        '/*開訓人次-就安*/
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID='02' then 1 else 0 end ) openstudcount2," & vbCrLf
        '/*開訓人次-公務*/
        'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID='01' then 1 else 0 end ) openstudcount3," & vbCrLf
        '/*開訓人次-協助*/ 公務ECFA
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID='97' then 1 else 0 end ) openstudcount97," & vbCrLf
        '/*開訓人次-合計*/ Cst_實際開訓人次加總
        '班級開訓，學員開訓後14天實際錄訓人數,且有選擇預算別(公務/就安/就保或公務(ECFA))
        'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID IN ('01','02','03','97') then 1 else 0 end ) openstudcountall," & vbCrLf
        ' 1.【綜合查詢統計表】- 實際開訓人次 班級開訓，學員開訓後14天實際錄訓人次,不管有沒有預算別
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 then 1 else 0 end) openstudcountall," & vbCrLf
        '/*預估補助費-就保*/
        sql &= " sum( case when ISNULL(cs.SupplyID,'') ='' or cc.ATNum=0 then 0" & vbCrLf
        sql &= " when cs.SupplyID='1' and cc.ATNum<>0 and cs.BudgetID='03' and cc.Notopen='N' and cs.IsApprPaper='Y' then (ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1))*0.8" & vbCrLf
        sql &= " when cs.SupplyID='9' and cc.ATNum<>0 and cs.BudgetID='03' and cc.Notopen='N' and cs.IsApprPaper='Y' then 0" & vbCrLf
        sql &= " when cs.SupplyID='2' and cc.ATNum<>0 and cs.BudgetID='03' and cc.Notopen='N' and cs.IsApprPaper='Y' then ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1) end ) cost1," & vbCrLf
        '/*預估補助費-就安*/
        sql &= " sum( case when ISNULL(cs.SupplyID,'') =''  or cc.ATNum=0 then 0" & vbCrLf
        sql &= " when cs.SupplyID='1' and cc.ATNum<>0 and cs.BudgetID='02' and cc.Notopen='N' and cs.IsApprPaper='Y' then (ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1))*0.8" & vbCrLf
        sql &= " when cs.SupplyID='9' and cc.ATNum<>0 and cs.BudgetID='02' and cc.Notopen='N' and cs.IsApprPaper='Y' then 0" & vbCrLf
        sql &= " when cs.SupplyID='2' and cc.ATNum<>0 and cs.BudgetID='02' and cc.Notopen='N' and cs.IsApprPaper='Y' then ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1) end ) cost2," & vbCrLf
        '/*預估補助費-公務*/
        'sql &= " sum( case when ISNULL(cs.SupplyID,'') =''  or cc.ATNum=0 then 0" & vbCrLf
        'sql &= " when cs.SupplyID='1' and cc.ATNum<>0 and cs.BudgetID='01' and cc.Notopen='N' and cs.IsApprPaper='Y' then (ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1))*0.8" & vbCrLf
        'sql &= " when cs.SupplyID='9' and cc.ATNum<>0 and cs.BudgetID='01' and cc.Notopen='N' and cs.IsApprPaper='Y' then 0" & vbCrLf
        'sql &= " when cs.SupplyID='2' and cc.ATNum<>0 and cs.BudgetID='01' and cc.Notopen='N' and cs.IsApprPaper='Y' then ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1) end ) cost3," & vbCrLf
        '/*預估補助費-協助*/ 公務ECFA
        sql &= " sum( case when ISNULL(cs.SupplyID,'') =''  or cc.ATNum=0 then 0" & vbCrLf
        sql &= " when cs.SupplyID='1' and cc.ATNum<>0 and cs.BudgetID='97' and cc.Notopen='N' and cs.IsApprPaper='Y' then (ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1))*0.8" & vbCrLf
        sql &= " when cs.SupplyID='9' and cc.ATNum<>0 and cs.BudgetID='97' and cc.Notopen='N' and cs.IsApprPaper='Y' then 0" & vbCrLf
        sql &= " when cs.SupplyID='2' and cc.ATNum<>0 and cs.BudgetID='97' and cc.Notopen='N' and cs.IsApprPaper='Y' then ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1) end ) cost97," & vbCrLf
        '/*預估補助費-合計*/
        sql &= " sum( case when ISNULL(cs.SupplyID,'') =''  or cc.ATNum=0 then 0" & vbCrLf
        sql &= String.Concat(" when cs.SupplyID='1' and cc.ATNum<>0 and cs.BudgetID IN (", cst_BudgetID_All_in, ") and cc.Notopen='N' and cs.IsApprPaper='Y' then (ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1))*0.8") & vbCrLf
        sql &= String.Concat(" when cs.SupplyID='9' and cc.ATNum<>0 and cs.BudgetID IN (", cst_BudgetID_All_in, ") and cc.Notopen='N' and cs.IsApprPaper='Y' then 0") & vbCrLf
        sql &= String.Concat(" when cs.SupplyID='2' and cc.ATNum<>0 and cs.BudgetID IN (", cst_BudgetID_All_in, ") and cc.Notopen='N' and cs.IsApprPaper='Y' then ISNULL(cc.TotalCost,0)/ISNULL(cc.ATNum,1) end ) costAll,") & vbCrLf
        '/*結訓-就保人次*/
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='03'"
        sql &= " then 1 else 0 end ) closestudcout03," & vbCrLf
        '/*結訓-就安人次*/
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='02'"
        sql &= " then 1 else 0 end ) closestudcout02," & vbCrLf
        '/*結訓-公務人次*/
        'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        'sql &= " and cs.CreditPoints is not NULL"
        'sql &= " and cs.StudStatus Not IN (2,3)"
        'sql &= " and cc.FTDate<GETDATE()"
        'sql &= " and cs.BudgetID='01'" & vbCrLf
        'sql &= " then 1 else 0 end ) closestudcout01," & vbCrLf
        '/*結訓-協助人次*/ 公務ECFA
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='97'"
        sql &= " then 1 else 0 end ) closestudcout97," & vbCrLf
        '/*結訓-合計人次*/Cst_結訓人次
        '班級開訓，學員補助符合補助者 沒有離退訓，只剩開結訓人數，且結訓日期已過今天,且有選擇預算別(公務/就安/就保或公務(ECFA))
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= String.Concat(" and cs.BudgetID IN (", cst_BudgetID_All_in, ")") & vbCrLf
        sql &= " then 1 else 0 end ) closestudcoutall," & vbCrLf

        '/*離訓人次-退訓人次*/
        'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.StudStatus=2 then 1 else 0 end ) std_cnt2," & vbCrLf
        'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.StudStatus=3 then 1 else 0 end ) std_cnt3," & vbCrLf
        'by AMU 20220401-學員資料維護尚未儲存(未確認)即離退-離訓人次-退訓人次
        sql &= " sum(case when cc.Notopen='N' and cs.StudStatus=2 then 1 else 0 end) std_cnt2," & vbCrLf
        sql &= " sum(case when cc.Notopen='N' and cs.StudStatus=3 then 1 else 0 end) std_cnt3," & vbCrLf

        '/*就保合計撥款人次*/
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and ss.AppliedStatus='1'"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='03'" & vbCrLf
        sql &= " then 1 else 0 end ) budcountall3," & vbCrLf
        '/*就安合計撥款人次*/
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and ss.AppliedStatus='1'"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='02'" & vbCrLf
        sql &= " then 1 else 0 end ) budcountall2," & vbCrLf
        '/*公務合計撥款人次*/
        'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        'sql &= " and cs.CreditPoints is not NULL"
        'sql &= " and ss.AppliedStatus='1'"
        'sql &= " and cs.StudStatus Not IN (2,3)"
        'sql &= " and cc.FTDate<GETDATE()"
        'sql &= " and cs.BudgetID='01'" & vbCrLf
        'sql &= " then 1 else 0 end ) budcountall3," & vbCrLf
        '/*協助合計撥款人次*/Cst_撥款人次 公務ECFA
        '班級開訓，學員補助符合補助者 沒有離退訓，只剩開結訓人數，且結訓日期已過今天,且有選擇預算別(公務/就安/就保或公務(ECFA))-學員經費撥款狀態：已撥款之人數 
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and ss.AppliedStatus='1'"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='97'"
        sql &= " then 1 else 0 end ) budcountall97," & vbCrLf

        For Each dr1 As DataRow In dtIdentity.Rows
            '/*就保一般特殊身分學員撥款人次-XX */
            sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
            sql &= " and cs.CreditPoints is not NULL"
            sql &= " and ss.AppliedStatus='1'"
            sql &= " and cs.StudStatus Not IN (2,3)"
            sql &= " and cc.FTDate<GETDATE()"
            sql &= " and cs.BudgetID='03'"
            sql &= " and cs.MIdentityID='" & dr1("IdentityID") & "'"
            sql &= " then 1 else 0 end ) bud03count" & dr1("IdentityID") & "," & vbCrLf

            '/*就安一般特殊身分學員撥款人次-XX */
            sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
            sql &= " and cs.CreditPoints is not NULL"
            sql &= " and ss.AppliedStatus='1'"
            sql &= " and cs.StudStatus Not IN (2,3)"
            sql &= " and cc.FTDate<GETDATE()"
            sql &= " and cs.BudgetID='02'"
            sql &= " and cs.MIdentityID='" & dr1("IdentityID") & "'"
            sql &= " then 1 else 0 end ) bud02count" & dr1("IdentityID") & "," & vbCrLf

            '/*公務一般特殊身分學員撥款人次-XX */
            'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
            'sql &= " and cs.CreditPoints is not NULL"
            'sql &= " and ss.AppliedStatus='1'"
            'sql &= " and cs.StudStatus Not IN (2,3)"
            'sql &= " and cc.FTDate<GETDATE()"
            'sql &= " and cs.BudgetID='01'" 
            'sql &= " and cs.MIdentityID='" & dr1("IdentityID") & "'" 
            'sql &= " then 1 else 0 end ) bud01count" & dr1("IdentityID") & "," & vbCrLf

            '/*協助一般特殊身分學員撥款人次-XX */ 公務ECFA
            sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
            sql &= " and cs.CreditPoints is not NULL"
            sql &= " and ss.AppliedStatus='1'"
            sql &= " and cs.StudStatus Not IN (2,3)"
            sql &= " and cc.FTDate<GETDATE()"
            sql &= " and cs.BudgetID='97'"
            sql &= " and cs.MIdentityID='" & dr1("IdentityID") & "'"
            sql &= " then 1 else 0 end ) bud97count" & dr1("IdentityID") & "," & vbCrLf
        Next
        '/*就保合計撥款金額*/
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and ss.AppliedStatus='1'"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='03'"
        sql &= " then ss.SumOfMoney else 0 end ) budmoneyall3," & vbCrLf
        '/*就安合計撥款金額*/
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and ss.AppliedStatus='1'"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='02'"
        sql &= " then ss.SumOfMoney else 0 end ) budmoneyall2," & vbCrLf
        '/*公務合計撥款金額*/
        'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        'sql &= " and cs.CreditPoints is not NULL"
        'sql &= " and ss.AppliedStatus = '1'" 
        'sql &= " and cs.StudStatus Not IN (2,3)"
        'sql &= " and cc.FTDate<GETDATE()"
        'sql &= " and cs.BudgetID='01'" 
        'sql &= " then ss.SumOfMoney else 0 end ) budmoneyall3," & vbCrLf
        '/*協助合計撥款金額*/ 公務ECFA
        sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
        sql &= " and cs.CreditPoints is not NULL"
        sql &= " and ss.AppliedStatus='1'"
        sql &= " and cs.StudStatus Not IN (2,3)"
        sql &= " and cc.FTDate<GETDATE()"
        sql &= " and cs.BudgetID='97'"
        sql &= " then ss.SumOfMoney else 0 end ) budmoneyall97," & vbCrLf

        For Each dr1 As DataRow In dtIdentity.Rows
            '/*就保 一般特殊身分學員撥款金額-XX*/
            sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
            sql &= " and cs.CreditPoints is not NULL"
            sql &= " and ss.AppliedStatus='1'"
            sql &= " and cs.StudStatus Not IN (2,3)"
            sql &= " and cc.FTDate<GETDATE()"
            sql &= " and cs.BudgetID='03'"
            sql &= " and cs.MIdentityID='" & dr1("IdentityID") & "'"
            sql &= " then ss.SumOfMoney else 0 end ) bud03money" & dr1("IdentityID") & "," & vbCrLf

            '/*就安 一般特殊身分學員撥款金額-XX*/
            sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
            sql &= " and cs.CreditPoints is not NULL"
            sql &= " and ss.AppliedStatus='1'"
            sql &= " and cs.StudStatus Not IN (2,3)"
            sql &= " and cc.FTDate<GETDATE()"
            sql &= " and cs.BudgetID='02'"
            sql &= " and cs.MIdentityID='" & dr1("IdentityID") & "'"
            sql &= " then ss.SumOfMoney else 0 end ) bud02money" & dr1("IdentityID") & "," & vbCrLf

            '/*公務 一般特殊身分學員撥款金額-XX*/
            'sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
            'sql &= " and cs.CreditPoints is not NULL"
            'sql &= " and ss.AppliedStatus='1'"
            'sql &= " and cs.StudStatus Not IN (2,3)"
            'sql &= " and cc.FTDate<GETDATE()"
            'sql &= " and cs.BudgetID='01'" 
            'sql &= " and cs.MIdentityID='" & dr1("IdentityID") & "'" 
            'sql &= " then ss.SumOfMoney else 0 end ) bud01money" & dr1("IdentityID") & "," & vbCrLf

            '/*協助 一般特殊身分學員撥款金額-XX*/ 公務ECFA
            sql &= " sum(case when cc.Notopen='N' and cs.IsApprPaper='Y'"
            sql &= " and cs.CreditPoints is not NULL"
            sql &= " and ss.AppliedStatus='1'"
            sql &= " and cs.StudStatus Not IN (2,3)"
            sql &= " and cc.FTDate<GETDATE()"
            sql &= " and cs.BudgetID='97'"
            sql &= " and cs.MIdentityID='" & dr1("IdentityID") & "'"
            sql &= " then ss.SumOfMoney else 0 end ) bud97money" & dr1("IdentityID") & "," & vbCrLf
        Next

        '/* 協助特殊身分男性學員人數 */ 公務ECFA
        sql &= " sum(case when cc.NotOpen='N' and cs.IsApprPaper='Y' and cs.StudStatus Not IN (2,3)"
        sql &= " and cs.BudgetID='97' and ssi.Sex='M' then 1 else 0 end) SexNumxM" & vbCrLf
        '/* 協助特殊身分女性學員人數 */ 公務ECFA
        sql &= " ,sum(case when cc.NotOpen='N' and cs.IsApprPaper='Y' and cs.StudStatus Not IN (2,3)"
        sql &= " and cs.BudgetID='97' and ssi.Sex='F' then 1 else 0 end) SexNumxF" & vbCrLf
        '/*實際開訓男性人數*/ 實際開訓性別人數
        sql &= " ,count(case when cc.NotOpen='N' and cs.IsApprPaper='Y' and cs.StudStatus Not IN (2,3)"
        sql &= String.Concat(" and cs.BudgetID IN (", cst_BudgetID_All_in, ") and ssi.Sex='M' then 1 end) SexCNTM", vbCrLf)
        '/*實際開訓女性人數*/ 實際開訓性別人數
        sql &= " ,count(case when cc.NotOpen='N' and cs.IsApprPaper='Y' and cs.StudStatus Not IN (2,3)"
        sql &= String.Concat(" and cs.BudgetID IN (", cst_BudgetID_All_in, ") and ssi.Sex='F' then 1 end) SexCNTF", vbCrLf)

        sql &= " FROM dbo.CLASS_STUDENTSOFCLASS cs WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO ssi WITH(NOLOCK) on ssi.SID=cs.SID" & vbCrLf
        sql &= " JOIN WC1 cc on cc.OCID =cs.ocid" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_SUBSIDYCOST ss WITH(NOLOCK) on ss.SOCID=cs.SOCID" & vbCrLf
        '取得搜尋範圍 (學員)
        sql &= " WHERE 1=1" & strSearchStd
        'and pp.TPlanID ='28'and pp.IsApprPaper ='Y'and ip.Years='2016'and ip.Distid='005'
        sql &= " GROUP BY cs.ocid" & vbCrLf
        Return sql
    End Function

    ''' <summary>ALL</summary>
    ''' <returns></returns>
    Function Get_SqlSchAll() As String
        Dim sql As String = ""
        '(YEARS)
        sql &= " SELECT c.TMID" & vbCrLf
        sql &= " ,c.YEARS" & vbCrLf '計畫年度
        sql &= " ,c.OrgTypeName" & vbCrLf '單位屬性 
        sql &= " ,c.OrgPlanName2" & vbCrLf '自主/產投(產投計畫別)
        sql &= " ,c.DistName" & vbCrLf
        sql &= " ,c.OCID" & vbCrLf
        sql &= " ,c.ORGNAME" & vbCrLf
        sql &= " ,c.ClassName" & vbCrLf
        sql &= " ,c.FIRSTSORT" & vbCrLf
        sql &= " ,c.CyclType" & vbCrLf
        '2019年啟用'申請階段 'APPSTAGE
        sql &= " ,c.APPSTAGE" & vbCrLf

        sql &= " ,c.ClassID" & vbCrLf
        sql &= " ,c.pSTDate STDate" & vbCrLf
        sql &= " ,c.pFTDate FDDate" & vbCrLf
        'sql &= " ,c.STDate" & vbCrLf
        'sql &= " ,c.FDDate" & vbCrLf
        sql &= " ,CONVERT(varchar, a.AllotDate, 111) AllotDate" & vbCrLf
        '課程申請流水號
        sql &= " ,c.PSNO28" & vbCrLf
        '上架日期
        'sql &= " ,c.ONSHELLDATE" & vbCrLf '(yyyy/MM/dd HH:mm)
        sql &= " ,format(c.ONSHELLDATE,'yyyy/MM/dd HH:mm') ONSHELLDATE" & vbCrLf
        '開放報名結束日期開放報名日期
        sql &= " ,CONVERT(varchar, c.SENTERDATE, 111) SENTERDATE" & vbCrLf
        sql &= " ,CONVERT(varchar, c.FENTERDATE, 111) FENTERDATE" & vbCrLf
        '課程備註 'https://jira.turbotech.com.tw/browse/TIMSC-218
        sql &= " ,c.memo8,c.memo82" & vbCrLf
        'https://jira.turbotech.com.tw/browse/TIMSC-301
        sql &= " ,c.FIXSUMCOST" & vbCrLf
        sql &= " ,c.ACTHUMCOST" & vbCrLf
        sql &= " ,c.FIXExceeDesc" & vbCrLf
        sql &= " ,c.METSUMCOST" & vbCrLf
        sql &= " ,c.METCOSTPER" & vbCrLf
        sql &= " ,c.METExceeDesc" & vbCrLf

        sql &= " ,c.PLANID,c.COMIDNO,c.SEQNO" & vbCrLf
        'oo.orgZipCode=iz2.ZipCode
        sql &= " ,c.CTName2" & vbCrLf
        'sql &= " ,c.PackageType" & vbCrLf
        sql &= " ,c.PackageTypeN" & vbCrLf
        sql &= " ,c.ADefGovCost" & vbCrLf
        sql &= " ,c.ATNum" & vbCrLf
        sql &= " ,c.DefGovCost" & vbCrLf
        sql &= " ,c.TNum" & vbCrLf
        sql &= " ,c.THours" & vbCrLf
        '/*人時成本*/ Cst_人時成本
        'sql &= " ,(a.DefGovCost/(case when a.TNum = 0 then 1 else a.TNum end))" & vbCrLf
        'sql &= " /(case when a.THours = 0 then 1 else a.THours end) as PhCost" & vbCrLf
        '/*人時成本 總費用/訓練人數/訓練時數 Cst_人時成本 */ PhCost
        sql &= " ,ROUND((c.TotalCost/(case when c.TNum=0 then 1 else c.TNum end))" & vbCrLf
        sql &= " /(case when c.THours = 0 then 1 else c.THours end),3) PhCost" & vbCrLf

        sql &= " ,c.WEEKSTIME" & vbCrLf
        sql &= " ,c.BusPackage" & vbCrLf
        sql &= " ,c.BusPackage2" & vbCrLf
        sql &= " ,c.BusPackage3" & vbCrLf
        sql &= " ,c.PlanTeacher" & vbCrLf
        '課程分類 (view_depot12 vd12)
        'sql &= " ,c.kname12" & vbCrLf
        sql &= " ,c.Pkname12" & vbCrLf
        '轄區重點產業
        sql &= " ,c.kname13" & vbCrLf
        sql &= " ,c.kname15" & vbCrLf
        '生產力4.0 'SELECT KID,KNAME FROM Key_Business WHERE DEPID='14' AND status is null
        sql &= " ,c.kname14" & vbCrLf
        '六大新興產業 'SELECT * FROM view_Depot10
        '十大重點服務業'SELECT * FROM view_Depot11
        '四大新興智慧型產業'SELECT * FROM view_Depot09
        sql &= " ,c.kname1" & vbCrLf
        sql &= " ,c.kname2" & vbCrLf
        sql &= " ,c.kname3" & vbCrLf
        'sql &= " ,c.KNAME17" & vbCrLf '政府政策性產業(108年之後不使用此欄)
        'sql &= " ,c.KNAME19" & vbCrLf '政府政策性產業(108年之後不使用此欄)
        sql &= " ,c.KNAME18" & vbCrLf
        '2019年啟用 work2019x01:2019 政府政策性產業
        sql &= " ,c.KID20" & vbCrLf
        sql &= " ,c.D20KNAME1,c.D20KNAME2,c.D20KNAME3,c.D20KNAME4,c.D20KNAME5,c.D20KNAME6" & vbCrLf
        'dd.KID22,dd.KNAME22 進階政策性產業類別
        sql &= " ,c.KNAME22" & vbCrLf
        '2025 政府政策性產業 (產投)
        sql &= " ,c.D25KNAME1,c.D25KNAME2,c.D25KNAME3,c.D25KNAME4,c.D25KNAME5,c.D25KNAME6,c.D25KNAME7,c.D25KNAME8" & vbCrLf
        '2026 政府政策性產業 (產投)
        sql &= " ,c.D26KNAME1,c.D26KNAME2,c.D26KNAME7,c.D26KNAME3,c.D26KNAME5,c.D26KNAME4,c.D26KNAME6,c.D26KNAME8,c.D26KNAME9"
        ',a.GovClass,a.GCode1,a.GCode2
        sql &= " ,c.GCodeName" & vbCrLf 'GCodeName: 訓練業別編碼
        sql &= " ,c.GCNAME" & vbCrLf '訓練業別
        'sql &= " ,c.TJOBNAME" & vbCrLf'(職訓業別)
        'CJOB_UNKEY
        sql &= " ,c.CJOB_UNKEY" & vbCrLf
        'sql &= " ,a.CJOBNAME1" & vbCrLf
        'sql &= " ,a.CJOBNAME2" & vbCrLf
        'ExportStr &= "<td>訓練業別</td>"  
        'ExportStr &= "<td>通俗職類-大類</td>"  
        'ExportStr &= "<td>通俗職類-小類</td>"  
        sql &= " ,c.CCName" & vbCrLf

        sql &= " ,c.AddressSciPTID ,c.AddressTechPTID" & vbCrLf
        'sql &= " ,a.p1ADDRESS" 
        'sql &= " ,a.p2ADDRESS" 

        '/* '組合 Cst_上課地址及教室*/
        sql &= " ,c.S1PLACEID ,c.S2PLACEID ,c.T1PLACEID ,c.T2PLACEID ,c.SPPLACEID ,c.TPPLACEID
 ,c.S1PTID ,c.S2PTID ,c.T1PTID ,c.T2PTID ,c.SPPTID ,c.TPPTID
 ,c.S1PLACENAME ,c.S2PLACENAME ,c.T1PLACENAME ,c.T2PLACENAME ,c.SPPLACENAME ,c.TPPLACENAME
 ,c.S1ZIPCODE ,c.S1ADDRESS ,c.S1ZIP6W
 ,c.S2ZIPCODE ,c.S2ADDRESS ,c.S2ZIP6W
 ,c.T1ZIPCODE ,c.T1ADDRESS ,c.T1ZIP6W
 ,c.T2ZIPCODE ,c.T2ADDRESS ,c.T2ZIP6W
 ,c.SPZIPCODE ,c.SPADDRESS ,c.SPZIP6W
 ,c.TPZIPCODE ,c.TPADDRESS ,c.TPZIP6W"

        sql &= " ,c.ProTechHours,c.ContactName,c.ContactPhone,c.ContactMobile" & vbCrLf
        sql &= " ,c.NotOpenN" & vbCrLf
        'ICAPNUM-iCAP標章證號-cst_iCAP標章證號及效期
        sql &= " ,c.ICAPNUM" & vbCrLf
        sql &= " ,format(c.iCAPMARKDATE,'yyyy/MM/dd') iCAPMARKDATE" & vbCrLf
        sql &= " ,case c.ENTERSUPPLYSTYLE when 1 then '全額' when 2 then '50%'end ENTERSUPPLYSTYLE_N" & vbCrLf '報名繳費方式
        sql &= " ,format(c.VERIFYDATE,'yyyy/MM/dd') VERIFYDATE" & vbCrLf '為分署於該功能登打之經查核確認日期。
        '辦理方式-遠距教學 DISTANCE,dbo.FN_GET_DISTANCE(DISTANCE) DISTANCE_N
        sql &= " ,c.DISTANCE" & vbCrLf
        sql &= " ,dbo.FN_GET_DISTANCE(c.DISTANCE) DISTANCE_N" & vbCrLf
        'sql &= " ,c.YR1,c.YR2,c.YR3" & vbCrLf
        'sql &= " ,c.PCNT11,c.PCNT12,c.PCNT13" & vbCrLf

        sql &= " ,a.openstudcount1" & vbCrLf
        sql &= " ,a.openstudcount2" & vbCrLf
        'sql &= " ,a.openstudcount3" & vbCrLf
        sql &= " ,a.openstudcount97" & vbCrLf
        sql &= " ,a.openstudcountall" & vbCrLf 'Cst_實際開訓人次加總

        sql &= " ,ISNULL(a.cost1,0) cost1" & vbCrLf
        sql &= " ,ISNULL(a.cost2,0) cost2" & vbCrLf
        'sql &= " ,ISNULL(a.cost3,0) cost3" & vbCrLf
        sql &= " ,ISNULL(a.cost97,0) cost97" & vbCrLf
        sql &= " ,ISNULL(a.costAll,0) costAll" & vbCrLf
        'sql &= " ,a.closestudcout" & vbCrLf
        'sql &= " ,a.closestudcout01" & vbCrLf '結訓-公務人次
        sql &= " ,a.closestudcout02" & vbCrLf
        sql &= " ,a.closestudcout03" & vbCrLf
        sql &= " ,a.closestudcout97" & vbCrLf
        sql &= " ,a.closestudcoutall" & vbCrLf
        sql &= " ,a.std_cnt2" & vbCrLf
        sql &= " ,a.std_cnt3" & vbCrLf

        sql &= " ,a.budcountall3"
        sql &= " ,a.budcountall2"
        'sql &= " ,a.budcountall3" & vbCrLf '公務合計撥款人次
        sql &= " ,a.budcountall97" & vbCrLf

        For Each dr1 As DataRow In dtIdentity.Rows
            sql &= " ,a.bud03count" & Convert.ToString(dr1("IdentityID"))
            sql &= " ,a.bud02count" & Convert.ToString(dr1("IdentityID"))
            'sql &= " ,a.bud01count" & Convert.ToString(dr1("IdentityID")) & vbCrLf  '/*公務一般特殊身分學員撥款人次-XX */
            sql &= " ,a.bud97count" & Convert.ToString(dr1("IdentityID")) & vbCrLf
        Next

        sql &= " ,a.budmoneyall3" & vbCrLf
        sql &= " ,a.budmoneyall2" & vbCrLf
        'sql &= " ,a.budmoneyall3" & vbCrLf'/*公務合計撥款金額*/
        sql &= " ,a.budmoneyall97" & vbCrLf

        For Each dr1 As DataRow In dtIdentity.Rows
            sql &= " ,a.bud03money" & Convert.ToString(dr1("IdentityID"))
            sql &= " ,a.bud02money" & Convert.ToString(dr1("IdentityID"))
            'sql &= " ,a.bud01money" & Convert.ToString(dr1("IdentityID")) & vbCrLf    '/*公務 一般特殊身分學員撥款金額-XX*/
            sql &= " ,a.bud97money" & Convert.ToString(dr1("IdentityID")) & vbCrLf
        Next

        '/*總特殊學員就保人次*/
        sql &= " ,a.budcountall3 - a.bud03count01 Sbudcount01"
        '/*總特殊學員就安人次*/
        sql &= " ,a.budcountall2 - a.bud02count01 Sbudcount02"
        '/*總特殊學員公務人次*/
        'sql &= " ,a.budcountall3 - a.bud01count01 Sbudcount03" & vbCrLf
        '/*總特殊學員協助人次*/ 公務ECFA
        sql &= " ,a.budcountall97 - a.bud97count01 Sbudcount97" & vbCrLf

        '/*總特殊學員就保金額*/
        sql &= " ,a.budmoneyall3 - a.bud03money01 Sbudmoney01"
        '/*總特殊學員就安金額*/
        sql &= " ,a.budmoneyall2 - a.bud02money01 Sbudmoney02"
        '/*總特殊學員公務金額*/
        'sql &= " ,a.budmoneyall3 - a.bud01money01 Sbudmoney03" & vbCrLf
        '/*總特殊學員協助金額*/ 公務ECFA
        sql &= " ,a.budmoneyall97 - a.bud97money01 Sbudmoney97" & vbCrLf

        sql &= " ,ISNULL(e.CUALL,0) CUALL" & vbCrLf '/*累計不預告實地抽訪次數*/
        sql &= " ,ISNULL(e.VW2CNT,0) VW2CNT" & vbCrLf '/*累計不預告視訊抽訪次數*/  '累計不預告視訊抽訪次數 視訊訪查 VW2CNT

        sql &= " ,ISNULL(e.vitN,0) vitN" & vbCrLf '/*累計不預告實地抽訪異常次數*/
        sql &= " ,ISNULL(e.vitTVN,0) vitTVN" & vbCrLf '/*累計不預告視訊抽訪異常次數*/
        sql &= " ,ISNULL(f.VitTelN,0) VitTelN" & vbCrLf '/*累計不預告電話抽訪異常次數*/

        sql &= " ,ISNULL(f.CTALL3,0) CTALL3" & vbCrLf '/*累計不預告電話抽訪次數*/
        sql &= " ,ISNULL(f.CTALL4,0) CTALL4" & vbCrLf
        'https://jira.turbotech.com.tw/browse/TIMSC-229
        '勾選綜合查詢統計表「不預告訪視次數-電話訪視」選項
        '，另新增顯示「訪視日期」欄位，欄位內容顯示之日期分號隔開。例：7/1;7/8;7/22
        '僅顯示電話抽訪原因為非「實地抽訪時未到」的電話訪視日期。AND ISNULL(U.TELVISITREASON,1)='1'
        sql &= " ,f.cuAPPLYDATE3 cuAPPLYDATE3" & vbCrLf '/*訪視日期*/
        '僅顯示電話抽訪原因=「實地抽訪時未到」的電話訪視日期。AND U.TELVISITREASON='2'
        sql &= " ,f.cuAPPLYDATE4 cuAPPLYDATE4" & vbCrLf '/*訪視日期*/
        'sql &= " ,ISNULL(e.vitN,0) + ISNULL(f.VitTelN,0) vtn" & vbCrLf

        '/*累計訪視異常原因*/
        sql &= " ,ISNULL(e.It22b01N,0) It22b01N" & vbCrLf
        sql &= " ,ISNULL(e.It22b02N,0) It22b02N" & vbCrLf
        sql &= " ,ISNULL(e.It22b03N,0) It22b03N" & vbCrLf
        sql &= " ,ISNULL(e.It22b06N,0) It22b06N" & vbCrLf
        sql &= " ,ISNULL(e.It22b04N,0) It22b04N" & vbCrLf
        sql &= " ,ISNULL(e.It22b05N,0) It22b05N" & vbCrLf
        'sql &= " ,ISNULL(e.It22b99N,0) It22b99N" & vbCrLf

        sql &= " ,e.ItAPPLYDATE ItAPPLYDATE" & vbCrLf 'ItAPPLYDATE 實地訪視日期
        sql &= " ,e.It22b99NOTE It22b99NOTE" & vbCrLf 'IT22B99NOTE 累計訪視異常原因 其他
        sql &= " ,e.LITEM23NOTE LITEM23NOTE" & vbCrLf 'LITEM23NOTE 累計訪視異常原因 其他補充說明
        sql &= " ,e.VW2APPLYDATE VW2APPLYDATE" & vbCrLf '視訊訪視日期 視訊訪查 VW2APPLYDATE

        sql &= " ,a.SexNumxM" & vbCrLf '/*協助男性人數*/ 公務ECFA
        sql &= " ,a.SexNumxF" & vbCrLf '/*協助女性人數*/ 公務ECFA
        sql &= " ,a.SexCNTM" & vbCrLf '/*實際開訓男性人數*/ 實際開訓性別人數
        sql &= " ,a.SexCNTF" & vbCrLf '/*實際開訓女性人數*/ 實際開訓性別人數

        sql &= " FROM WC1 c" & vbCrLf
        sql &= " LEFT JOIN WS1 a on a.OCID=c.OCID" & vbCrLf
        sql &= " LEFT JOIN WC2 e on e.OCID=c.OCID" & vbCrLf
        sql &= " LEFT JOIN WC3 f on f.OCID=c.OCID" & vbCrLf

        Return sql
    End Function

    ''' <summary>返回組合後的sql string</summary>
    ''' <param name="iType"></param>
    ''' <returns></returns>
    Function Get_SQL_Exp1(ByVal iType As Integer) As String
        'iType:1: old (HTML/XLS) EXP / 2: new xlsx EXP

        '使用歷史資料（為當天凌晨定版資料，匯出時間會較短）
        Dim flag_USE_MVD As Boolean = Cbl_His_MV_DATA.Checked

        '使用歷史資料（為當天凌晨定版資料，匯出時間會較短）
        Dim strWC1 As String = If(flag_USE_MVD, Get_SqlSchC1_MV(), Get_SqlSchC1_CC()) 'c PLAN_PLANINFO/CLASS

        Dim strWC2 As String = Get_SqlSchC2() 'e CLASS_UNEXPECTVISITOR
        Dim strWC3 As String = Get_SqlSchC3() 'f CLASS_UNEXPECTTEL
        Dim strWS1 As String = Get_SqlSchS1() 'a CLASS_STUDENTSOFCLASS / STUD_STUDENTINFO
        Dim strWAll1 As String = Get_SqlSchAll() 'ALL
        'c WC1:／'a WS1:／'e WC2:／'f WC3:／

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= String.Concat("WITH WC1 AS (", strWC1, ")", vbCrLf) 'c
        sql &= String.Concat(",WC2 AS (", strWC2, ")", vbCrLf) 'e
        sql &= String.Concat(",WC3 AS (", strWC3, ")", vbCrLf) 'f
        sql &= String.Concat(",WS1 AS (", strWS1, ")", vbCrLf) 'a
        Select Case iType
            Case 1 'iType:1: old (HTML/XLS) EXP / 2: new xlsx EXP
                sql &= strWAll1
                'Return sql
            Case 2 'iType:1: old (HTML/XLS) EXP / 2: new xlsx EXP
                sql &= String.Concat(",WA1 AS (", strWAll1, ")", vbCrLf) 'ALL
                sql &= String.Concat(Get_SqlExpRptXLSX(), " FROM WA1 a", vbCrLf)
                'Return sql
        End Select

        'test 測試環境測試
        If (flag_chktest) Then TIMS.WriteLog(Me, $"--##SD_15_012.aspx,{vbCrLf}--,Get_SQL_Exp1 sql:{vbCrLf}{sql}")

        Return sql
    End Function

    '匯出SUB (SQL)'取得搜尋範圍 (班級) (SQL WHERE)-MV_CLASS_1-'Batch\Dbt_20190313 'Batch\dbt_20210217  '綜合查詢統計表(定版)
    Sub ExpRpt()
        Const cst_max_Timeout As Integer = 500
        'ByVal da As SqlDataAdapter
        '判斷是否在測試環境中。
        Dim flag_chktest As Boolean = If(TIMS.sUtl_ChkTest(), True, False) '(測試環境中)
        'test_flag = False
        'Dim sFileName1 As String = "綜合查詢統計表"
        Dim s_FileName1 As String = String.Format("綜合查詢統計表_{0}", TIMS.GetDateNo2(5))
        'Dim sql As String = ""
        'sql = " SELECT KID, KNAME FROM dbo.KEY_BUSINESS WHERE DEPID='20' ORDER BY KID"
        'Dim dtKID_N20 As DataTable = DbAccess.GetDataTable(sql, objConn)
        'Call TIMS.GET_NAME_KID20(CBLKID20_1, dtKID_N20, 1)
        'Call TIMS.GET_NAME_KID20(CBLKID20_2, dtKID_N20, 2)
        'Call TIMS.GET_NAME_KID20(CBLKID20_3, dtKID_N20, 3)
        'Call TIMS.GET_NAME_KID20(CBLKID20_4, dtKID_N20, 4)
        'Call TIMS.GET_NAME_KID20(CBLKID20_5, dtKID_N20, 5)
        'Call TIMS.GET_NAME_KID20(CBLKID20_6, dtKID_N20, 6)

        'Call setSelYears1(yearlist.SelectedValue)
        'Dim intYears As Integer = yearlist.SelectedValue 'sm.UserInfo.Years
        'Dim strSearchCls As String = Get_SearchStr() '取得搜尋範圍  (班級)
        'Dim strSearchStd As String = Get_SearchStr2() '取得搜尋範圍 (學員)

        '判斷是否在測試環境中。
        'If flag_chktest Then
        '    Dim slogMsg1 As String = "##SD_15_012, sql: " & sql & vbCrLf
        '    slogMsg1 &= "##SD_15_012, myParam: " & TIMS.GetMyValue3(myParam) & vbCrLf
        '    Dim rp_sql As String = Replace(sql, vbCrLf, "<br/>")
        '    Common.RespWrite(Me, TIMS.sUtl_AntiXss(rp_sql))
        '    Response.End() 'Exit Sub
        '    TIMS.writeLog(Me, String.Concat("##SD_15_012.aspx,", vbCrLf, ",ExpRpt sql:", vbCrLf, sql))
        'End If

        'Dim sql As String = "" ' Get_SQL_Exp1(1)
        Call TIMS.OpenDbConn(objConn)
        Dim dt As New DataTable

        'Const Cst_FileSavePath As String = "~/SD/01/Temp/"
        Dim parmsExp As New Hashtable
        Dim v_RBListExpType As String = TIMS.GetListValue(RBListExpType)
        Select Case v_RBListExpType
            Case "XLSX"
                g_ErrSql = Get_SQL_Exp1(2)
                Dim sCmd As New SqlCommand(g_ErrSql, objConn)
                sCmd.CommandTimeout = cst_max_Timeout '500
                With sCmd
                    .Parameters.Clear()
                    dt.Load(.ExecuteReader())
                End With

                dt.Columns("通俗職類-大類").ReadOnly = False
                dt.Columns("通俗職類-小類").ReadOnly = False
                If flag_use_上課地址及教室 Then dt.Columns("上課地址及教室").ReadOnly = False
                Dim sCJOB_UNKEY As String = ""
                For Each dr As DataRow In dt.Rows
                    sCJOB_UNKEY = Convert.ToString(dr("CJOB_UNKEY"))
                    'sql &= " <td>" & TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 1) & "</td>" & vbCrLf  '通俗職類-大類
                    'sql &= " <td>" & TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 2) & "</td>" & vbCrLf  '通俗職類-小類
                    dr("通俗職類-大類") = TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 1)
                    dr("通俗職類-小類") = TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 2)
                    'REM'REM'REM上課地址及教室
                    'Case Cst_上課地址及教室
                    'Dim spAddress As String = "" '組合地址用
                    ''組合 Cst_上課地址及教室
                    If flag_use_上課地址及教室 Then dr("上課地址及教室") = Get_AddressPlaceName(dtZip, dr)
                Next

                'Columns.Remove
                Dim s_CanDelColumn As String = "PLANID,COMIDNO,SEQNO,CJOB_UNKEY,S1PLACEID,S2PLACEID,T1PLACEID,T2PLACEID,S1PTID,S2PTID,T1PTID,T2PTID,S1PLACENAME,S2PLACENAME,T1PLACENAME,T2PLACENAME,S1ZIPCODE,S1ADDRESS,S1ZIP6W,S2ZIPCODE,S2ADDRESS,S2ZIP6W,T1ZIPCODE,T1ADDRESS,T1ZIP6W,T2ZIPCODE,T2ADDRESS,T2ZIP6W"
                For Each V_COL As String In s_CanDelColumn.Split(",")
                    dt.Columns.Remove(V_COL)
                Next
                'dt.DefaultView.Sort = "DistName,ClassID,orgname,OrgTypeName,ClassName,CyclType"
                'dt.DefaultView.Sort = "OrgPlanName2,DistName,orgname,FIRSTSORT,ClassName,CyclType"
                dt.DefaultView.Sort = "計畫別,各分署,訓練機構,提案意願順序,班別名稱,期別"
                dt = TIMS.dv2dt(dt.DefaultView)

                Call TIMS.CloseDbConn(objConn)
                Call TIMS.Get_XLSX_Response(Me, dt) ', Cst_FileSavePath
                If Response IsNot Nothing AndAlso (Response.IsClientConnected) Then Response.End()
                Return

            Case Else
                g_ErrSql = Get_SQL_Exp1(1)
                Dim sCmd As New SqlCommand(g_ErrSql, objConn)
                sCmd.CommandTimeout = cst_max_Timeout '500
                With sCmd
                    .Parameters.Clear()
                    dt.Load(.ExecuteReader())
                End With
                'dt.DefaultView.Sort = "DistName,ClassID,orgname,OrgTypeName,ClassName,CyclType"
                dt.DefaultView.Sort = "OrgPlanName2,DistName,orgname,FIRSTSORT,ClassName,CyclType"
                dt = TIMS.dv2dt(dt.DefaultView)

        End Select

        '設定-YR123
        'Dim out_parms As New Hashtable
        'Call TIMS.SET_YR123(sm, yearlist.SelectedValue, out_parms)
        'Dim T_YR1 As String = TIMS.GetMyValue2(out_parms, "T_YR1")
        'Dim T_YR2 As String = TIMS.GetMyValue2(out_parms, "T_YR2")
        'Dim T_YR3 As String = TIMS.GetMyValue2(out_parms, "T_YR3")

        'If dt.Rows.Count = 0 Then
        '    Common.MessageBox(Me, "目前條件查無資料!!")
        '    Exit Sub
        'End If

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("綜合查詢統計表", System.Text.Encoding.UTF8) & ".xls")
        ''Response.ContentType = "Application/octet-stream"
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        ''Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        ''Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        ''文件內容指定為Excel
        ''Response.ContentType = "application/ms-excel;charset=utf-8"
        'Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        'Common.RespWrite(Me, "<html>")
        'Common.RespWrite(Me, "<head>")
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        ''<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        ''套CSS值
        'Common.RespWrite(Me, "<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        'Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        ''mso-number-format:"0" 
        'Common.RespWrite(Me, "</style>")
        'Common.RespWrite(Me, "</head>")
        'Common.RespWrite(Me, "<body>")

        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        '將所有td欄位格式改 為"文字"
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        'mso-number-format:"0"  10進位無小數點 
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim sbHTML As New StringBuilder
        'Dim strHTML As String = ""
        'strHTML = "<meta http-equiv='Content-Type' content='application/vnd.ms-excel; charset=UTF-8'>" & vbCrLf
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        sbHTML.Append("<tr>")
        'Dim tmpAddr As String = "" '組合地址暫存用
        'Dim spAddress As String = "" '組合地址用
        Dim ExportStr As String = ""
        'ExportStr = ""
        If ChbExit.SelectedIndex = -1 Then
            'ExportStr &= "<td>分署</td>"  
            'If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then ExportStr &= "<td>計畫</td>"    '自主/產投(產投計畫別)
            ExportStr &= "<td>計畫年度</td>"
            ExportStr &= "<td>計畫</td>"
            ExportStr &= "<td>各分署</td>"
            ExportStr &= "<td>單位屬性</td>"
            ExportStr &= "<td>訓練機構</td>"

            ExportStr &= "<td>班別名稱</td>"
            If hid_ssYears.Value >= cst_y2019 Then 'sCaseYears
                ExportStr &= "<td>提案意願順序</td>"
            End If
            ExportStr &= "<td>期別</td>"
            ExportStr &= "<td>課程代碼</td>"
            If gflag_SHOW_APPSTAGE Then ExportStr &= "<td>申請階段</td>"   'APPSTAGE

            ExportStr &= "<td>開訓日期</td>"
            ExportStr &= "<td>結訓日期</td>"
            ExportStr &= "<td>課程分類</td>" '課程分類

            '政府政策性產業  If gflag_SHOW_2019_1 Then
            'If gflag_SHOW_2019_1 Then
            '「5+2」產業創新計畫 5+2產業'【台灣AI行動計畫】 KID='08''【數位國家創新經濟發展方案】KID='09''【國家資通安全發展方案】KID='10''【前瞻基礎建設計畫】'【新南向政策】KID='19'
            '    ExportStr &= "<td>5+2產業創新計畫</td>"
            '    ExportStr &= "<td>台灣AI行動計畫</td>"
            '    ExportStr &= "<td>數位國家創新經濟發展方案</td>"
            '    ExportStr &= "<td>國家資通安全發展方案</td>"
            '    ExportStr &= "<td>前瞻基礎建設計畫</td>"
            '    ExportStr &= "<td>新南向政策</td>"
            'End If
            If fg_Work2026x02 Then
                Dim KID26STR1 As String() = {"五大信賴產業推動方案", "六大區域產業及生活圈", "智慧國家2.0綱領", "新南向政策推動計畫", "國家人才競爭力躍升方案", "AI新十大建設推動方案", "台灣AI行動計畫2.0", "智慧機器人產業推動方案", "臺灣2050淨零轉型"}
                For i_kid26 As Integer = 0 To KID26STR1.Length - 1
                    ExportStr &= $"<td>{KID26STR1(i_kid26)}</td>"
                Next
            End If
            If gfg_SHOW_2025_1 Then
                '亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
                ExportStr &= "<td>亞洲矽谷</td>"
                ExportStr &= "<td>重點產業</td>"
                ExportStr &= "<td>台灣AI行動計畫</td>"
                ExportStr &= "<td>智慧國家方案</td>"
                ExportStr &= "<td>國家人才競爭力躍升方案</td>"
                ExportStr &= "<td>新南向政策</td>"
                ExportStr &= "<td>AI加值應用</td>"
                ExportStr &= "<td>職場續航</td>"
            End If
            'dd.KID22,dd.KNAME22 進階政策性產業類別
            ExportStr &= "<td>進階政策性產業類別</td>"

            ExportStr &= "<td>新興產業</td>"   '※六大新興產業			
            ExportStr &= "<td>重點服務業</td>"   '※十大重點服務業			
            'Select Case hid_ssYears.Value 'sCaseYears
            '    Case cst_y2017
            '    Case Else
            '        ExportStr &= "<td>新興智慧型產業</td>"   '※四大新興智慧型產業			
            'End Select
            ExportStr &= "<td>新興智慧型產業</td>"   '※四大新興智慧型產業			

            ExportStr &= "<td>訓練業別編碼</td>"
            ExportStr &= "<td>訓練業別</td>"
            ExportStr &= "<td>通俗職類-大類</td>"
            ExportStr &= "<td>通俗職類-小類</td>"
            ExportStr &= "<td>訓練職能</td>"

            ExportStr &= "<td>學科辦訓地縣市</td>"
            ExportStr &= "<td>術科辦訓地縣市</td>"
            ExportStr &= "<td>包班種類</td>"   '包班總類 包班種類
        Else
            'If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then ExportStr &= "<td>計畫</td>"    '自主/產投(產投計畫別)
            ExportStr &= "<td>計畫年度</td>"
            ExportStr &= "<td>計畫</td>"
            ExportStr &= "<td>各分署</td>"
            ExportStr &= "<td>單位屬性</td>"
            ExportStr &= "<td>訓練機構</td>"
            For a As Integer = 0 To Me.ChbExit.Items.Count - 1
                If ChbExit.Items.Item(a).Selected Then
                    Select Case ChbExit.Items.Item(a).Value
                        Case Cst_統一編號
                            ExportStr &= "<td>統一編號</td>"
                        Case Cst_立案縣市
                            ExportStr &= "<td>立案縣市</td>"
                    End Select
                End If
            Next
            ExportStr &= "<td>班別名稱</td>"
            ExportStr &= "<td>期別</td>"
            ExportStr &= "<td>課程代碼</td>"
            If gflag_SHOW_APPSTAGE Then ExportStr &= "<td>申請階段</td>"   'APPSTAGE

            ExportStr &= "<td>開訓日期</td>"
            ExportStr &= "<td>結訓日期</td>"
            ExportStr &= "<td>課程分類</td>"   '課程分類

            '政府政策性產業  If gflag_SHOW_2019_1 Then
            'If gflag_SHOW_2019_1 Then
            '    '「5+2」產業創新計畫 5+2產業
            '    '【台灣AI行動計畫】 KID='08'
            '    '【數位國家創新經濟發展方案】KID='09'
            '    '【國家資通安全發展方案】KID='10'
            '    '【前瞻基礎建設計畫】
            '    '【新南向政策】KID='19'
            '    ExportStr &= "<td>5+2產業創新計畫</td>"
            '    ExportStr &= "<td>台灣AI行動計畫</td>"
            '    ExportStr &= "<td>數位國家創新經濟發展方案</td>"
            '    ExportStr &= "<td>國家資通安全發展方案</td>"
            '    ExportStr &= "<td>前瞻基礎建設計畫</td>"
            '    ExportStr &= "<td>新南向政策</td>"
            'End If
            '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
            '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
            If fg_Work2026x02 Then
                Dim KID26STR1 As String() = {"五大信賴產業推動方案", "六大區域產業及生活圈", "智慧國家2.0綱領", "新南向政策推動計畫", "國家人才競爭力躍升方案", "AI新十大建設推動方案", "台灣AI行動計畫2.0", "智慧機器人產業推動方案", "臺灣2050淨零轉型"}
                For i_kid26 As Integer = 0 To KID26STR1.Length - 1
                    ExportStr &= $"<td>{KID26STR1(i_kid26)}</td>"
                Next
            End If
            If gfg_SHOW_2025_1 Then
                '亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
                ExportStr &= "<td>亞洲矽谷</td>"
                ExportStr &= "<td>重點產業</td>"
                ExportStr &= "<td>台灣AI行動計畫</td>"
                ExportStr &= "<td>智慧國家方案</td>"
                ExportStr &= "<td>國家人才競爭力躍升方案</td>"
                ExportStr &= "<td>新南向政策</td>"
                ExportStr &= "<td>AI加值應用</td>"
                ExportStr &= "<td>職場續航</td>"
            End If
            'dd.KID22,dd.KNAME22 進階政策性產業類別
            ExportStr &= "<td>進階政策性產業類別</td>"

            ExportStr &= "<td>新興產業</td>"   '※六大新興產業			
            ExportStr &= "<td>重點服務業</td>"   '※十大重點服務業			
            'Select Case hid_ssYears.Value 'sCaseYears
            '    Case cst_y2017
            '    Case Else
            '        ExportStr &= "<td>新興智慧型產業</td>"   '※四大新興智慧型產業			
            'End Select
            ExportStr &= "<td>新興智慧型產業</td>"   '※四大新興智慧型產業			

            ExportStr &= "<td>訓練業別編碼</td>"
            ExportStr &= "<td>訓練業別</td>"
            ExportStr &= "<td>通俗職類-大類</td>"
            ExportStr &= "<td>通俗職類-小類</td>"
            ExportStr &= "<td>訓練職能</td>"

            ExportStr &= "<td>學科辦訓地縣市</td>"
            ExportStr &= "<td>術科辦訓地縣市</td>"
            ExportStr &= "<td>包班種類</td>"   '包班總類 包班種類

            For a As Integer = 0 To Me.ChbExit.Items.Count - 1
                If ChbExit.Items.Item(a).Selected Then
                    Select Case ChbExit.Items.Item(a).Value
                        Case Cst_政府政策性產業_114NOUSE 'Cst_政府政策性產業_108NOUSE '(108年之後不使用此欄)
                            'ExportStr &= "<td>政府政策性產業(108年之後不使用此欄)</td>"
                            'ExportStr &= "<td>政府政策性產業(114年後)</td>"
                            ExportStr &= "<td>5+2產業創新計畫</td>"
                            ExportStr &= "<td>台灣AI行動計畫</td>"
                            ExportStr &= "<td>數位國家創新經濟發展方案</td>"
                            ExportStr &= "<td>國家資通安全發展方案</td>"
                            ExportStr &= "<td>前瞻基礎建設計畫</td>"
                            ExportStr &= "<td>新南向政策</td>"
                        Case Cst_新南向政策
                            ExportStr &= "<td>新南向政策</td>"
                        Case Cst_轄區重點產業
                            ExportStr &= "<td>轄區重點產業</td>"
                        Case Cst_生產力40
                            ExportStr &= "<td>生產力4.0</td>"
                        Case Cst_申請人次
                            ExportStr &= "<td>申請人次</td>"
                        Case Cst_申請補助費
                            ExportStr &= "<td>申請補助費</td>"
                        Case Cst_核定人次
                            ExportStr &= "<td>核定人次</td>"
                        Case Cst_核定補助費
                            ExportStr &= "<td>核定補助費</td>"
                        Case Cst_實際開訓人次
                            ExportStr &= "<td>實際就保開訓人次</td>"
                            ExportStr &= "<td>實際就安開訓人次</td>"
                            'ExportStr &= "<td>實際公務開訓人次</td>"
                            ExportStr &= String.Concat("<td>實際", cst_BUDGET97_N, "開訓人次</td>")'公務ECFA // 協助
                        Case Cst_實際開訓人次加總
                            ExportStr &= "<td>實際合計開訓人次</td>"
                        Case Cst_預估補助費
                            ExportStr &= "<td>就保預估補助費金額</td>"
                            ExportStr &= "<td>就安預估補助費金額</td>"
                            'ExportStr &= "<td>公務預估補助費金額</td>"
                            ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "預估補助費金額</td>") '公務ECFA // 協助
                        Case Cst_預估補助費加總
                            ExportStr &= "<td>合計預估補助費金額</td>"
                        Case Cst_結訓人次 '合計結訓人次
                            ExportStr &= "<td>就保結訓人次</td>"
                            ExportStr &= "<td>就安結訓人次</td>"
                            'ExportStr &= "<td>公務結訓人次</td>"
                            ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "結訓人次</td>") '公務ECFA // 協助
                            ExportStr &= "<td>合計結訓人次</td>"
                        Case Cst_撥款人次 'Key_Identity
                            ExportStr &= "<td>就保撥款人次</td>"
                            ExportStr &= "<td>就安撥款人次</td>"
                            'ExportStr &= "<td>公務撥款人次</td>"
                            ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "撥款人次</td>") '公務ECFA // 協助
                        Case Cst_各身分別撥款人次
                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        ExportStr &= "<td>就保一般身分撥款人次</td>"   '01
                                    Case Else
                                        ExportStr &= "<td>就保特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款人次</td>"   '07
                                End Select
                            Next
                            ExportStr &= "<td>就保特殊身分總撥款人次</td>"

                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        ExportStr &= "<td>就安一般身分撥款人次</td>"   '01
                                    Case Else
                                        ExportStr &= "<td>就安特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款人次</td>"   '07
                                End Select
                            Next
                            ExportStr &= "<td>就安特殊身分總撥款人次</td>"

                            'For Each dr1 As DataRow In dtIdentity.Rows
                            '    Select Case Convert.ToString(dr1("IdentityID"))
                            '        Case "01"
                            '            ExportStr &= "<td>公務一般身分撥款人次</td>"   '01
                            '        Case Else
                            '            ExportStr &= "<td>公務特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款人次</td>"   '07
                            '    End Select
                            'Next
                            'ExportStr &= "<td>公務特殊身分總撥款人次</td>"

                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "一般身分撥款人次</td>") '公務ECFA // 協助 IdentityID:01
                                    Case Else
                                        ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "特殊身分(", Convert.ToString(dr1("NAME")), ")撥款人次</td>") '公務ECFA // 協助 IdentityID:07
                                End Select
                            Next
                            ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "特殊身分總撥款人次</td>") '公務ECFA // 協助

                        Case Cst_撥款補助費
                            ExportStr &= "<td>就保撥款補助費</td>"
                            ExportStr &= "<td>就安撥款補助費</td>"
                            'ExportStr &= "<td>公務撥款補助費</td>"
                            ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "撥款補助費</td>")'公務ECFA // 協助

                        Case Cst_各身分別撥款補助費
                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        ExportStr &= "<td>就保一般身分撥款補助費</td>"   '01
                                    Case Else
                                        ExportStr &= "<td>就保特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款補助費</td>"   '07
                                End Select
                            Next
                            ExportStr &= "<td>就保特殊身分總撥款補助費</td>"

                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        ExportStr &= "<td>就安一般身分撥款補助費</td>"   '01
                                    Case Else
                                        ExportStr &= "<td>就安特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款補助費</td>"   '07
                                End Select
                            Next
                            ExportStr &= "<td>就安特殊身分總撥款補助費</td>"

                            'For Each dr1 As DataRow In dtIdentity.Rows
                            '    Select Case Convert.ToString(dr1("IdentityID"))
                            '        Case "01"
                            '            ExportStr &= "<td>公務一般身分撥款補助費</td>"   ' IdentityID:01
                            '        Case Else
                            '            ExportStr &= "<td>公務特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款補助費</td>"   ' IdentityID:07
                            '    End Select
                            'Next
                            'ExportStr &= "<td>公務特殊身分總撥款補助費</td>"

                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "一般身分撥款補助費</td>") '公務ECFA // 協助 IdentityID:01
                                    Case Else
                                        ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "特殊身分(", Convert.ToString(dr1("NAME")), ")撥款補助費</td>") '公務ECFA // 協助  IdentityID:07
                                End Select
                            Next
                            ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "特殊身分總撥款補助費</td>") '公務ECFA // 協助

                        Case Cst_不預告訪視次數_實地抽訪
                            ExportStr &= "<td>累計不預告實地抽訪次數</td>"
                            ExportStr &= "<td>實地訪視日期</td>"   'ItAPPLYDATE 實地訪視日期
                        Case Cst_不預告訪視次數_電話抽訪
                            ExportStr &= "<td>累計不預告電話抽訪次數</td>"   'CTALL3
                            ExportStr &= "<td>電話訪視日期</td>"   'cuAPPLYDATE3
                            ExportStr &= "<td>累計不預告電話抽訪次數(實地抽訪未到)</td>"   'CTALL4
                            ExportStr &= "<td>電話訪視日期(實地抽訪未到)</td>"  'cuAPPLYDATE4
                        Case Cst_不預告訪視次數_視訊訪查
                            ExportStr &= "<td>累計不預告視訊抽訪次數</td>"   'CTALL3
                            ExportStr &= "<td>視訊訪視日期</td>"   'cuAPPLYDATE3
                        Case Cst_累積訪視異常次數
                            ExportStr &= "<td>累計不預告實地抽訪異常次數</td>"   'vitN 累計不預告實地抽訪異常次數
                            ExportStr &= "<td>累計不預告視訊抽訪異常次數</td>"   'vitTVN 累計不預告視訊抽訪異常次數
                            ExportStr &= "<td>累計不預告電話抽訪異常次數</td>"   'VitTelN 累計不預告電話抽訪異常次數

                        Case Cst_累計訪視異常原因
                            ExportStr &= "<td>出席率不佳</td>"   'It22b01N
                            ExportStr &= "<td>簽到退未落實</td>"   'It22b02N
                            ExportStr &= "<td>師資不符</td>"   'It22b03N
                            ExportStr &= "<td>助教不符</td>"   'It22b06N
                            ExportStr &= "<td>課程內容不符</td>"   'It22b04N
                            ExportStr &= "<td>上課地點不符</td>"   'It22b05N

                            ExportStr &= "<td>其他</td>"   'IT22B99NOTE 累計訪視異常原因 其他
                            ExportStr &= "<td>其他補充說明</td>"   'LITEM23NOTE 累計訪視異常原因 其他補充說明
                        'Case Cst_會計查帳次數 ExportStr &= "<td>會計查帳次數</td>"  
                        Case Cst_離訓人次
                            ExportStr &= "<td>離訓人次</td>"
                        Case Cst_退訓人次
                            ExportStr &= "<td>退訓人次</td>"
                        Case Cst_訓練時數
                            ExportStr &= "<td>訓練時數</td>"

                        Case Cst_固定費用總額
                            ExportStr &= "<td>固定費用總額</td>"
                        Case Cst_固定費用單一人時成本
                            ExportStr &= "<td>固定費用單一人時成本</td>"
                        Case Cst_人時成本超出原因說明
                            ExportStr &= "<td>人時成本超出原因說明</td>"
                        Case Cst_材料費總額
                            ExportStr &= "<td>材料費總額</td>"
                        Case Cst_材料費占比
                            ExportStr &= "<td>材料費占比</td>"
                        Case Cst_超出材料費比率上限原因說明
                            ExportStr &= "<td>超出材料費比率上限原因說明</td>"
                            'Case Cst_人時成本
                            'ExportStr &= "<td>人時成本</td>"  
                        Case Cst_上課時間
                            ExportStr &= "<td>上課時間</td>"
                        Case Cst_撥款日期
                            ExportStr &= "<td>撥款日期</td>"

                        Case Cst_包班事業單位
                            ExportStr &= "<td>包班事業單位</td>"
                        Case Cst_師資名單
                            ExportStr &= "<td>師資名單</td>"
                        Case Cst_上課地址及教室
                            ExportStr &= "<td>上課地址及教室</td>"
                            'Case Cst_上課地址及教室2
                            '    ExportStr &= "<td>上課地址及教室2</td>"  
                        Case Cst_包班事業單位保險證號
                            ExportStr &= "<td>包班事業單位保險證號</td>"
                        Case Cst_包班事業單位統一編號
                            ExportStr &= "<td>包班事業單位統一編號</td>"
                        Case Cst_公務ECFA性別人數 '公務ECFA // 協助
                            ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "男性人數</td>") '公務ECFA // 協助
                            ExportStr &= String.Concat("<td>", cst_BUDGET97_N, "女性人數</td>") '公務ECFA // 協助

                        Case cst_課程申請流水號
                            ExportStr &= "<td>課程申請流水號</td>"
                        Case cst_上架日期
                            ExportStr &= "<td>上架日期</td>"
                        Case cst_開放報名結束日期
                            ExportStr &= "<td>開放報名日期</td>"
                            ExportStr &= "<td>結束報名日期</td>"
                        Case cst_課程備註
                            ExportStr &= "<td>課程備註1</td>"
                            ExportStr &= "<td>課程備註2</td>"
                        Case cst_術科時數
                            ExportStr &= "<td>術科時數</td>"
                        Case cst_聯絡人
                            ExportStr &= "<td>聯絡人</td>"
                        Case cst_聯絡電話
                            ExportStr &= "<td>聯絡電話</td>"
                        Case cst_是否停辦
                            ExportStr &= "<td>是否停辦</td>"
                        Case cst_iCAP標章證號及效期
                            ExportStr &= "<td>iCAP標章證號</td>"
                            ExportStr &= "<td>iCAP效期</td>"
                            'Case cst_政策性產業課程可辦理班數
                            '    ExportStr &= "<td>" & T_YR1 & "</td>"  
                            '    ExportStr &= "<td>" & T_YR2 & "</td>"  
                            '    ExportStr &= "<td>" & T_YR3 & "</td>"  
                        Case Cst_辦理方式
                            ExportStr &= "<td>辦理方式</td>"
                        Case Cst_實際開訓性別人數
                            ExportStr &= "<td>實際開訓男性人數</td>"
                            ExportStr &= "<td>實際開訓女性人數</td>"
                        Case Cst_行政管理疏失重大異常狀況 '53
                            ExportStr &= "<td>行政管理疏失/重大異常狀況</td>"
                        Case Cst_報名繳費方式 '54 'ENTERSUPPLYSTYLE
                            ExportStr &= "<td>報名繳費方式</td>"
                    End Select
                End If
            Next
        End If
        ExportStr &= "</tr>" & vbCrLf
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        sbHTML.Append(ExportStr)

        '建立資料面
        For Each dr As DataRow In dt.Rows
            ExportStr = "<tr>"
            'sbHTML.Append("<tr>")
            If ChbExit.SelectedIndex = -1 Then '匯出欄位 未選
                ExportStr &= String.Concat("<td>", dr("YEARS"), "</td>")    '計畫年度
                ExportStr &= String.Concat("<td>", dr("OrgPlanName2"), "</td>")    '自主/產投(產投計畫別)
                ExportStr &= String.Concat("<td>", dr("DistName"), "</td>")    '分署
                ExportStr &= "<td>" & Convert.ToString(dr("OrgTypeName")) & "</td>"      '單位屬性
                ExportStr &= "<td>" & dr("orgname") & "</td>"    '訓練機構
                ExportStr &= "<td>" & dr("ClassName") & "</td>"    '班別名稱
                If hid_ssYears.Value >= cst_y2019 Then 'sCaseYears
                    ExportStr &= "<td>" & Convert.ToString(dr("FIRSTSORT")) & "</td>"   '提案意願順序
                End If
                ExportStr &= "<td>" & dr("CyclType") & "</td>"    '期別
                ExportStr &= "<td>" & dr("ClassID") & "</td>"    '課程代碼
                If gflag_SHOW_APPSTAGE Then 'APPSTAGE
                    ExportStr &= "<td>" & Convert.ToString(dr("APPSTAGE")) & "</td>"   '申請階段
                End If

                ExportStr &= "<td>" & dr("STDate") & "</td>"    '開訓日期
                ExportStr &= "<td>" & dr("FDDate") & "</td>"    '結訓日期
                ExportStr &= "<td>" & dr("Pkname12") & "</td>"    '課程分類

                'If gflag_SHOW_2019_1 Then
                '政府政策性產業「5+2」產業創新計畫 5+2產業【台灣AI行動計畫】 KID='08'【數位國家創新經濟發展方案】KID='09'【國家資通安全發展方案】KID='10'【前瞻基礎建設計畫】【新南向政策】KID='19'2019年啟用 work2019x01:2019 政府政策性產業
                '    'Dim vKID20_1 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 1)
                '    'Dim vKID20_2 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 2)
                '    'Dim vKID20_3 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 3)
                '    'Dim vKID20_4 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 4)
                '    'Dim vKID20_5 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 5)
                '    'Dim vKID20_6 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 6)
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME1")) & "</td>"   '5+2產業創新計畫
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME2")) & "</td>"   '台灣AI行動計畫
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME3")) & "</td>"   '"數位國家創新經濟發展方案</td>"  
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME4")) & "</td>"   '"國家資通安全發展方案</td>"  
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME5")) & "</td>"   '"前瞻基礎建設計畫</td>"  
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME6")) & "</td>"   '"新南向政策</td>"  
                'End If
                '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
                '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
                If fg_Work2026x02 Then
                    'Dim KID26STR1 As String() = {"五大信賴產業推動方案", "六大區域產業及生活圈", "智慧國家2.0綱領", "新南向政策推動計畫", "國家人才競爭力躍升方案", "AI新十大建設推動方案", "台灣AI行動計畫2.0", "智慧機器人產業推動方案", "臺灣2050淨零轉型"}
                    For i_kid26 As Integer = 1 To 9
                        ExportStr &= $"<td>{dr($"D26KNAME{i_kid26}")}</td>"
                    Next
                End If
                If gfg_SHOW_2025_1 Then
                    '亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
                    ExportStr &= String.Concat("<td>", dr("D25KNAME1"), "</td>") '亞洲矽谷
                    ExportStr &= String.Concat("<td>", dr("D25KNAME2"), "</td>") '重點產業</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME3"), "</td>") '台灣AI行動計畫</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME4"), "</td>") '智慧國家方案</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME5"), "</td>") '國家人才競爭力躍升方案</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME6"), "</td>") '新南向政策</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME7"), "</td>") 'AI加值應用</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME8"), "</td>") '職場續航</td>"
                End If
                'dd.KID22,dd.KNAME22 進階政策性產業類別
                ExportStr &= "<td>" & Convert.ToString(dr("KNAME22")) & "</td>"

                '20100114 add andy
                ExportStr &= "<td>" & dr("kname1") & "</td>"    '※六大新興產業
                ExportStr &= "<td>" & dr("kname2") & "</td>"    '※十大重點服務業	
                'Select Case hid_ssYears.Value 'sCaseYears
                '    Case cst_y2017
                '    Case Else
                '        ExportStr &= "<td>" & dr("kname3") & "</td>"   '※四大新興智慧型產業
                'End Select
                ExportStr &= "<td>" & dr("kname3") & "</td>"   '※四大新興智慧型產業

                ExportStr &= "<td>" & dr("GCodeName") & "</td>"    '訓練業別編碼
                ExportStr &= "<td>" & Convert.ToString(dr("GCNAME")) & "</td>"    '訓練業別
                'ExportStr &= "<td>" & Convert.ToString(dr("TJOBNAME")) & "</td>"    '訓練業別(職訓業別)

                sCJOB_UNKEY = Convert.ToString(dr("CJOB_UNKEY"))
                ExportStr &= "<td>" & TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 1) & "</td>"    '通俗職類-大類
                ExportStr &= "<td>" & TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 2) & "</td>"    '通俗職類-小類
                ExportStr &= "<td>" & Convert.ToString(dr("CCName")) & "</td>"     '訓練職能

                ExportStr &= "<td>" & dr("AddressSciPTID") & "</td>"    '學科辦訓地縣市
                ExportStr &= "<td>" & dr("AddressTechPTID") & "</td>"   '術科辦訓地縣市
                'ExportStr &= "<td>" & dr("PackageType") & "</td>"    '包班種類
                ExportStr &= "<td>" & dr("PackageTypeN") & "</td>"    '包班總類 包班種類 

            Else
                ExportStr &= String.Concat("<td>", dr("YEARS"), "</td>")    '計畫年度
                ExportStr &= String.Concat("<td>", dr("OrgPlanName2"), "</td>")    '自主/產投(產投計畫別)
                ExportStr &= String.Concat("<td>", dr("DistName"), "</td>")    '分署
                ExportStr &= "<td>" & Convert.ToString(dr("OrgTypeName")) & "</td>"      '單位屬性
                ExportStr &= "<td>" & dr("orgname") & "</td>"    '訓練機構
                For a As Integer = 0 To Me.ChbExit.Items.Count - 1
                    If ChbExit.Items.Item(a).Selected Then
                        Select Case ChbExit.Items.Item(a).Value
                            Case Cst_統一編號
                                ExportStr &= String.Concat("<td>", dr("ComIDNO"), "</td>")
                            Case Cst_立案縣市
                                ExportStr &= String.Concat("<td>", dr("CTName2"), "</td>")
                        End Select
                    End If
                Next
                ExportStr &= "<td>" & dr("ClassName") & "</td>"    '班別名稱
                ExportStr &= "<td>" & dr("CyclType") & "</td>"    '期別
                ExportStr &= "<td>" & dr("ClassID") & "</td>"    '課程代碼
                If gflag_SHOW_APPSTAGE Then 'APPSTAGE
                    ExportStr &= String.Concat("<td>", dr("APPSTAGE"), "</td>") '申請階段
                End If

                ExportStr &= "<td>" & dr("STDate") & "</td>"    '開訓日期
                ExportStr &= "<td>" & dr("FDDate") & "</td>"    '結訓日期
                ExportStr &= "<td>" & dr("Pkname12") & "</td>"    '課程分類

                'If gflag_SHOW_2019_1 Then
                '    '政府政策性產業'「5+2」產業創新計畫 5+2產業'【台灣AI行動計畫】 KID='08''【數位國家創新經濟發展方案】KID='09''【國家資通安全發展方案】KID='10''【前瞻基礎建設計畫】'【新南向政策】KID='19'
                '    '2019年啟用 work2019x01:2019 政府政策性產業
                '    'Dim vKID20_1 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 1)
                '    'Dim vKID20_2 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 2)
                '    'Dim vKID20_3 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 3)
                '    'Dim vKID20_4 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 4)
                '    'Dim vKID20_5 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 5)
                '    'Dim vKID20_6 As String = TIMS.GET_NAME_KID20(Convert.ToString(dr("KID20")), dtKID_N20, 6)
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME1")) & "</td>"   '5+2產業創新計畫
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME2")) & "</td>"   '台灣AI行動計畫
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME3")) & "</td>"   '"數位國家創新經濟發展方案</td>"  
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME4")) & "</td>"   '"國家資通安全發展方案</td>"  
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME5")) & "</td>"   '"前瞻基礎建設計畫</td>"  
                '    ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME6")) & "</td>"   '"新南向政策</td>"  
                'End If
                '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
                '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
                If fg_Work2026x02 Then
                    'Dim KID26STR1 As String() = {"五大信賴產業推動方案", "六大區域產業及生活圈", "智慧國家2.0綱領", "新南向政策推動計畫", "國家人才競爭力躍升方案", "AI新十大建設推動方案", "台灣AI行動計畫2.0", "智慧機器人產業推動方案", "臺灣2050淨零轉型"}
                    For i_kid26 As Integer = 1 To 9
                        ExportStr &= $"<td>{dr($"D26KNAME{i_kid26}")}</td>"
                    Next
                End If
                If gfg_SHOW_2025_1 Then
                    '亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
                    ExportStr &= String.Concat("<td>", dr("D25KNAME1"), "</td>") '亞洲矽谷
                    ExportStr &= String.Concat("<td>", dr("D25KNAME2"), "</td>") '重點產業</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME3"), "</td>") '台灣AI行動計畫</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME4"), "</td>") '智慧國家方案</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME5"), "</td>") '國家人才競爭力躍升方案</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME6"), "</td>") '新南向政策</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME7"), "</td>") 'AI加值應用</td>"
                    ExportStr &= String.Concat("<td>", dr("D25KNAME8"), "</td>") '職場續航</td>"
                End If
                'dd.KID22,dd.KNAME22 進階政策性產業類別
                ExportStr &= "<td>" & Convert.ToString(dr("KNAME22")) & "</td>"

                '20100114 add andy
                ExportStr &= "<td>" & dr("kname1") & "</td>"   '※六大新興產業	
                ExportStr &= "<td>" & dr("kname2") & "</td>"   '※十大重點服務業
                'Select Case hid_ssYears.Value 'sCaseYears
                '    Case cst_y2017
                '    Case Else
                '        ExportStr &= "<td>" & dr("kname3") & "</td>"    '※四大新興智慧型產業
                'End Select
                ExportStr &= "<td>" & dr("kname3") & "</td>"    '※四大新興智慧型產業

                ExportStr &= "<td>" & dr("GCodeName") & "</td>"     '訓練業別編碼
                ExportStr &= "<td>" & Convert.ToString(dr("GCNAME")) & "</td>"    '訓練業別
                'ExportStr &= "<td>" & Convert.ToString(dr("TJOBNAME")) & "</td>"    '訓練業別(職訓業別)

                sCJOB_UNKEY = Convert.ToString(dr("CJOB_UNKEY"))
                ExportStr &= "<td>" & TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 1) & "</td>"    '通俗職類-大類
                ExportStr &= "<td>" & TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 2) & "</td>"    '通俗職類-小類
                ExportStr &= "<td>" & Convert.ToString(dr("CCName")) & "</td>"     '訓練職能

                ExportStr &= "<td>" & dr("AddressSciPTID") & "</td>"    '學科辦訓地縣市
                ExportStr &= "<td>" & dr("AddressTechPTID") & "</td>"   '術科辦訓地縣市
                'ExportStr &= "<td>" & dr("PackageType") & "</td>"    '包班種類
                ExportStr &= "<td>" & dr("PackageTypeN") & "</td>"    '包班總類 包班種類

                For j As Integer = 0 To Me.ChbExit.Items.Count - 1
                    If Me.ChbExit.Items(j).Selected Then
                        Select Case ChbExit.Items.Item(j).Value
                            Case Cst_政府政策性產業_114NOUSE 'Cst_政府政策性產業_108NOUSE '(108年之後不使用此欄)
                                'If KID_17.Visible Then
                                '    ExportStr &= "<td>" & dr("KNAME17") & "</td>"    '政府政策性產業(108年之後不使用此欄)
                                'Else
                                '    ExportStr &= "<td>" & dr("KNAME19") & "</td>"    '政府政策性產業(108年之後不使用此欄)
                                'End If
                                ExportStr &= String.Concat("<td>", dr("D20KNAME1"), "</td>") '5+2產業創新計畫
                                ExportStr &= String.Concat("<td>", dr("D20KNAME2"), "</td>") '台灣AI行動計畫
                                ExportStr &= String.Concat("<td>", dr("D20KNAME3"), "</td>") '"數位國家創新經濟發展方案</td>"  
                                ExportStr &= String.Concat("<td>", dr("D20KNAME4"), "</td>") '"國家資通安全發展方案</td>"  
                                ExportStr &= String.Concat("<td>", dr("D20KNAME5"), "</td>") '"前瞻基礎建設計畫</td>"  
                                ExportStr &= String.Concat("<td>", dr("D20KNAME6"), "</td>") '"新南向政策</td>"  
                            Case Cst_新南向政策
                                ExportStr &= "<td>" & dr("KNAME18") & "</td>"    '新南向政策
                            Case Cst_轄區重點產業
                                '空白或新年度，使用新欄位
                                Dim s_KNAME1315 As String = Convert.ToString(dr("KNAME13"))
                                If s_KNAME1315 = "" OrElse hid_ssYears.Value >= cst_y2018 Then s_KNAME1315 = Convert.ToString(dr("KNAME15"))
                                ExportStr &= "<td>" & s_KNAME1315 & "</td>"    '轄區重點產業
                            Case Cst_生產力40
                                ExportStr &= "<td>" & dr("KNAME14") & "</td>"    '生產力4.0
                            Case Cst_申請人次
                                ExportStr &= "<td class=""noDecFormat"">" & dr("ATNum") & "</td>"    '申請人數
                            Case Cst_申請補助費
                                ExportStr &= "<td class=""noDecFormat"">" & dr("ADefGovCost") & "</td>"    '申請補助費
                            Case Cst_核定人次
                                ExportStr &= "<td class=""noDecFormat"">" & dr("TNum") & "</td>"    '核定人數
                            Case Cst_核定補助費
                                ExportStr &= "<td class=""noDecFormat"">" & dr("DefGovCost") & "</td>"   '核定補助費
                            Case Cst_實際開訓人次
                                ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcount1") & "</td>"
                                ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcount2") & "</td>"
                                'ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcount3") & "</td>"'/*開訓人次-公務*/
                                ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcount97") & "</td>"
                            Case Cst_實際開訓人次加總 'openstudcountall
                                ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcountall") & "</td>"    '開訓人次加總
                            Case Cst_預估補助費
                                ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("cost1"), 2) & "</td>"
                                ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("cost2"), 2) & "</td>"
                                'ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("cost3"), 2) & "</td>" '/*預估補助費-公務*/
                                ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("cost97"), 2) & "</td>"
                            Case Cst_預估補助費加總
                                ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("costAll"), 2) & "</td>"   '預估補助費加總
                            Case Cst_結訓人次 '合計結訓人次
                                ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcout03") & "</td>"
                                ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcout02") & "</td>"
                                'ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcout01") & "</td>"
                                ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcout97") & "</td>"
                                ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcoutall") & "</td>"

                            Case Cst_撥款人次
                                ExportStr &= "<td class=""noDecFormat"">" & dr("budcountall3") & "</td>"
                                ExportStr &= "<td class=""noDecFormat"">" & dr("budcountall2") & "</td>"
                                'ExportStr &= "<td class=""noDecFormat"">" & dr("budcountall3") & "</td>"'/*公務合計撥款人次*/
                                ExportStr &= "<td class=""noDecFormat"">" & dr("budcountall97") & "</td>"
                            Case Cst_各身分別撥款人次
                                For Each dr1 As DataRow In dtIdentity.Rows
                                    ExportStr &= "<td class=""noDecFormat"">" & dr("bud03count" & dr1("IdentityID")) & "</td>"
                                Next
                                ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudcount01") & "</td>"

                                For Each dr1 As DataRow In dtIdentity.Rows
                                    ExportStr &= "<td class=""noDecFormat"">" & dr("bud02count" & dr1("IdentityID")) & "</td>"
                                Next
                                ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudcount02") & "</td>"

                                'For Each dr1 As DataRow In dtIdentity.Rows
                                '    ExportStr &= "<td class=""noDecFormat"">" & dr("bud01count" & dr1("IdentityID")) & "</td>"
                                'Next
                                'ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudcount03") & "</td>"

                                For Each dr1 As DataRow In dtIdentity.Rows
                                    ExportStr &= "<td class=""noDecFormat"">" & dr("bud97count" & dr1("IdentityID")) & "</td>"
                                Next
                                ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudcount97") & "</td>"

                            Case Cst_撥款補助費
                                ExportStr &= "<td class=""noDecFormat"">" & dr("budmoneyall3") & "</td>"
                                ExportStr &= "<td class=""noDecFormat"">" & dr("budmoneyall2") & "</td>"
                                'ExportStr &= "<td class=""noDecFormat"">" & dr("budmoneyall3") & "</td>" '/*公務合計撥款金額*/
                                ExportStr &= "<td class=""noDecFormat"">" & dr("budmoneyall97") & "</td>"
                            Case Cst_各身分別撥款補助費
                                For Each dr1 As DataRow In dtIdentity.Rows
                                    ExportStr &= "<td class=""noDecFormat"">" & dr("bud03money" & dr1("IdentityID")) & "</td>"
                                Next
                                ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudmoney01") & "</td>"

                                For Each dr1 As DataRow In dtIdentity.Rows
                                    ExportStr &= "<td class=""noDecFormat"">" & dr("bud02money" & dr1("IdentityID")) & "</td>"
                                Next
                                ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudmoney02") & "</td>"

                                'For Each dr1 As DataRow In dtIdentity.Rows
                                '    ExportStr &= "<td class=""noDecFormat"">" & dr("bud01money" & dr1("IdentityID")) & "</td>"
                                'Next
                                'ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudmoney03") & "</td>"

                                For Each dr1 As DataRow In dtIdentity.Rows
                                    ExportStr &= "<td class=""noDecFormat"">" & dr("bud97money" & dr1("IdentityID")) & "</td>"
                                Next
                                ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudmoney97") & "</td>"

                            Case Cst_不預告訪視次數_實地抽訪
                                ExportStr &= "<td class=""noDecFormat"">" & dr("cuall") & "</td>"   '不預告訪視次數-實地抽訪
                                Dim ItAPPLYDATE As String = If(Convert.ToString(dr("ItAPPLYDATE")) <> "", Replace(Convert.ToString(dr("ItAPPLYDATE")), ",", ";"), "") ' Convert.ToString(dr("ItAPPLYDATE"))
                                ExportStr &= "<td>" & ItAPPLYDATE & "</td>"   'Cst_不預告訪視次數_實地抽訪 訪視日期

                            Case Cst_不預告訪視次數_電話抽訪
                                '僅計算電話抽訪原因為非「實地抽訪時未到」的件數
                                ExportStr &= "<td class=""noDecFormat"">" & dr("CTALL3") & "</td>"   '不預告訪視次數-電話抽訪
                                Dim cuAPPLYDATE3 As String = If(Convert.ToString(dr("cuAPPLYDATE3")) <> "", Replace(Convert.ToString(dr("cuAPPLYDATE3")), ",", ";"), "")
                                ExportStr &= "<td>" & cuAPPLYDATE3 & "</td>"   '訪視日期

                                '僅計算電話抽訪原因=「實地抽訪時未到」的件數
                                ExportStr &= "<td class=""noDecFormat"">" & dr("CTALL4") & "</td>"   '不預告訪視次數-電話抽訪
                                Dim cuAPPLYDATE4 As String = If(Convert.ToString(dr("cuAPPLYDATE4")) <> "", Replace(Convert.ToString(dr("cuAPPLYDATE4")), ",", ";"), "") 'Convert.ToString(dr("cuAPPLYDATE4"))
                                ExportStr &= "<td>" & cuAPPLYDATE4 & "</td>"   '訪視日期

                            Case Cst_不預告訪視次數_視訊訪查
                                ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", dr("VW2CNT"))    '累計不預告視訊抽訪次數 視訊訪查 VW2CNT
                                Dim vw2APPLYDATE As String = If(Convert.ToString(dr("VW2APPLYDATE")) <> "", Replace(Convert.ToString(dr("VW2APPLYDATE")), ",", ";"), "") 'Convert.ToString(dr("cuAPPLYDATE4"))
                                ExportStr &= String.Format("<td>{0}</td>", vw2APPLYDATE)    '視訊訪視日期 視訊訪查 VW2APPLYDATE

                            Case Cst_累積訪視異常次數
                                ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", dr("vitN"))   'vitN 累計不預告實地抽訪異常次數
                                ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", dr("vitTVN"))   'vitTVN 累計不預告視訊抽訪異常次數
                                ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", dr("VitTelN"))   'VitTelN 累計不預告電話抽訪異常次數

                            Case Cst_累計訪視異常原因
                                ExportStr &= "<td class=""noDecFormat"">" & dr("It22b01N") & "</td>"     '累計訪視異常原因-出席率不佳
                                ExportStr &= "<td class=""noDecFormat"">" & dr("It22b02N") & "</td>"     '累計訪視異常原因-簽到退未落實
                                ExportStr &= "<td class=""noDecFormat"">" & dr("It22b03N") & "</td>"     '累計訪視異常原因-師資不符
                                ExportStr &= "<td class=""noDecFormat"">" & dr("It22b06N") & "</td>"     '累計訪視異常原因-助教不符
                                ExportStr &= "<td class=""noDecFormat"">" & dr("It22b04N") & "</td>"     '累計訪視異常原因-課程內容不符
                                ExportStr &= "<td class=""noDecFormat"">" & dr("It22b05N") & "</td>"     '累計訪視異常原因-上課地點不符
                                ExportStr &= "<td>" & Convert.ToString(dr("IT22B99NOTE")) & "</td>"      '累計訪視異常原因-其他
                                ExportStr &= "<td>" & Convert.ToString(dr("LITEM23NOTE")) & "</td>"      '累計訪視異常原因-其他補充說明
                            'Case Cst_會計查帳次數 ExportStr &= "<td class=""noDecFormat"">" & "" & "</td>"   '會計查帳次數
                            Case Cst_離訓人次
                                ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("std_cnt2")) & "</td>"     '離訓人次
                            Case Cst_退訓人次
                                ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("std_cnt3")) & "</td>"     '退訓人次
                            Case Cst_訓練時數
                                ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("THours")) & "</td>"     '訓練時數
                            Case Cst_固定費用總額
                                ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("FIXSUMCOST")) & "</td>"
                            Case Cst_固定費用單一人時成本
                                Select Case hid_ssYears.Value 'sCaseYears
                                    Case Is >= cst_y2018
                                        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("ACTHUMCOST")) & "</td>"
                                    Case Else
                                        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("PHCOST")) & "</td>"
                                End Select

                            Case Cst_人時成本超出原因說明
                                ExportStr &= "<td>" & Convert.ToString(dr("FIXExceeDesc")) & "</td>"
                            Case Cst_材料費總額
                                ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("METSUMCOST")) & "</td>"
                            Case Cst_材料費占比
                                ExportStr &= "<td>" & Convert.ToString(dr("METCOSTPER")) & "</td>"
                            Case Cst_超出材料費比率上限原因說明
                                ExportStr &= "<td>" & Convert.ToString(dr("METExceeDesc")) & "</td>"
                                'Case Cst_人時成本
                                'ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("PhCost")) & "</td>"     '人時成本
                            Case Cst_上課時間
                                ExportStr &= "<td>" & Convert.ToString(dr("WEEKSTIME")) & "</td>"     '上課時間
                            Case Cst_撥款日期
                                ExportStr &= "<td>" & Convert.ToString(dr("AllotDate")) & "</td>"

                            Case Cst_包班事業單位
                                ExportStr &= "<td>" & Convert.ToString(dr("BusPackage")) & "</td>"
                            Case Cst_師資名單
                                ExportStr &= "<td>" & Convert.ToString(dr("PlanTeacher")) & "</td>"
                            Case Cst_上課地址及教室
                                Dim spAddress As String = "" '組合地址用
                                '組合 Cst_上課地址及教室
                                spAddress = Get_AddressPlaceName(dtZip, dr)
                                ExportStr &= "<td>" & spAddress & "</td>"
                            Case Cst_包班事業單位保險證號
                                ExportStr &= "<td>" & Convert.ToString(dr("BusPackage2")) & "</td>"
                            Case Cst_包班事業單位統一編號
                                ExportStr &= "<td>" & Convert.ToString(dr("BusPackage3")) & "</td>"
                            Case Cst_公務ECFA性別人數 '公務ECFA // 協助
                                ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("SexNumxM")) & "</td>"    '/*男性人數*/
                                ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("SexNumxF")) & "</td>"    '/*女性人數*/
                            Case cst_課程申請流水號
                                ExportStr &= "<td>" & Convert.ToString(dr("PSNO28")) & "</td>"
                            Case cst_上架日期
                                ExportStr &= "<td>" & Convert.ToString(dr("ONSHELLDATE")) & "</td>"
                            Case cst_開放報名結束日期
                                ExportStr &= "<td>" & Convert.ToString(dr("SENTERDATE")) & "</td>"
                                ExportStr &= "<td>" & Convert.ToString(dr("FENTERDATE")) & "</td>"
                            Case cst_課程備註
                                ExportStr &= "<td>" & Convert.ToString(dr("memo8")) & "</td>"
                                ExportStr &= "<td>" & Convert.ToString(dr("memo82")) & "</td>"
                            Case cst_術科時數 'ProTechHours
                                ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("ProTechHours")) & "</td>"
                            Case cst_聯絡人
                                ExportStr &= "<td>" & Convert.ToString(dr("ContactName")) & "</td>"
                            Case cst_聯絡電話
                                Dim s_ContactPhone As String = Convert.ToString(dr("ContactPhone"))
                                Dim s_ContactMobile As String = Convert.ToString(dr("ContactMobile"))
                                Dim s_ContactPhoneMobile As String = s_ContactMobile
                                If s_ContactPhone <> "" AndAlso s_ContactMobile <> "" Then
                                    s_ContactPhoneMobile = String.Concat(s_ContactPhone, "、", s_ContactMobile)
                                ElseIf s_ContactPhone <> "" Then
                                    s_ContactPhoneMobile = s_ContactPhone
                                End If
                                ExportStr &= String.Concat("<td>", s_ContactPhoneMobile, "</td>")
                            Case cst_是否停辦
                                ExportStr &= "<td>" & Convert.ToString(dr("NotOpenN")) & "</td>"
                            Case cst_iCAP標章證號及效期 'ICAPNUM-iCAP標章證號
                                ExportStr &= "<td>" & Convert.ToString(dr("ICAPNUM")) & "</td>"
                                ExportStr &= "<td>" & Convert.ToString(dr("iCAPMARKDATE")) & "</td>"
                                'Case cst_政策性產業課程可辦理班數
                                '    ExportStr &= "<td>" & Convert.ToString(dr("PCNT11")) & "</td>"  
                                '    ExportStr &= "<td>" & Convert.ToString(dr("PCNT12")) & "</td>"  
                                '    ExportStr &= "<td>" & Convert.ToString(dr("PCNT13")) & "</td>"  
                            Case Cst_辦理方式
                                ExportStr &= String.Format("<td>{0}</td>", dr("DISTANCE_N"))
                            Case Cst_實際開訓性別人數
                                ExportStr &= String.Format("<td>{0}</td>", dr("SexCNTM"))
                                ExportStr &= String.Format("<td>{0}</td>", dr("SexCNTF"))
                            Case Cst_行政管理疏失重大異常狀況 '53
                                ExportStr &= String.Format("<td>{0}</td>", dr("VERIFYDATE")) '為分署於該功能登打之經查核確認日期。
                            Case Cst_報名繳費方式 '54 'ENTERSUPPLYSTYLE
                                ExportStr &= String.Format("<td>{0}</td>", dr("ENTERSUPPLYSTYLE_N")) '"<td>報名繳費方式</td>"

                        End Select
                    End If
                Next
            End If
            ExportStr &= "</tr>" & vbCrLf
            'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
            sbHTML.Append(ExportStr)
        Next
        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        '"<meta http-equiv='Content-Type' content='application/vnd.ms-excel; charset=UTF-8'>"
        'Dim strMETAXLS As String = "<meta http-equiv=Content-Type content=""text/html; charset=big5"">"
        'Const cst_meta_big5 As String = "<meta http-equiv=Content-Type content=""application/vnd.ms-excel; charset=big5"">"
        'Const cst_meta_utf8 As String = "<meta http-equiv='Content-Type' content='application/vnd.ms-excel; charset=UTF-8'>"
        'Dim v_RBL_CharsetType As String = TIMS.GetListValue(RBL_CharsetType)
        'Dim strMETAXLS As String = If(v_RBL_CharsetType = "BIG5", cst_meta_big5, cst_meta_utf8)
        'Dim s_Charset As String = If(v_RBL_CharsetType = "BIG5", TIMS.cst_Charset_big5, TIMS.cst_Charset_UTF8)

        'Dim parmsExp As New Hashtable
        parmsExp.Clear()
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", s_FileName1)
        'parmsExp.Add("strMETAXLS", strMETAXLS)
        'parmsExp.Add("Charset", s_Charset)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        'strHTML &= ("</body>")
        flagExp1 = True '匯出全域變數判斷
    End Sub

    Function Get_SqlExpRptXLSX() As String
        'Dim response As HttpResponse = HttpResponse(,) '(FileIO(path_to_file).read(), mimetype ='application/force-download')
        'Dim Reason As String '儲存錯誤的原因
        Dim dtW As New DataTable '儲存錯誤資料的DataTable
        'dtWrong.Columns.Add(New DataColumn("Index"))
        'dtWrong.Columns.Add(New DataColumn("COMIDNO"))
        'dtWrong.Columns.Add(New DataColumn("Reason"))
        'Dim ExportStr As String = ""
        Dim sql As String = ""
        sql &= " SELECT a.PLANID,a.COMIDNO,a.SEQNO,a.CJOB_UNKEY" & vbCrLf 'PK
        '/* '組合 Cst_上課地址及教室*/
        sql &= " ,a.S1PLACEID" & vbCrLf
        sql &= " ,a.S2PLACEID" & vbCrLf
        sql &= " ,a.T1PLACEID" & vbCrLf
        sql &= " ,a.T2PLACEID" & vbCrLf
        'sql &= " ,a.SPPLACEID" & vbCrLf
        'sql &= " ,a.TPPLACEID" & vbCrLf
        sql &= " ,a.S1PTID" & vbCrLf
        sql &= " ,a.S2PTID" & vbCrLf
        sql &= " ,a.T1PTID" & vbCrLf
        sql &= " ,a.T2PTID" & vbCrLf
        'sql &= " ,a.SPPTID" & vbCrLf
        'sql &= " ,a.TPPTID" & vbCrLf
        sql &= " ,a.S1PLACENAME"
        sql &= " ,a.S2PLACENAME"
        sql &= " ,a.T1PLACENAME"
        sql &= " ,a.T2PLACENAME" & vbCrLf
        'sql &= " ,a.SPPLACENAME" & vbCrLf
        'sql &= " ,a.TPPLACENAME" & vbCrLf
        sql &= " ,a.S1ZIPCODE"
        sql &= " ,a.S1ADDRESS"
        sql &= " ,a.S1ZIP6W" & vbCrLf
        sql &= " ,a.S2ZIPCODE"
        sql &= " ,a.S2ADDRESS"
        sql &= " ,a.S2ZIP6W" & vbCrLf
        sql &= " ,a.T1ZIPCODE"
        sql &= " ,a.T1ADDRESS"
        sql &= " ,a.T1ZIP6W" & vbCrLf
        sql &= " ,a.T2ZIPCODE"
        sql &= " ,a.T2ADDRESS"
        sql &= " ,a.T2ZIP6W" & vbCrLf

        If ChbExit.SelectedIndex = -1 Then
            'ExportStr &= "<td>分署</td>"  
            sql &= " ,a.YEARS 計畫年度" & vbCrLf
            sql &= " ,a.OrgPlanName2 計畫別" & vbCrLf  '自主/產投(產投計畫別)
            sql &= " ,a.DistName 各分署" & vbCrLf
            sql &= " ,a.OrgTypeName 單位屬性" & vbCrLf    '單位屬性
            sql &= " ,a.orgname 訓練機構" & vbCrLf  '訓練機構
            sql &= " ,a.ClassName 班別名稱" & vbCrLf  '班別名稱
            sql &= " ,a.FIRSTSORT 提案意願順序" & vbCrLf  '提案意願順序
            sql &= " ,a.CyclType 期別" & vbCrLf  '期別
            sql &= " ,a.ClassID 課程代碼" & vbCrLf  '課程代碼
            'APPSTAGE
            If gflag_SHOW_APPSTAGE Then sql &= " ,a.APPSTAGE 申請階段" & vbCrLf '申請階段
            sql &= " ,format(a.STDate,'yyyy/MM/dd') 開訓日期" & vbCrLf  '開訓日期
            sql &= " ,format(a.FDDate,'yyyy/MM/dd') 結訓日期" & vbCrLf  '結訓日期
            'sql &= " ,a.STDate 開訓日期" & vbCrLf  '開訓日期
            'sql &= " ,a.FDDate 結訓日期" & vbCrLf  '結訓日期
            sql &= " ,a.Pkname12 課程分類" & vbCrLf  '課程分類

            'sql &= " ,a.D20KNAME1 ""5+2產業創新計畫""" & vbCrLf '5+2產業創新計畫
            'sql &= " ,a.D20KNAME2 ""台灣AI行動計畫""" & vbCrLf '台灣AI行動計畫
            'sql &= " ,a.D20KNAME3 ""數位國家創新經濟發展方案""" & vbCrLf '"數位國家創新經濟發展方案</td>" & vbcrlf
            'sql &= " ,a.D20KNAME4 ""國家資通安全發展方案""" & vbCrLf '"國家資通安全發展方案</td>" & vbcrlf
            'sql &= " ,a.D20KNAME5 ""前瞻基礎建設計畫""" & vbCrLf '"前瞻基礎建設計畫</td>" & vbcrlf
            'sql &= " ,a.D20KNAME6 ""新南向政策""" & vbCrLf '"新南向政策</td>" & vbcrlf
            If fg_Work2026x02 Then
                sql &= "
,a.D26KNAME1 ""五大信賴產業推動方案""
,a.D26KNAME2 ""六大區域產業及生活圈""
,a.D26KNAME3 ""智慧國家2.0綱領""
,a.D26KNAME4 ""新南向政策推動計畫""
,a.D26KNAME5 ""國家人才競爭力躍升方案""
,a.D26KNAME6 ""AI新十大建設推動方案""
,a.D26KNAME7 ""台灣AI行動計畫2.0""
,a.D26KNAME8 ""智慧機器人產業推動方案""
,a.D26KNAME9 ""臺灣2050淨零轉型""
"
            End If
            '亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
            sql &= " ,a.D25KNAME1 ""亞洲矽谷""" & vbCrLf
            sql &= " ,a.D25KNAME2 ""重點產業""" & vbCrLf
            sql &= " ,a.D25KNAME3 ""台灣AI行動計畫""" & vbCrLf
            sql &= " ,a.D25KNAME4 ""智慧國家方案""" & vbCrLf
            sql &= " ,a.D25KNAME5 ""國家人才競爭力躍升方案""" & vbCrLf
            sql &= " ,a.D25KNAME6 ""新南向政策""" & vbCrLf
            sql &= " ,a.D25KNAME7 ""AI加值應用""" & vbCrLf
            sql &= " ,a.D25KNAME8 ""職場續航""" & vbCrLf
            'dd.KID22,dd.KNAME22 進階政策性產業類別
            sql &= " ,a.KNAME22 ""進階政策性產業類別""" & vbCrLf '進階政策性產業類別

            sql &= " ,a.kname1 新興產業" & vbCrLf '※六大新興產業	
            sql &= " ,a.kname2 重點服務業" & vbCrLf '※十大重點服務業

            sql &= " ,a.kname3 新興智慧型產業" & vbCrLf  '※四大新興智慧型產業

            sql &= " ,a.GCodeName 訓練業別編碼" & vbCrLf   '訓練業別編碼
            sql &= " ,a.GCNAME 訓練業別" & vbCrLf  '訓練業別

            sql &= " ,CAST('' as NVARCHAR(100)) ""通俗職類-大類""" & vbCrLf
            sql &= " ,CAST('' as NVARCHAR(100)) ""通俗職類-小類""" & vbCrLf
            sql &= " ,a.CCName 訓練職能" & vbCrLf   '訓練職能

            sql &= " ,a.AddressSciPTID 學科辦訓地縣市" & vbCrLf  '學科辦訓地縣市
            sql &= " ,a.AddressTechPTID 術科辦訓地縣市" & vbCrLf '術科辦訓地縣市
            sql &= " ,a.PackageTypeN 包班種類" & vbCrLf  '包班總類 包班種類
        Else
            sql &= " ,a.YEARS 計畫年度" & vbCrLf
            sql &= " ,a.OrgPlanName2 計畫別" & vbCrLf  '自主/產投(產投計畫別)
            sql &= " ,a.DistName 各分署" & vbCrLf
            sql &= " ,a.OrgTypeName 單位屬性" & vbCrLf    '單位屬性
            sql &= " ,a.orgname 訓練機構" & vbCrLf  '訓練機構

            For a As Integer = 0 To Me.ChbExit.Items.Count - 1
                If ChbExit.Items.Item(a).Selected Then
                    Select Case ChbExit.Items.Item(a).Value
                        Case Cst_統一編號 'Case Cst_統一編號
                            sql &= " ,a.ComIDNO 統一編號" & vbCrLf
                        Case Cst_立案縣市 'Case Cst_立案縣市 oo.orgZipCode=iz2.ZipCode
                            sql &= " ,a.CTName2 立案縣市" & vbCrLf
                    End Select
                End If
            Next
            sql &= " ,a.ClassName 班別名稱" & vbCrLf  '班別名稱
            sql &= " ,a.FIRSTSORT 提案意願順序" & vbCrLf  '提案意願順序
            sql &= " ,a.CyclType 期別" & vbCrLf  '期別
            sql &= " ,a.ClassID 課程代碼" & vbCrLf  '課程代碼
            'APPSTAGE
            If gflag_SHOW_APPSTAGE Then sql &= " ,a.APPSTAGE 申請階段" & vbCrLf '申請階段
            sql &= " ,format(a.STDate,'yyyy/MM/dd') 開訓日期" & vbCrLf  '開訓日期
            sql &= " ,format(a.FDDate,'yyyy/MM/dd') 結訓日期" & vbCrLf  '結訓日期
            'sql &= " ,a.STDate 開訓日期" & vbCrLf  '開訓日期
            'sql &= " ,a.FDDate 結訓日期" & vbCrLf  '結訓日期
            sql &= " ,a.Pkname12 課程分類" & vbCrLf  '課程分類

            'If gflag_SHOW_2019_1 Then
            '    '政府政策性產業'「5+2」產業創新計畫 5+2產業'【台灣AI行動計畫】 KID='08''【數位國家創新經濟發展方案】KID='09''【國家資通安全發展方案】KID='10''【前瞻基礎建設計畫】'【新南向政策】KID='19' sql &= " ,a.D20KNAME1 ""5+2產業創新計畫""" & vbCrLf '5+2產業創新計畫
            '    sql &= " ,a.D20KNAME2 ""台灣AI行動計畫""" & vbCrLf '台灣AI行動計畫
            '    sql &= " ,a.D20KNAME3 ""數位國家創新經濟發展方案""" & vbCrLf '"數位國家創新經濟發展方案</td>" & vbcrlf
            '    sql &= " ,a.D20KNAME4 ""國家資通安全發展方案""" & vbCrLf '"國家資通安全發展方案</td>" & vbcrlf
            '    sql &= " ,a.D20KNAME5 ""前瞻基礎建設計畫""" & vbCrLf '"前瞻基礎建設計畫</td>" & vbcrlf
            '    sql &= " ,a.D20KNAME6 ""新南向政策""" & vbCrLf '"新南向政策</td>" & vbcrlf
            'End If
            '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
            '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
            If fg_Work2026x02 Then
                sql &= "
,a.D26KNAME1 ""五大信賴產業推動方案""
,a.D26KNAME2 ""六大區域產業及生活圈""
,a.D26KNAME3 ""智慧國家2.0綱領""
,a.D26KNAME4 ""新南向政策推動計畫""
,a.D26KNAME5 ""國家人才競爭力躍升方案""
,a.D26KNAME6 ""AI新十大建設推動方案""
,a.D26KNAME7 ""台灣AI行動計畫2.0""
,a.D26KNAME8 ""智慧機器人產業推動方案""
,a.D26KNAME9 ""臺灣2050淨零轉型""
"
            End If
            If gfg_SHOW_2025_1 Then
                '亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
                sql &= " ,a.D25KNAME1 ""亞洲矽谷""" & vbCrLf
                sql &= " ,a.D25KNAME2 ""重點產業""" & vbCrLf
                sql &= " ,a.D25KNAME3 ""台灣AI行動計畫""" & vbCrLf
                sql &= " ,a.D25KNAME4 ""智慧國家方案""" & vbCrLf
                sql &= " ,a.D25KNAME5 ""國家人才競爭力躍升方案""" & vbCrLf
                sql &= " ,a.D25KNAME6 ""新南向政策""" & vbCrLf
                sql &= " ,a.D25KNAME7 ""AI加值應用""" & vbCrLf
                sql &= " ,a.D25KNAME8 ""職場續航""" & vbCrLf
            End If
            'dd.KID22,dd.KNAME22 進階政策性產業類別
            sql &= " ,a.KNAME22 ""進階政策性產業類別""" & vbCrLf

            sql &= " ,a.kname1 新興產業" & vbCrLf '※六大新興產業	
            sql &= " ,a.kname2 重點服務業" & vbCrLf '※十大重點服務業
            sql &= " ,a.kname3 新興智慧型產業" & vbCrLf  '※四大新興智慧型產業

            sql &= " ,a.GCodeName 訓練業別編碼" & vbCrLf   '訓練業別編碼
            sql &= " ,a.GCNAME 訓練業別" & vbCrLf  '訓練業別

            sql &= " ,CAST('' as NVARCHAR(100)) ""通俗職類-大類""" & vbCrLf
            sql &= " ,CAST('' as NVARCHAR(100)) ""通俗職類-小類""" & vbCrLf

            sql &= " ,a.CCName 訓練職能" & vbCrLf   '訓練職能

            sql &= " ,a.AddressSciPTID 學科辦訓地縣市" & vbCrLf  '學科辦訓地縣市
            sql &= " ,a.AddressTechPTID 術科辦訓地縣市" & vbCrLf '術科辦訓地縣市
            'sql &= " ,a.PackageType" & vbCrLf  '包班種類
            sql &= " ,a.PackageTypeN 包班種類" & vbCrLf  '包班總類 包班種類

            For a As Integer = 0 To Me.ChbExit.Items.Count - 1
                If ChbExit.Items.Item(a).Selected Then
                    Select Case ChbExit.Items.Item(a).Value
                        Case Cst_政府政策性產業_114NOUSE 'Cst_政府政策性產業_108NOUSE '(108年之後不使用此欄)
                            'sql &= " ,a.KNAME19 ""政府政策性產業(108年之後不使用此欄)""" & vbCrLf  '政府政策性產業(108年之後不使用此欄)
                            sql &= " ,a.D20KNAME1 ""5+2產業創新計畫""" & vbCrLf '5+2產業創新計畫
                            sql &= " ,a.D20KNAME2 ""台灣AI行動計畫""" & vbCrLf '台灣AI行動計畫
                            sql &= " ,a.D20KNAME3 ""數位國家創新經濟發展方案""" & vbCrLf '"數位國家創新經濟發展方案</td>" & vbcrlf
                            sql &= " ,a.D20KNAME4 ""國家資通安全發展方案""" & vbCrLf '"國家資通安全發展方案</td>" & vbcrlf
                            sql &= " ,a.D20KNAME5 ""前瞻基礎建設計畫""" & vbCrLf '"前瞻基礎建設計畫</td>" & vbcrlf
                            sql &= " ,a.D20KNAME6 ""新南向政策""" & vbCrLf '"新南向政策</td>" & vbcrlf
                        Case Cst_新南向政策
                            sql &= " ,a.KNAME18 新南向政策" & vbCrLf  '新南向政策
                        Case Cst_轄區重點產業
                            sql &= " ,a.KNAME15 ""轄區重點產業""" & vbCrLf  '生產力4.0
                        Case Cst_生產力40
                            sql &= " ,a.KNAME14 ""生產力4.0""" & vbCrLf  '生產力4.0
                        Case Cst_申請人次
                            sql &= " ,a.ATNum  ""申請人數""" & vbCrLf  '申請人數
                        Case Cst_申請補助費
                            sql &= " ,a.ADefGovCost ""申請補助費""" & vbCrLf  '申請補助費
                        Case Cst_核定人次
                            sql &= " ,a.TNUM 核定人數" & vbCrLf  '核定人數
                        Case Cst_核定補助費
                            sql &= " ,a.DefGovCost 核定補助費" & vbCrLf '核定補助費
                        Case Cst_實際開訓人次
                            sql &= " ,a.openstudcount1 實際就保開訓人次" & vbCrLf
                            sql &= " ,a.openstudcount2 實際就安開訓人次" & vbCrLf
                            'sql &= " ,a.openstudcount3 實際公務開訓人次" & vbCrLf
                            sql &= String.Concat(" ,a.openstudcount97 實際", cst_BUDGET97_N, "開訓人次") & vbCrLf '公務ECFA // 協助

                        Case Cst_實際開訓人次加總
                            sql &= " ,a.openstudcountall 實際合計開訓人次" & vbCrLf  '開訓人次加總

                        Case Cst_預估補助費
                            sql &= " ,ROUND(a.cost1,2) 就保預估補助費金額" & vbCrLf
                            sql &= " ,ROUND(a.cost2,2) 就安預估補助費金額" & vbCrLf
                            'sql &= " ,ROUND(a.cost3,2) 公務預估補助費金額" & vbCrLf
                            sql &= String.Concat(" ,ROUND(a.cost97,2) ", cst_BUDGET97_N, "預估補助費金額") & vbCrLf '公務ECFA // 協助

                        Case Cst_預估補助費加總
                            sql &= " ,ROUND(a.costAll,2) 合計預估補助費金額" & vbCrLf

                        Case Cst_結訓人次 '合計結訓人次
                            sql &= " ,a.closestudcout03 就保結訓人次" & vbCrLf
                            sql &= " ,a.closestudcout02 就安結訓人次" & vbCrLf
                            'sql &= " ,a.closestudcout01 公務結訓人次" & vbCrLf
                            sql &= String.Concat(" ,a.closestudcout97 ", cst_BUDGET97_N, "結訓人次") & vbCrLf '公務ECFA // 協助
                            sql &= " ,a.closestudcoutall 合計結訓人次" & vbCrLf

                        Case Cst_撥款人次 'Key_Identity
                            sql &= " ,a.budcountall3 就保撥款人次" & vbCrLf
                            sql &= " ,a.budcountall2 就安撥款人次" & vbCrLf
                            'sql &= " ,a.budcountall3 公務撥款人次" & vbCrLf
                            sql &= String.Concat(" ,a.budcountall97 ", cst_BUDGET97_N, "撥款人次") & vbCrLf '公務ECFA // 協助

                        Case Cst_各身分別撥款人次
                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        sql &= String.Format(" ,a.bud03count{0} ""{1}""", dr1("IdentityID"), "就保一般身分撥款人次")
                                    Case Else
                                        sql &= String.Format(" ,a.bud03count{0} ""{1}""", dr1("IdentityID"), String.Format("就保特殊身分({0})撥款人次", dr1("NAME")))
                                End Select
                            Next
                            sql &= " ,a.Sbudcount01 就保特殊身分總撥款人次" & vbCrLf

                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        sql &= String.Format(" ,a.bud02count{0} ""{1}""", dr1("IdentityID"), "就安一般身分撥款人次")
                                    Case Else
                                        sql &= String.Format(" ,a.bud02count{0} ""{1}""", dr1("IdentityID"), String.Format("就安特殊身分({0})撥款人次", dr1("NAME")))
                                End Select
                            Next
                            sql &= " ,a.Sbudcount02 就安特殊身分總撥款人次" & vbCrLf

                            'For Each dr1 As DataRow In dtIdentity.Rows
                            '    Select Case Convert.ToString(dr1("IdentityID"))
                            '        Case "01"
                            '            sql &= String.Format(" ,a.bud01count{0} ""{1}""", dr1("IdentityID"), "公務一般身分撥款人次")
                            '        Case Else
                            '            sql &= String.Format(" ,a.bud01count{0} ""{1}""", dr1("IdentityID"), String.Format("公務特殊身分({0})撥款人次", dr1("NAME")))
                            '    End Select
                            'Next
                            'sql &= " ,a.Sbudcount03 公務特殊身分總撥款人次" & vbCrLf

                            '公務ECFA // 協助
                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        sql &= String.Format(" ,a.bud97count{0} ""{1}{2}""", dr1("IdentityID"), cst_BUDGET97_N, "一般身分撥款人次")
                                    Case Else
                                        sql &= String.Format(" ,a.bud97count{0} ""{1}{2}""", dr1("IdentityID"), cst_BUDGET97_N, String.Format("特殊身分({0})撥款人次", dr1("NAME")))
                                End Select
                            Next
                            sql &= String.Concat(" ,a.Sbudcount97 ", cst_BUDGET97_N, "特殊身分總撥款人次") & vbCrLf

                        Case Cst_撥款補助費
                            sql &= " ,a.budmoneyall3 就保撥款補助費" & vbCrLf
                            sql &= " ,a.budmoneyall2 就安撥款補助費" & vbCrLf
                            'sql &= " ,a.budmoneyall3 公務撥款補助費" & vbCrLf
                            sql &= String.Concat(" ,a.budmoneyall97 ", cst_BUDGET97_N, "撥款補助費") & vbCrLf '公務ECFA // 協助

                        Case Cst_各身分別撥款補助費
                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        sql &= String.Format(" ,a.bud03money{0} ""{1}""", dr1("IdentityID"), "就保一般身分撥款補助費")
                                    Case Else
                                        sql &= String.Format(" ,a.bud03money{0} ""{1}""", dr1("IdentityID"), String.Format("就保特殊身分({0})撥款補助費", dr1("NAME")))
                                End Select
                            Next
                            sql &= " ,a.Sbudmoney01 就保特殊身分總撥款補助費" & vbCrLf

                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        sql &= String.Format(" ,a.bud02money{0} ""{1}""", dr1("IdentityID"), "就安一般身分撥款補助費")
                                    Case Else
                                        sql &= String.Format(" ,a.bud02money{0} ""{1}""", dr1("IdentityID"), String.Format("就安特殊身分({0})撥款補助費", dr1("NAME")))
                                End Select
                                'sql &= " ,a.bud02money" & dr1("IdentityID" & vbCrLf
                            Next
                            sql &= " ,a.Sbudmoney02 就安特殊身分總撥款補助費" & vbCrLf

                            'For Each dr1 As DataRow In dtIdentity.Rows
                            '    Select Case Convert.ToString(dr1("IdentityID"))
                            '        Case "01"
                            '            sql &= String.Format(" ,a.bud01money{0} ""{1}""", dr1("IdentityID"), "公務一般身分撥款補助費")
                            '        Case Else
                            '            sql &= String.Format(" ,a.bud01money{0} ""{1}""", dr1("IdentityID"), String.Format("公務特殊身分({0})撥款補助費", dr1("NAME")))
                            '    End Select
                            'Next
                            'sql &= " ,a.Sbudmoney03 公務特殊身分總撥款補助費" & vbCrLf

                            '公務ECFA // 協助
                            For Each dr1 As DataRow In dtIdentity.Rows
                                Select Case Convert.ToString(dr1("IdentityID"))
                                    Case "01"
                                        sql &= String.Format(" ,a.bud97money{0} ""{1}{2}""", dr1("IdentityID"), cst_BUDGET97_N, "一般身分撥款補助費")
                                    Case Else
                                        sql &= String.Format(" ,a.bud97money{0} ""{1}{2}""", dr1("IdentityID"), cst_BUDGET97_N, String.Format("特殊身分({0})撥款補助費", dr1("NAME")))
                                End Select
                                'sql &= " ,a.bud97money" & dr1("IdentityID" & vbCrLf
                            Next
                            sql &= String.Concat(" ,a.Sbudmoney97 ", cst_BUDGET97_N, "特殊身分總撥款補助費") & vbCrLf

                        Case Cst_不預告訪視次數_實地抽訪
                            sql &= " ,a.VW2CNT ""累計不預告視訊抽訪次數""" & vbCrLf
                            sql &= " ,REPLACE(a.VW2APPLYDATE,',',';') ""視訊訪視日期""" & vbCrLf

                        Case Cst_不預告訪視次數_電話抽訪
                            sql &= " ,a.CTALL3 累計不預告電話抽訪次數" & vbCrLf    '累計不預告實地抽訪異常次數
                            sql &= " ,a.cuAPPLYDATE3 電話訪視日期" & vbCrLf    '累計不預告視訊抽訪異常次數
                            sql &= " ,a.CTALL4 ""累計不預告電話抽訪次數(實地抽訪未到)""" & vbCrLf    '累計不預告電話抽訪異常次數
                            sql &= " ,a.cuAPPLYDATE4 ""電話訪視日期(實地抽訪未到)""" & vbCrLf    '累計不預告電話抽訪異常次數

                        Case Cst_不預告訪視次數_視訊訪查
                            sql &= " ,a.VW2CNT ""累計不預告視訊抽訪次數""" & vbCrLf
                            sql &= " ,REPLACE(a.VW2APPLYDATE,',',';') ""視訊訪視日期""" & vbCrLf

                        Case Cst_累積訪視異常次數
                            sql &= " ,a.vitN 累計不預告實地抽訪異常次數" & vbCrLf    '累計不預告實地抽訪異常次數
                            sql &= " ,a.vitTVN 累計不預告視訊抽訪異常次數" & vbCrLf    '累計不預告視訊抽訪異常次數
                            sql &= " ,a.VitTelN 累計不預告電話抽訪異常次數" & vbCrLf    '累計不預告電話抽訪異常次數

                        Case Cst_累計訪視異常原因
                            sql &= " ,a.It22b01N ""累計訪視異常原因-出席率不佳""" & vbCrLf      '累計訪視異常原因-出席率不佳
                            sql &= " ,a.It22b02N ""累計訪視異常原因-簽到退未落實""" & vbCrLf    '累計訪視異常原因-簽到退未落實
                            sql &= " ,a.It22b03N ""累計訪視異常原因-師資不符""" & vbCrLf        '累計訪視異常原因-師資不符
                            sql &= " ,a.It22b06N ""累計訪視異常原因-助教不符""" & vbCrLf        '累計訪視異常原因-助教不符
                            sql &= " ,a.It22b04N ""累計訪視異常原因-課程內容不符""" & vbCrLf    '累計訪視異常原因-課程內容不符
                            sql &= " ,a.It22b05N ""累計訪視異常原因-上課地點不符""" & vbCrLf    '累計訪視異常原因-上課地點不符
                            sql &= " ,a.IT22B99NOTE ""累計訪視異常原因-其他""" & vbCrLf         '累計訪視異常原因-其他
                            sql &= " ,a.LITEM23NOTE ""累計訪視異常原因-其他補充說明""" & vbCrLf '累計訪視異常原因-其他補充說明

                            'Case Cst_會計查帳次數 dtW.Columns.Add(New DataColumn("會計查帳次數"))   
                        Case Cst_離訓人次
                            sql &= " ,a.std_cnt2 離訓人次" & vbCrLf   '離訓人次
                        Case Cst_退訓人次
                            sql &= " ,a.std_cnt3 退訓人次" & vbCrLf   '退訓人次
                        Case Cst_訓練時數
                            sql &= " ,a.THOURS 訓練時數" & vbCrLf   '訓練時數
                        Case Cst_固定費用總額
                            sql &= " ,a.FIXSUMCOST 固定費用總額" & vbCrLf
                        Case Cst_固定費用單一人時成本
                            sql &= " ,a.ACTHUMCOST 固定費用單一人時成本" & vbCrLf
                        Case Cst_人時成本超出原因說明
                            sql &= " ,a.FIXExceeDesc 人時成本超出原因說明" & vbCrLf
                        Case Cst_材料費總額
                            sql &= " ,a.METSUMCOST 材料費總額" & vbCrLf
                        Case Cst_材料費占比
                            sql &= " ,a.METCOSTPER 材料費占比" & vbCrLf
                        Case Cst_超出材料費比率上限原因說明
                            sql &= " ,a.METExceeDesc 超出材料費比率上限原因說明" & vbCrLf
                            'Case Cst_人時成本
                            'dtW.Columns.Add(New DataColumn("人時成本"))   
                        Case Cst_上課時間
                            sql &= " ,a.WEEKSTIME 上課時間" & vbCrLf   '上課時間
                        Case Cst_撥款日期
                            sql &= " ,a.AllotDate 撥款日期" & vbCrLf
                        Case Cst_包班事業單位
                            sql &= " ,a.BusPackage 包班事業單位" & vbCrLf
                        Case Cst_師資名單
                            sql &= " ,a.PlanTeacher 師資名單" & vbCrLf
                        Case Cst_上課地址及教室
                            sql &= " ,CAST('' as NVARCHAR(1000)) ""上課地址及教室""" & vbCrLf
                            flag_use_上課地址及教室 = True
                            'Case Cst_上課地址及教室2
                            '    dtW.Columns.Add(New DataColumn("上課地址及教室2"))   
                        Case Cst_包班事業單位保險證號
                            sql &= " ,a.BusPackage2 包班事業單位保險證號" & vbCrLf
                        Case Cst_包班事業單位統一編號
                            sql &= " ,a.BusPackage3 包班事業單位統一編號" & vbCrLf
                        Case Cst_公務ECFA性別人數 '公務ECFA // 協助
                            sql &= String.Concat(" ,a.SexNumxM ", cst_BUDGET97_N, "男性人數") & vbCrLf  '/*男性人數*/ 協助男性人數 '公務ECFA // 協助
                            sql &= String.Concat(" ,a.SexNumxF ", cst_BUDGET97_N, "女性人數") & vbCrLf  '/*女性人數*/ 協助女性人數 '公務ECFA // 協助

                        Case cst_課程申請流水號
                            sql &= " ,a.PSNO28 課程申請流水號" & vbCrLf
                        Case cst_上架日期
                            sql &= " ,a.ONSHELLDATE 上架日期" & vbCrLf
                        Case cst_開放報名結束日期
                            sql &= " ,a.SENTERDATE 開放報名日期" & vbCrLf
                            sql &= " ,a.FENTERDATE 結束報名日期" & vbCrLf
                        Case cst_課程備註
                            sql &= " ,a.memo8 ""課程備註1""" & vbCrLf
                            sql &= " ,a.memo82 ""課程備註2""" & vbCrLf
                        Case cst_術科時數
                            sql &= " ,a.ProTechHours 術科時數" & vbCrLf
                        Case cst_聯絡人
                            sql &= " ,a.ContactName 聯絡人" & vbCrLf
                        Case cst_聯絡電話
                            sql &= " ,a.ContactPhone 聯絡電話" & vbCrLf
                        Case cst_是否停辦
                            sql &= " ,a.NotOpenN 是否停辦" & vbCrLf
                        Case cst_iCAP標章證號及效期
                            sql &= " ,a.ICAPNUM ""iCAP標章證號""" & vbCrLf
                            sql &= " ,a.iCAPMARKDATE ""iCAP效期""" & vbCrLf
                            'Case cst_政策性產業課程可辦理班數
                            '    dtW.Columns.Add(New DataColumn(T_YR1 ))  
                            '    dtW.Columns.Add(New DataColumn(T_YR2 ))  
                            '    dtW.Columns.Add(New DataColumn(T_YR3 ))  
                        Case Cst_辦理方式
                            '辦理方式-遠距教學 DISTANCE,dbo.FN_GET_DISTANCE(DISTANCE) DISTANCE_N
                            sql &= " ,dbo.FN_GET_DISTANCE(a.DISTANCE) ""辦理方式""" & vbCrLf
                        Case Cst_實際開訓性別人數
                            sql &= " ,a.SexCNTM ""實際開訓男性人數""" & vbCrLf
                            sql &= " ,a.SexCNTF ""實際開訓女性人數""" & vbCrLf
                        Case Cst_行政管理疏失重大異常狀況 '53
                            sql &= " ,a.VERIFYDATE ""行政管理疏失/重大異常狀況""" & vbCrLf '為分署於該功能登打之經查核確認日期。
                        Case Cst_報名繳費方式 '54 'ENTERSUPPLYSTYLE
                            sql &= " ,a.ENTERSUPPLYSTYLE_N ""報名繳費方式""" & vbCrLf

                    End Select
                End If
            Next
        End If
        Return sql
    End Function

    '組合 Cst_上課地址及教室
    Public Shared Function Get_AddressPlaceName(ByRef dtZip As DataTable, ByVal dr As DataRow) As String
        Dim rst As String = ""
        Dim tmpAddr As String = ""
        Const cst_spTag As String = "、"
        For i As Integer = 1 To 4
            Dim strTag As String = If(i = 1, "s1", If(i = 2, "s2", If(i = 3, "t1", If(i = 4, "t2", ""))))
            Dim sZipCode As String = If(Convert.ToString(dr(strTag & "ZIP6W")) <> "", Convert.ToString(dr(strTag & "ZIP6W")), Convert.ToString(dr(strTag & "zipCode")))
            tmpAddr = TIMS.getZipName6(sZipCode, Convert.ToString(dr(strTag & "ADDRESS")), Convert.ToString(dr(strTag & "PLACENAME")), dtZip)
            If tmpAddr <> "" Then rst &= String.Concat(If(rst <> "", cst_spTag, ""), tmpAddr)
        Next
        Return rst
    End Function

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    ''' <summary> 匯出 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnExp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExp.Click
        Dim okFlag As Boolean = False 'okFlag = False '結束狀況有誤
        flagExp1 = False '匯出全域變數判斷

        Dim Errmsg As String = ""
        SDate1.Text = TIMS.ClearSQM(SDate1.Text)
        SDate2.Text = TIMS.ClearSQM(SDate2.Text)
        EDate1.Text = TIMS.ClearSQM(EDate1.Text)
        EDate2.Text = TIMS.ClearSQM(EDate2.Text)

        Dim v_Syear As String = TIMS.GetListValue(yearlist)
        Dim iMOK As Integer = 0
        If v_Syear <> "" Then iMOK += 2
        If SDate1.Text <> "" Then iMOK += 1
        If SDate2.Text <> "" Then iMOK += 1
        If EDate1.Text <> "" Then iMOK += 1
        If EDate2.Text <> "" Then iMOK += 1

        Dim fg_NG_ALL_NO_SEL1 As Boolean = (v_Syear = "" AndAlso SDate1.Text <> "" AndAlso SDate2.Text <> "" AndAlso EDate1.Text <> "" AndAlso EDate2.Text <> "")
        If fg_NG_ALL_NO_SEL1 OrElse iMOK < 2 Then Errmsg &= "請選擇年度或開結訓日期(至少要有2項填寫)" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return
        End If

        Try
            'Dim da As New SqlDataAdapter
            'da.SelectCommand = New SqlCommand
            'da.SelectCommand.Connection = objConn
            'da.SelectCommand.CommandTimeout = 100
            Call ExpRpt() '匯出SUB'SQL
            okFlag = True '結束狀況無誤
            'Call TIMS.CloseDbConn(objConn)
        Catch ex As System.Threading.ThreadAbortException
            TIMS.LOG.Warn(ex.Message, ex)
            Server.ClearError()
        Catch ex As Exception
            Dim sErrMsg1 As String = String.Concat("發生錯誤:", vbCrLf, ex.ToString, vbCrLf, "g_ErrSql : ", vbCrLf, g_ErrSql)
            Call TIMS.WriteTraceLog(Page, ex, sErrMsg1)

            'If conn.State = ConnectionState.Open Then conn.Close()
            Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.Message)
            Call TIMS.CloseDbConn(objConn)
            If Response IsNot Nothing AndAlso (Response.IsClientConnected) Then Response.End()
            Return
        End Try

        Call TIMS.CloseDbConn(objConn)
        'str_TIME_MSG_A = "此查詢動作，總共花費" & DateDiff(DateInterval.Second, DateSec1, DateSec2) & "秒，共" & CStr(dt.Rows.Count) & "筆" & vbCrLf
        '結束狀況無誤
        If okFlag AndAlso flagExp1 Then
            If Response IsNot Nothing AndAlso (Response.IsClientConnected) Then Response.End() 'Response.End()
        End If
    End Sub

    '年度選擇。
    Private Sub yearlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yearlist.SelectedIndexChanged
        If yearlist.SelectedValue = "" Then Exit Sub
        Call setSelYears1(yearlist.SelectedValue)
    End Sub

End Class
