Partial Class TC_03_001
    Inherits AuthBasePage

    'PLAN_TRAINDESC'PLAN_COSTITEM'PLAN_DEPOT'PLAN_ABILITY'PLAN_PLANINFO'CLASS_CLASSINFO'AUTH_RELSHIP'ORG_ORGPLANINFO

    '(每次loading執行)
    Sub SUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("PLAN_PLANINFO,PLAN_TRAINDESC,PLAN_COSTITEM", objconn)
        If TIMS.dtNODATA(dt) Then Return 'Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "PLANEMAIL", EMail) 'EMAIL
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER1", Other1) '其他一
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER2", Other2) '其他二
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER3", Other3) '其他三
        Call TIMS.sUtl_SetMaxLen(dt, "PNAME", PName) '單元名稱.PLAN_TRAINDESC
        Call TIMS.sUtl_SetMaxLen(dt, "PCONT", PCont) '課程大綱.PLAN_TRAINDESC
        Call TIMS.sUtl_SetMaxLen(dt, "TMSCIENCE", TMScience) '學科
        Call TIMS.sUtl_SetMaxLen(dt, "TMTECH", TMTech) '術科
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSNAME", ClassName) '班別名稱
        Call TIMS.sUtl_SetMaxLen(dt, "CYCLTYPE", CyclType) '期別
        Call TIMS.sUtl_SetMaxLen(dt, "TAddress", TAddress) '上課地址
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTNAME", ContactName) '聯絡人
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTPHONE", ContactPhone) '聯絡人電話
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTEMAIL", ContactEmail) '聯絡人電子郵件
        Call TIMS.sUtl_SetMaxLen(dt, "MASTEREMAIL", MasterEmail) '直屬主管電子郵件
        Call TIMS.sUtl_SetMaxLen(dt, "ITEMOTHER", ItemOther) '項目中的其他費用
        Call TIMS.sUtl_SetMaxLen(dt, "ITEMOTHER", ItemOther4) '項目中的其他費用
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSENGNAME", ClassEngName) '班級英文名稱'
        Call TIMS.sUtl_SetMaxLen(dt, "CTNAME", CTName) '導師名稱

        '就服單位協助報名
        'txtEpNum.Enabled = False '該欄位鎖定 by AMU 2013
        'Trwork2013a.Visible = False
        'Trwork2013b.Visible = False
        'Trwork2013c
        'UPDATE VIEW : SELECT * FROM VIEW_JOBCLASSINFO
        'If sm.UserInfo.Years >= 2013 _
        '    AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If TIMS.Utl_GetConfigSet("work2013") = "Y" Then
        '        Trwork2013a.Visible = True
        '        Trwork2013b.Visible = True
        '    End If
        'End If

        'Dim dt As DataTable
        box9.Visible = True
        box10.Visible = True
        Title1.Text = "訓練目標"
        Title2.Text = "就業展望"
        fill2.ErrorMessage = "目標「訓練目標」為必填欄位"
        fill3.ErrorMessage = "目標「就業展望」為必填欄位"
        fill4.Enabled = False
        fill13.Enabled = False
        fill14.Enabled = False
        fill23.Enabled = True
        fill24.Enabled = True
        fill25.Enabled = True
        fill26.Enabled = True
        fill27.Enabled = True
        fill31.Enabled = True
        fill32.Enabled = True
        TR_2005_01.Visible = False
        TR_2006_01.Visible = True
        If sm.UserInfo.Years <= 2005 Then
            '舊年度資料
            box9.Visible = False
            box10.Visible = False
            Title1.Text = "學科"
            Title2.Text = "技能"
            fill2.ErrorMessage = "目標「學科」為必填欄位"
            fill3.ErrorMessage = "目標「技能」為必填欄位"
            fill4.Enabled = True
            fill13.Enabled = True
            fill14.Enabled = True
            fill23.Enabled = False
            fill24.Enabled = False
            fill25.Enabled = False
            fill26.Enabled = False
            fill27.Enabled = False
            fill31.Enabled = False
            fill32.Enabled = False
            TR_2005_01.Visible = True
            TR_2006_01.Visible = False
        End If

        'OJT-21061501：<系統> 自辦在職、區域 - 班級申請：隱藏【企業負擔金額】欄位 欄位為接受企業委託計畫才會使用
        tr_DefUnitCost.Style("display") = If(TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1, "", "none")

        Dim flag_SHOW_2019_2 As Boolean = TIMS.SHOW_2019_2(sm)
        Dim flag_SHOW_2020x70 As Boolean = TIMS.SHOW_2020x70(sm)
        '2019年啟用 work2019x01:2019 政府政策性產業
        '為配合政府政策性產業推動，將分署自辦在職訓練、接受企業委託訓練2計畫增加【政府政策性產業】功能(項目同產投 '70:區域產業據點職業訓練計畫(在職)
        trKID20.Visible = False
        If flag_SHOW_2019_2 Then trKID20.Visible = True
        If flag_SHOW_2020x70 Then trKID20.Visible = True

        Dim fg_SHOW_2026_1 As Boolean = TIMS.SHOW_2026_1(sm)
        trKID20.Visible = If(fg_SHOW_2026_1, False, True)
        trKID25.Visible = If(fg_SHOW_2026_1, True, False)

        '如登入計畫別為 區域產業據點職業訓練計畫之變更 待計畫別碼
        'If sm.UserInfo.TPlanID = "" Then
        '    Label9.Enabled = True 
        '    Label9.Visible = True
        '    Label1.Enabled = True
        '    Label1.Visible = True
        '    Label8.Enabled = True 
        '    Label8.Visible = True
        '    GetTrain3.Items(0).Enabled = True
        '    GetTrain3.Items(3).Enabled = True
        '    GetTrain3.Items(4).Enabled = True
        '    GetTrain3.Items(5).Enabled = True
        '    intertype.Attributes.Add("class", "bluecol")
        '    others.Attributes.Add("class", "bluecol")
        '    interdate.Attributes.Add("class", "bluecol")
        'End If
        '2009年e網暫不需要

        '70:區域產業據點職業訓練計畫(在職)
        flag_TPlanID70_1 = (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1)
        '70:區域產業據點職業訓練計畫(在職)
        Hid_TPlanID.Value = sm.UserInfo.TPlanID
        tr_ACTHUMCOST.Visible = False
        tr_METCOSTPER.Visible = False
        'Dim v_GetTrain3 As String = TIMS.GetListValue(GetTrain3)
        If flag_TPlanID70_1 Then
            fill23.Enabled = False '不啟用
            tr_GetTrain1.Visible = False
            fill25.Enabled = False '不啟用
            fill26.Enabled = False '不啟用
            tr_ACTHUMCOST.Visible = True
            tr_METCOSTPER.Visible = True
            TIMS.SET_CLASS_BLUECOL_1(td_GetTrain3, lab_msg_GetTrain3)
            TIMS.SET_CLASS_BLUECOL_1(td_GetTrain4, lab_msg_GetTrain4)
            TIMS.SET_CLASS_BLUECOL_1(td_ExamDate, lab_msg_ExamDate)
        End If
        'If v_GetTrain3 <> "" Then TIMS.SetCblValue(GetTrain3, v_GetTrain3)

        TR_2006_01.Visible = False
        '47 補助辦理照顧服務員職業訓練'68 照顧服務員自訓自用訓練計畫
        HyperLink1.Visible = True
        HyperLink2.Visible = False
        If TIMS.Cst_TPlanID47AppPlan7.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            HyperLink1.Visible = False
            HyperLink2.Visible = True
        End If

        TIMS.Tooltip(THours, " 與課程編配的總時數一致", True)
        THours.Attributes.Add("readonly", "readonly") ' 與課程編配的總時數一致 ,不得修改

        '產生郵遞區號JavaScript
        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))
        '取出鍵詞-放入DataTable
        Dim sql As String = " SELECT * FROM KEY_COSTITEM ORDER BY SORT"
        dtKEYCOSTITEM = DbAccess.GetDataTable(sql, objconn)

        '段2 'andy ' btu_sel.Attributes("onclick") =GetPostBackEventReference(trainValue)
        ' tb_title.Attributes("onclick") =GetPostBackEventReference(trainValue)

        '登入計畫是學習券，改變課程名稱
        ClassName.ReadOnly = False
        Button26.Visible = False
        'If sm.UserInfo.TPlanID = "15" Then
        '    'ClassName.ReadOnly = True'禁修改 2012年啟用開放修改 BY AMU
        '    Button26.Visible = True '產生班級名稱
        '    Button26.Attributes("onclick") = "wopen('../01/TC_01_004_unit.aspx?textField=ClassName&valueField=Class_Unit&classunit='+document.form1.Class_Unit.value,'班級名稱',450,400,1);"
        '    rdoAge1.Checked = True
        '    rdoAge2.Checked = False
        'End If
        '登入計畫為地方政府補助則把行政事務費的警告視窗取消
        'If sm.UserInfo.TPlanID = "17" Then hidden17.Value = sm.UserInfo.TPlanID

        gflag_ccopy = False
        If Convert.ToString(Request(cst_ccopy)) = "1" Then gflag_ccopy = True

        'Dim rqPlanID As String = Request("PlanID")
        'Dim rqComIDNO As String = Request("ComIDNO")
        'Dim rqSeqNO As String = Request("SeqNO")
        rqPlanID = TIMS.ClearSQM(Request("PlanID"))
        rqComIDNO = TIMS.ClearSQM(Request("ComIDNO"))
        rqSeqNO = TIMS.ClearSQM(Request("SeqNO"))
        '使用  Request("PlanID") 若是false:使用 sm.UserInfo.PlanID
        '依 Request: PlanID/ComIDNO/SeqNO 查出是否有該班1筆資料。true:有 false:異常
        Dim blnuseRePlanID As Boolean = If(rqPlanID <> "" AndAlso rqComIDNO <> "" AndAlso rqSeqNO <> "", TIMS.ChkPPInfo(rqPlanID, rqComIDNO, rqSeqNO, objconn), False)
        '判斷計畫種類，選擇要顯示的經費項目
        sql = " SELECT TPLANID, PLANKIND, YEARS FROM ID_PLAN WHERE PLANID=@PlanID"
        'Call TIMS.OpenDbConn(objconn)
        Dim dtIP As New DataTable
        Using sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Clear()
                '似乎有登入異常，或查詢異常情況(有可能是要新增資料)
                .Parameters.Add("PlanID", SqlDbType.VarChar).Value = If(blnuseRePlanID, rqPlanID, sm.UserInfo.PlanID)
                dtIP.Load(.ExecuteReader())
            End With
        End Using
        If TIMS.dtNODATA(dtIP) Then
            '查無計畫 'vMsg = "程式出現例外狀況(查無計畫)，請聯絡TIMS系統駐點人員!"
            vMsg = "程式出現例外狀況(查無計畫)，請重新操作查詢!!"
            Common.MessageBox(Me, vMsg)
            Call TIMS.Utl_RespWriteEnd(Me, objconn, vMsg)
            Return 'Exit Sub
        End If

        Dim dr1 As DataRow = dtIP.Rows(0)
        iPlanKind = dr1("PlanKind") '判斷計畫種類
        sTPlanID = dr1("TPlanID")
        If iPlanKind = 1 Then
            Table1_Email.Visible = False
            Page.RegisterStartupScript("0900", "<script>ShowOther('CostID','ItemOther');</script>")
        Else
            Table1_Email.Visible = True
            Page.RegisterStartupScript("0900", "<script>ShowOther('CostID4','ItemOther4');</script>")
        End If

        Button8.Enabled = True
        btnAdd.Enabled = True

        '找出訓練計畫，並選擇計價種類
        If iPlanKind = 1 Then
            TableCost1.Style("display") = ""
        Else
            'PlanKind (2) 1:自辦 2:委外
            i_CostMode = GetCostMode(1)
            If i_CostMode = 0 Then
                vMsg = "程式出現例外狀況(查無計價種類)，請先設定-計價種類!!"
                Common.MessageBox(Me, vMsg)
                Call TIMS.Utl_RespWriteEnd(Me, objconn, vMsg)
                Return 'Exit Sub
            End If
        End If

        '段3 by 2020 AMU
        SHOW_TRNUNIT1()

        'OJT-23041104：區域據點-開班資料查詢：【班級英文名稱】改為非必填 
        'OJT-23041103：區域據點-班級申請：【班級英文名稱】改為非必填
        td_ClassEngName.Attributes.Add("class", If(TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) = -1, "bluecol_need", "bluecol"))
    End Sub

    '--UPDATE 班級申請作業'PLAN_PLANINFO '計畫主檔'Plan_TrainDesc '課程資料'Plan_CostItem '經費資料
    'Dim gda As SqlDataAdapter = Nothing
    'Dim giAgeType As Integer = 0 '1:一般計畫 2:計畫代碼：47及68 3:計畫代碼：58
    'HidAgeType.Value
    'PLAN_DEPOT / V_PLAN_DEPOT
    Const cst_AgeGt1 As String = "1" '1:一般計畫 2:計畫代碼：47及68 3:計畫代碼：58
    Const cst_AgeGt2 As String = "2" '1:一般計畫 2:計畫代碼：47及68 3:計畫代碼：58
    Const cst_AgeGt3 As String = "3" '1:一般計畫 2:計畫代碼：47及68 3:計畫代碼：58
    Const cst_tplanid47age2 As String = "47,68"
    Const cst_tplanid58age3 As String = "58"
    Const cst_tplanid70 As String = "70"

    'SAVE_PLAN_PLANINFO
    Dim PlanID_value As String = ""
    Dim ComIDNO_value As String = ""
    Dim SeqNO_value As String = ""
    'select * from key_plan where tplanid in (47,68)
    'Const cst_AgeStr1 As String = "年滿15歲以上。"
    '補助辦理照顧服務員職業訓練及照顧服務員自訓自用訓練計畫(計畫代碼：47及68)
    Const cst_AgeStr2 As String = "年滿16歲以上"
    '補助辦理托育人員職業訓練(計畫代碼：58)
    Const cst_AgeStr3 As String = "年滿20歲以上"

    Dim gflag_ccopy As Boolean = False 'Request(cst_ccopy) true:copy /false: not copy 
    Const cst_ccopy As String = "ccopy" 'Request(cst_ccopy)

    Dim iPlanKind As Integer = 0
    Dim sTPlanID As String = ""
    Dim dtKEYCOSTITEM As DataTable = Nothing
    Dim i_CostMode As Integer = 0
    Dim rqPlanID As String = "" 'Request("PlanID")
    Dim rqComIDNO As String = "" 'Request("ComIDNO")
    Dim rqSeqNO As String = "" 'Request("SeqNO")
    Dim flag_TPlanID70_1 As Boolean = False ' '70:區域產業據點職業訓練計畫(在職)
    'Dim TFstr As String = "false" '用來控制甄選方式(GetTrain3)checkbox勾選項目 /自辦:只有口試跟筆試能選
    '產業別(管考) true:使用/false:不可使用
    Dim fg_USE_CBLKID60_TP06 As Boolean = False

    'Dim blnCanAdds As Boolean = False '新增'Dim blnCanMod As Boolean = False '修改'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢'Dim blnCanPrnt As Boolean = False '列印

    Const cst_searchTC02001 As String = "searchTC02001"
    'Const cst_PlanTrainDesc As String = "Plan_TrainDesc"
    'Const cst_CostItemTable As String = "CostItemTable"
    Const cst_PlanTrainDescPKName As String = "PTDID"

    'Dim blnCanUpdataAuth As Boolean = False
    Dim blnCanUpdataTestUser As Boolean = False '非測試使用環境(正式)
    Dim vMsg As String = ""

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        'blnCanUpdataAuth = TIMS.CheckAuthUse(sm.UserInfo.UserID, objconn, 1)
        Call SUtl_PageInit1()
        '每次執行
        Call CCreate1_Every()

        '甄試日期/時間 'If TIMS.GFG_OJT_25050801_NoUse_ExamDateTime Then spExamDateTime.Visible = False:HR6.Visible = False : MM6.Visible = False

        '放置畫面上的Dropdonwlist
        If Not IsPostBack Then
            '產生新的GUID 避免記憶體相同 而異常
            Call CREATE_NEW_GUID21()
            Session("AdmGrant") = Nothing '行政管理費百分比
            Session("TaxGrant") = Nothing '營業稅費用百分比
            'If Not Session(cst_searchTC02001) Is Nothing Then
            '   ViewState(cst_searchTC02001) = Session(cst_searchTC02001)
            '    Session(cst_searchTC02001) = Nothing
            'End If

            '建立下拉選單物件
            Call cCreateItem()

            AdmGrantTR.Visible = False
            TaxGrantTR.Visible = False
            TableCost1.Style("display") = "none"
            TableCost2.Style("display") = "none"
            TableCost3.Style("display") = "none"
            TableCost4.Style("display") = "none"
            DataGrid1Table.Style.Item("display") = "none"
            DataGrid2Table.Style.Item("display") = "none"
            DataGrid3Table.Style.Item("display") = "none"
            DataGrid4Table.Style.Item("display") = "none"

            '20080811 andy 暫時先不擋
            'If sm.UserInfo.LID = "2" Then CheckPlain15()
            If (LayerState.Value = "") Then LayerState.Value = "1"
            Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("window_onload", s_js1)
            Call PageLoad_cCreate1() '導入網頁時 驗證基本訊息。
        End If

        '2004/12/7- -前端增加javascript屬性- -Start
        'Me.rblAge.Attributes("onclick") = "set_Agelu();"
        date1.Attributes("onclick") = "javascript:show_calendar('STDate','','','CY/MM/DD');"
        date2.Attributes("onclick") = "javascript:show_calendar('FDDate','','','CY/MM/DD');"


        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
            Org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            Org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx?btnName=Button28');"
        End If

        '增加快速點選機構清單
        If Org.Disabled = False Then
            TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "Button28")
            If HistoryRID.Rows.Count <> 0 Then
                center.Attributes("onclick") = "showObj('HistoryList2');"
                center.Style("CURSOR") = "hand"
            End If
        End If

        ''選擇全部
        'Me.CapMilitary.Attributes("onclick") = "SelectAllcbkList('CapMilitary','CapMilitaryHidden');"

        '//計算學科時數
        GenSciHours.Attributes("onblur") = "set_SciHours();"
        ProSciHours.Attributes("onblur") = "set_SciHours();"
        ProTechHours.Attributes("onblur") = "set_SciHours();"
        OtherHours.Attributes("onblur") = "set_SciHours();"

        '//班別資料訓練時數
        THours.Attributes("onblur") = "set_THours();"

        '訓練課程內容簡介
        Button29.Attributes("onclick") = "return CheckDescData('PName','PHour','PCont');"

        '自辦申請計畫
        Button2.Attributes("onclick") = "return check_Cost1();"

        'Hid_CostItem_GUID1 /CIGD TC_03_Adm/TC_03_Adm2
        Dim s_b3wo As String = ""
        s_b3wo = String.Format("{0}?CIGD={1}", "TC_03_Adm.aspx", Hid_CostItem_GUID1.Value)
        Button3.Attributes("onclick") = "if(checkAdm(1)){wopen('" & s_b3wo & "','Adm',400,300,0);document.form1.AdmGrant.value='1';}"
        s_b3wo = String.Format("{0}?CIGD={1}", "TC_03_Adm2.aspx", Hid_CostItem_GUID1.Value)
        Button3b.Attributes("onclick") = "if(checkAdm(3)){wopen('" & s_b3wo & "','Tax',400,300,0);document.form1.TaxGrant.value='1';}"

        Button9.Attributes("onclick") = "return check_Cost2();"
        Button8.Attributes("onclick") = "return Check_Temp();"
        Button10.Attributes("onclick") = "return check_Cost3();"
        Button11.Attributes("onclick") = "return check_Cost4();"

        'Hid_CostItem_GUID1 /CIGD
        s_b3wo = String.Format("{0}?CIGD={1}", "TC_03_Adm.aspx", Hid_CostItem_GUID1.Value)
        Button25.Attributes("onclick") = "if(checkAdm(2)){wopen('" & s_b3wo & "','Adm',400,300,0);document.form1.AdmGrant4.value='1';}"
        s_b3wo = String.Format("{0}?CIGD={1}", "TC_03_Adm2.aspx", Hid_CostItem_GUID1.Value)
        Button25b.Attributes("onclick") = "if(checkAdm(4)){wopen('" & s_b3wo & "','Tax',400,300,0);document.form1.TaxGrant4.value='1';}"
        'TAddressZip.Attributes("onblur") = "getZipName('CTName','TAddressZip',this.value);"
        'getzipname(zip,CityTextBox,ZipTextBox)

        If (1 = 1) Then
            TAddressZip.Attributes("onblur") = "getzipname(this.value, 'CCTName','TAddressZip');"

            Dim BtnTAddressZip_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(TAddressZip, TAddressZIPB3, hidTAddressZIP6W, CCTName, TAddress)
            BtnTAddressZip.Attributes.Add("onclick", BtnTAddressZip_Attr_VAL)
        End If
        If (1 = 1) Then
            '報名地點/甄試地點
            EAddressZip.Attributes("onblur") = "getzipname(this.value, 'ECTName','EAddressZip');"

            Dim BtnEAddressZip_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(EAddressZip, EAddressZIPB3, hidEAddressZIP6W, ECTName, EAddress)
            BtnEAddressZip.Attributes.Add("onclick", BtnEAddressZip_Attr_VAL)
        End If

        'CostSort.Attributes("onclick") = " return CostModeChange()"
        'CostMode2.Attributes("onclick") = " return CostModeChange()"
        'CostMode3.Attributes("onclick") = " return CostModeChange()"
        'CostMode4.Attributes("onclick") = " return CostModeChange()"
        CostID.Attributes("onchange") = "ShowOther('CostID','ItemOther')"
        CostID4.Attributes("onchange") = "ShowOther('CostID4','ItemOther4')"

        '計算經費來源的加總
        TNum.Attributes("onblur") = "CountCostSource();"
        DefGovCost.Attributes("onblur") = "CountCostSource();"
        DefUnitCost.Attributes("onblur") = "CountCostSource();"
        DefStdCost.Attributes("onblur") = "CountCostSource();"
        Page.RegisterStartupScript("CountCostSource", "<script>CountCostSource();</script>")

        Dim flag_UseGrant As Boolean = False
        Dim s_BlockName As String = "1111"
        If AdmGrant.Value = "1" Then '畫面回復時顯示正常化
            AdmGrant.Value = "0"
            flag_UseGrant = True
            'If Not IsPostBack Then CreateCostItem()
            'Page.RegisterStartupScript("1111", "<script>Layer_change(6);</script>")
        End If
        If AdmGrant4.Value = "1" Then '畫面回復時顯示正常化
            AdmGrant4.Value = "0"
            flag_UseGrant = True
            s_BlockName = "2524"
            'If Not IsPostBack Then CreateCostItem()
            'Page.RegisterStartupScript("2524", "<script>Layer_change(6);</script>")
        End If

        If TaxGrant.Value = "1" Then '畫面回復時顯示正常化
            TaxGrant.Value = "0"
            flag_UseGrant = True
            'If Not IsPostBack Then CreateCostItem()
            'Page.RegisterStartupScript("1111", "<script>Layer_change(6);</script>")
        End If
        If TaxGrant4.Value = "1" Then '畫面回復時顯示正常化
            TaxGrant4.Value = "0"
            flag_UseGrant = True
            s_BlockName = "2524"
            'If Not IsPostBack Then CreateCostItem()
            'Page.RegisterStartupScript("2524", "<script>Layer_change(6);</script>")
        End If
        If flag_UseGrant Then
            If Not IsPostBack Then CreateCostItem()
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript(s_BlockName, s_js11)
        End If

        '郵遞區號查詢
        LitTAddressZip.Text = TIMS.Get_WorkZIPB3Link2()
        LitEAddressZip.Text = TIMS.Get_WorkZIPB3Link2()

        '20090521 by Jimmy add 3+2郵遞區號驗證--begin
        TAddressZIPB3.Attributes("onchange") = "return CheckZIPB3_Event(this,'班別資料上課地址');"
        '報名地點/甄試地點
        EAddressZIPB3.Attributes("onchange") = "return CheckZIPB3_Event(this,'班別資料報名地點');"
        '20090521 by Jimmy add 3+2郵遞區號驗證--end

        '確認機構是否為黑名單
        'Dim vsMsg2 As String = ""
        'vsMsg2 = ""
        Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        Dim flagBlack As Boolean = False '未列入處分名單
        If TIMS.Check_OrgBlackList(Me, Hid_ComIDNO.Value, objconn) Then flagBlack = True '已列入處分名單。
        Dim vsMsg2 As String = "" '已列入處分名單(錯誤訊息。)
        If flagBlack Then
            vsMsg2 = sm.UserInfo.OrgName & "，已列入處分名單!!"
            isBlack.Value = "Y"
            Blackorgname.Value = sm.UserInfo.OrgName
            'Errmsg = vsMsg2
        End If
        If vsMsg2 <> "" Then
            Button8.Visible = False '草稿儲存
            btnAdd.Visible = False '正式處存
            TIMS.Tooltip(Button8, vsMsg2)
            TIMS.Tooltip(btnAdd, vsMsg2)
            Dim strScript As String = ""
            strScript = ""
            strScript += "<script>alert('"
            strScript += vsMsg2
            strScript += "');</script>"
            Page.RegisterStartupScript(TIMS.xBlockName, strScript)
        End If
    End Sub

    ''' <summary>每次執行</summary>
    Sub CCreate1_Every()
        '產業別(管考) true:使用/false:不可使用
        fg_USE_CBLKID60_TP06 = If(TIMS.Utl_GetConfigVAL(objconn, "USE_CBLKID60_TP06") = "Y", True, False)
        trCBLKID60.Visible = fg_USE_CBLKID60_TP06
    End Sub

    '導入網頁時 驗證基本訊息。
    Sub CREATE_PPINFO() '新增
        '新增狀態、帶入預設值
        '如果是自辦計劃，或者是委外並且是委訓登入，則帶入預設值
        If iPlanKind = 1 OrElse sm.UserInfo.LID = 2 Then Call Get_OrgPlanInfo() '取得機構資訊帶入預設值

        '建立訓練的DATATABLE
        Org.Disabled = False
        Button24.Visible = False '回上一頁

        LabCapMilitary.Visible = False '不限提示字
        CapMilitary.Visible = True 'CheckBoxList
        If sm.UserInfo.Years >= 2013 Then
            LabCapMilitary.Visible = True '不限提示字
            CapMilitary.Visible = False 'CheckBoxList
            '「受訓資格」統一選擇為「不限」
            Dim CapMValues As String = ""
            'If CapMValues <> "" Then CapMValues += ","
            'CapMValues += item.Value
            For Each item As ListItem In CapMilitary.Items
                If ("00").ToString().IndexOf(item.Value) > -1 Then
                    item.Selected = True '不限，值存在 (被選擇)
                    'Exit For
                Else
                    If CapMValues <> "" Then CapMValues += ","
                    CapMValues += item.Value
                End If
            Next
            If CapMValues <> "" Then
                Dim CapMValueA As String() = CapMValues.Split(",")
                For i As Integer = 0 To CapMValueA.Length - 1
                    If CapMValueA(i) <> "" AndAlso CapMilitary.Items.FindByValue(CapMValueA(i)) IsNot Nothing Then CapMilitary.Items.Remove(CapMilitary.Items.FindByValue(CapMValueA(i)))
                Next
            End If
            'Me.CapMilitary.Enabled = False '不可修改
            Common.SetListItem(CapMilitary, "00") '預設為不限
            Dim tmp_Tooltip As String = "" '受訓資格
            tmp_Tooltip = String.Format("「{0}」兵役-統一選擇為「不限」", lab_LayerC2.Text)
            TIMS.Tooltip(CapMilitary, tmp_Tooltip)
            TIMS.Tooltip(LabCapMilitary, tmp_Tooltip)
        End If
    End Sub

    '導入網頁時 驗證基本訊息。
    Sub SHOW_PLAN_PLANINFO() '修改/顯示
        '修改
        Button24.Visible = True '回上一頁

        Dim dtP As New DataTable
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql &= " SELECT a.*" & vbCrLf
        sql &= " ,b.OrgName" & vbCrLf
        sql &= " ,c.RID RIDValue" & vbCrLf
        sql &= " ,d.TrainID,d.TrainName" & vbCrLf
        sql &= " FROM PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN ORG_ORGINFO b ON a.ComIDNO = b.ComIDNO" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP c ON c.RID = a.RID" & vbCrLf
        sql &= " LEFT JOIN KEY_TRAINTYPE d ON a.TMID = d.TMID" & vbCrLf
        'sql &= " LEFT JOIN SHARE_CJOB cj ON a.CJOB_UNKEY = cj.CJOB_UNKEY" & vbCrLf
        sql &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNO=@SeqNO" & vbCrLf
        Using sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("PlanID", SqlDbType.VarChar).Value = rqPlanID
                .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = rqComIDNO
                .Parameters.Add("SeqNO", SqlDbType.VarChar).Value = rqSeqNO
                dtP.Load(.ExecuteReader())
            End With
        End Using
        'Dim dr As DataRow = Nothing
        Dim fg_PPIdataOK1 As Boolean = (TIMS.dtHaveDATA(dtP) AndAlso dtP.Rows.Count = 1)
        If Not fg_PPIdataOK1 Then
            '應該只能有1筆資料
            vMsg = "程式出現例外狀況(查無計畫)，請重新操作查詢!!"
            Common.MessageBox(Me, vMsg)
            Call TIMS.Utl_RespWriteEnd(Me, objconn, vMsg)
        End If
        If Not fg_PPIdataOK1 Then Return 'Exit Sub
        Dim dr As DataRow = dtP.Rows(0)

        Dim dr2 As DataRow = TIMS.GET_PLANDEPOT(rqPlanID, rqComIDNO, rqSeqNO, objconn)
        If dr2 IsNot Nothing Then
            '2019年啟用 work2019x01:2019 政府政策性產業
            Dim cvKID20 As String = Convert.ToString(dr2("KID20"))
            If gflag_ccopy Then cvKID20 = "" '(複制狀態清空)
            Call TIMS.SetCblValue(CBLKID20_1, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_2, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_3, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_4, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_5, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_6, cvKID20)

            Dim cvKID25 As String = Convert.ToString(dr2("KID25"))
            If gflag_ccopy Then cvKID25 = "" '(複制狀態清空)
            Call TIMS.SetCblValue(CBLKID25_1, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_2, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_3, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_4, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_5, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_6, cvKID25)

            '進階政策性產業類別 'Dim KID22 As String = Convert.ToString(dr2("KID22")) 'Call TIMS.SetCblValue(CBLKID22, KID22)
            Dim cvKID60 As String = Convert.ToString(dr2("KID60"))
            Call TIMS.SetCblValue(CBLKID60, cvKID60)
        End If

        If gflag_ccopy Then
            Org.Disabled = False
            RIDValue.Value = sm.UserInfo.RID
            HidOrgID.Value = TIMS.Get_OrgID(sm.UserInfo.RID, objconn)
            ComidValue.Value = TIMS.Get_ComIDNOforOrgID(HidOrgID.Value, objconn)
        Else
            Org.Disabled = True
            RIDValue.Value = dr("RID").ToString()
            ComidValue.Value = dr("ComIDNO").ToString()
            center.Text = dr("orgname").ToString()
        End If

        '班級英文名稱
        ClassEngName.Text = Convert.ToString(dr("CLASSENGNAME"))
        '訓練時段'取得鍵值-訓練時段
        Common.SetListItem(TPeriodList, Convert.ToString(dr("TPeriod")))
        TB_NOTE3.Text = Convert.ToString(dr("NOTE3"))
        '「訓練期限」
        Common.SetListItem(TDeadline_List, Convert.ToString(dr("TDeadline")))
        '導師名稱
        CTName.Text = Convert.ToString(dr("CTName"))
        '是否為輔導考照班
        RBCOACHING_Y.Checked = If(Convert.ToString(dr("COACHING")).Equals("Y"), True, False)
        RBCOACHING_N.Checked = If(Convert.ToString(dr("COACHING")).Equals("N"), True, False)
        '檢定職類代碼1/2/3 與考試級別1/2/3 KEY_EXAM3
        'txtGP1', 'txtXM1', 'txtLV1', 'EXAM1val', 'EXLV1val' 
        'Const cst_xExamLevelv1 As String = "1,2,3,4,5"
        'Const cst_xExamLeveln1 As String = "甲級,乙級,丙級,單一級,不分級"
        If RBCOACHING_Y.Checked Then
            EXAM1val.Value = Convert.ToString(dr("EXAMIDS1"))
            EXAM2val.Value = Convert.ToString(dr("EXAMIDS2"))
            EXAM3val.Value = Convert.ToString(dr("EXAMIDS3"))
            EXLV1val.Value = Convert.ToString(dr("EXAMLVID1"))
            EXLV2val.Value = Convert.ToString(dr("EXAMLVID2"))
            EXLV3val.Value = Convert.ToString(dr("EXAMLVID3"))

            Dim drEXS1 As DataRow = TIMS.Get_EXAM3DATA(objconn, Convert.ToString(dr("EXAMIDS1")), txtGP1, txtXM1)
            Dim drEXS2 As DataRow = TIMS.Get_EXAM3DATA(objconn, Convert.ToString(dr("EXAMIDS2")), txtGP2, txtXM2)
            Dim drEXS3 As DataRow = TIMS.Get_EXAM3DATA(objconn, Convert.ToString(dr("EXAMIDS3")), txtGP3, txtXM3)
            txtLV1.Text = TIMS.GET_EXAM3LEVEL(EXLV1val.Value)
            txtLV2.Text = TIMS.GET_EXAM3LEVEL(EXLV2val.Value)
            txtLV3.Text = TIMS.GET_EXAM3LEVEL(EXLV3val.Value)
            If (txtLV1.Text <> "") Then btnExamC1.Disabled = False '(暫時解除)
            If (txtLV2.Text <> "") Then btnExamC2.Disabled = False '(暫時解除)
            If (txtLV3.Text <> "") Then btnExamC3.Disabled = False '(暫時解除)
        End If

        Dim s_PCS As String = String.Format("{0}x{1}x{2}", rqPlanID, rqComIDNO, rqSeqNO)
        Dim drPC2 As DataRow = TIMS.GetPCSDate2(s_PCS, objconn)
        If drPC2 IsNot Nothing Then
            '若pp為空 cc不為空
            '班級英文名稱
            If Convert.ToString(dr("CLASSENGNAME")) = "" AndAlso Convert.ToString(drPC2("CLASSENGNAME")) <> "" Then ClassEngName.Text = Convert.ToString(drPC2("CLASSENGNAME"))
            '訓練時段'取得鍵值-訓練時段
            If Convert.ToString(dr("TPeriod")) = "" AndAlso Convert.ToString(drPC2("TPeriod")) <> "" Then Common.SetListItem(TPeriodList, Convert.ToString(drPC2("TPeriod")))
            If Convert.ToString(dr("NOTE3")) = "" AndAlso Convert.ToString(drPC2("NOTE3")) <> "" Then TB_NOTE3.Text = Convert.ToString(drPC2("NOTE3"))
            '「訓練期限」
            If Convert.ToString(dr("TDeadline")) = "" AndAlso Convert.ToString(drPC2("TDeadline")) <> "" Then Common.SetListItem(TDeadline_List, Convert.ToString(drPC2("TDeadline")))
            '導師名稱
            If Convert.ToString(dr("CTName")) = "" AndAlso Convert.ToString(drPC2("CTName")) <> "" Then CTName.Text = Convert.ToString(drPC2("CTName"))
        End If

        trainValue.Value = dr("TMID").ToString()
        TB_career_id.Text = If(dr("TMID").ToString() <> "", String.Format("[{0}]{1}", dr("TrainID").ToString(), dr("TrainName").ToString()), "")
        cjobValue.Value = dr("CJOB_UNKEY").ToString()
        'Dim dtSCJOB As DataTable = TIMS.Get_SHARECJOBdt(Me, objconn)
        txtCJOB_NAME.Text = TIMS.Get_CJOBNAME(TIMS.Get_SHARECJOBdt(Me, objconn), cjobValue.Value)
        'If dr("CJOB_UNKEY").ToString() <> "" Then txtCJOB_NAME.Text = "[" & dr("CJOB_NO").ToString() & "]" & dr("CJOB_NAME").ToString()
        '職群代碼 'If Convert.ToString(dr("JGID")) <> "" Then Common.SetListItem(ddlJGID, dr("JGID").ToString())

        Label3.Text = If(gflag_ccopy, sm.UserInfo.Years.ToString(), dr("PlanYear").ToString())
        PlanCause.Text = dr("PlanCause").ToString()
        PurScience.Text = dr("PurScience").ToString()
        PurTech.Text = dr("PurTech").ToString()
        PurMoral.Text = dr("PurMoral").ToString()
        Common.SetListItem(Degree, Convert.ToString(dr("CapDegree")))

        rdoAge1.Checked = True '年滿15歲以上
        rdoAge2.Checked = False '有上限，年滿15歲~
        If Convert.ToString(dr("CapAge2")) <> "" Then
            rdoAge1.Checked = False
            rdoAge2.Checked = True
            txtAge2.Text = "99"
            If Convert.ToString(dr("CapAge2")) <> "" Then txtAge2.Text = Val(dr("CapAge2"))
        End If

        '啟動年度2013
        'Me.CapMilitary.Enabled = True
        LabCapMilitary.Visible = False
        CapMilitary.Visible = True
        If sm.UserInfo.Years >= 2013 Then
            LabCapMilitary.Visible = True
            CapMilitary.Visible = False
            Dim CapMValues As String = "" '待移除選項
            'If CapMValues <> "" Then CapMValues += ","
            'CapMValues += item.Value
            For Each item As ListItem In CapMilitary.Items
                If ("00").ToString().IndexOf(item.Value) > -1 Then
                    item.Selected = True '不限，值存在 (被選擇)
                    'Exit For
                Else
                    If CapMValues <> "" Then CapMValues += ","
                    CapMValues += item.Value
                End If
            Next
            If CapMValues <> "" Then
                Dim CapMValueA As String() = CapMValues.Split(",")
                For i As Integer = 0 To CapMValueA.Length - 1
                    If CapMValueA(i) <> "" AndAlso Not CapMilitary.Items.FindByValue(CapMValueA(i)) Is Nothing Then
                        CapMilitary.Items.Remove(CapMilitary.Items.FindByValue(CapMValueA(i)))
                    End If
                Next
            End If
            CapMilitary.Enabled = False '不可修改
            Common.SetListItem(CapMilitary, "00") '預設為不限
            Dim tmp_Tooltip As String = "" '受訓資格
            tmp_Tooltip = String.Format("「{0}」兵役-統一選擇為「不限」", lab_LayerC2.Text)
            TIMS.Tooltip(CapMilitary, tmp_Tooltip)
            TIMS.Tooltip(LabCapMilitary, tmp_Tooltip)
        Else
            'Common.SetListItem(Solder, dr("CapMilitary").ToString())
            For Each item As ListItem In CapMilitary.Items
                If dr("CapMilitary").ToString().IndexOf(item.Value) > -1 Then item.Selected = True '值存在 (被選擇)
            Next
        End If

        TRNUNITNAME.Text = TIMS.ClearSQM(dr("TRNUNITNAME"))
        'TRNUNITCHO-委訓單位類型
        '1:政府機關/ 2:公民營事業機構/ 3:學校/ 4:團體/ 9:其他(請說明)
        Common.SetListItem(TRNUNITCHO, Convert.ToString(dr("TRNUNITCHO")))
        TRNUNITTYPE.Text = TIMS.ClearSQM(dr("TRNUNITTYPE"))
        TRNUNITEE.Text = TIMS.ClearSQM(dr("TRNUNITEE"))

        Other1.Text = TIMS.ClearSQM(dr("CapOther1"))
        Other2.Text = TIMS.ClearSQM(dr("CapOther2"))
        Other3.Text = TIMS.ClearSQM(dr("CapOther3"))
        TMScience.Text = dr("TMScience").ToString()
        TMTech.Text = dr("TMTech").ToString()
        TMScience.Text = TIMS.ClearSQM(TMScience.Text)
        TMTech.Text = TIMS.ClearSQM(TMTech.Text)

        Common.SetListItem(GetTrain1, Convert.ToString(dr("GetTrain1")))
        GetTrain2.Text = TIMS.ClearSQM(dr("GetTrain2")) '.ToString

        Dim v_GetTrain3 As String = TIMS.ClearSQM(dr("GetTrain3"))
        TIMS.SetCblValue(GetTrain3, v_GetTrain3)
        GetTrain3Other.Text = TIMS.ClearSQM(dr("GetTrain3Other")) '.ToString

        TIMS.SetCblValue(GetTrain4, Convert.ToString(dr("GetTrain4")))
        GetTrain4Other.Text = TIMS.ClearSQM(dr("GetTrain4Other")) '.ToString

        ''適用就業保險人非自願離職者一律免試入訓機制
        'If Convert.ToString(dr("InvExem")) <> "" Then
        '    Common.SetListItem(rblInvExem, dr("InvExem").ToString())
        'Else
        '    Common.SetListItem(rblInvExem, "X")
        '    'rblInvExem.SelectedIndex = -1
        'End If

        ''錄訓百分比代碼 
        'If Convert.ToString(dr("EnterPoint")) <> "" Then Common.SetListItem(ddlEnterPoint, dr("EnterPoint").ToString())
        ''手開推介單人數
        'txtEpNum.Text = Convert.ToString(dr("EpNum"))
        'txtEpNum.Enabled = False
        'If TIMS.Utl_GetConfigSet("work2013EpNum") = "Y" Then txtEpNum.Enabled = True '依參數關閉 txtEpNum

        GenSciHours.Text = TIMS.ClearSQM(Convert.ToString(dr("GenSciHours")))
        If GenSciHours.Text = "" Then GenSciHours.Text = "0"
        ProSciHours.Text = TIMS.ClearSQM(Convert.ToString(dr("ProSciHours")))
        If ProSciHours.Text = "" Then ProSciHours.Text = "0"
        Dim iGenSciHours As Double = If(GenSciHours.Text <> "", Val(GenSciHours.Text), 0)
        Dim iProSciHours As Double = If(ProSciHours.Text <> "", Val(ProSciHours.Text), 0)
        SciHours.Text = CInt(Val(iGenSciHours + iProSciHours))
        'SciHours.Text = Int(If(dr("GenSciHours").ToString() = "", 0, dr("GenSciHours"))) + Int(If(dr("ProSciHours").ToString() = "", 0, dr("ProSciHours").ToString()))
        ProTechHours.Text = TIMS.ClearSQM(Convert.ToString(dr("ProTechHours"))) '.ToString
        OtherHours.Text = TIMS.ClearSQM(Convert.ToString(dr("OtherHours"))) '.ToString
        TotalHours.Text = TIMS.ClearSQM(Convert.ToString(dr("TotalHours"))) '.ToString
        If ProTechHours.Text = "" Then ProTechHours.Text = "0"
        If OtherHours.Text = "" Then OtherHours.Text = "0"
        If TotalHours.Text = "" Then TotalHours.Text = "0"

        EMail.Text = TIMS.ClearSQM(Convert.ToString(dr("PlanEMail"))) '.ToString
        ClassName.Text = TIMS.ClearSQM(Convert.ToString(dr("ClassName"))) '.ToString

        CCTName.Text = ""
        TAddressZip.Value = ""
        hidTAddressZIP6W.Value = ""
        TAddressZIPB3.Value = ""
        If dr("TAddressZip").ToString() <> "" Then
            TAddressZip.Value = Convert.ToString(dr("TAddressZip")) 'TIMS.AddZero(Convert.ToString(dr("TAddressZip")), 3)
            hidTAddressZIP6W.Value = Convert.ToString(dr("TAddressZIP6W"))
            TAddressZIPB3.Value = TIMS.GetZIPCODEB3(hidTAddressZIP6W.Value) 'TIMS.AddZero(Convert.ToString(dr("TAddressZIPB3")), 2)
            CCTName.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("TAddressZip")), hidTAddressZIP6W.Value)
        End If
        TAddress.Text = dr("TAddress").ToString()

        '報名地點/甄試地點
        ECTName.Text = ""
        EAddressZip.Value = ""
        hidEAddressZIP6W.Value = ""
        EAddressZIPB3.Value = ""
        If dr("EAddressZip").ToString() <> "" Then
            EAddressZip.Value = Convert.ToString(dr("EAddressZip")) 'TIMS.AddZero(Convert.ToString(dr("EAddressZip")), 3)
            hidEAddressZIP6W.Value = Convert.ToString(dr("EAddressZIP6W"))
            EAddressZIPB3.Value = TIMS.GetZIPCODEB3(hidEAddressZIP6W.Value) 'TIMS.AddZero(Convert.ToString(dr("EAddressZIPB3")), 2)
            ECTName.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("EAddressZip")), hidEAddressZIP6W.Value)
        End If
        EAddress.Text = dr("EAddress").ToString()

        '是否報名時間已過
        Dim flag_Entered As Boolean = False
        Dim flag_AppliedResult As Boolean = False
        'Dim flag_TransFlag As Boolean = False 'AndAlso flag_TransFlag _
        Dim flag_IsApprPaper As Boolean = False
        If Convert.ToString(dr("SEnterDate")) <> "" Then
            If dr("AppliedResult").ToString() = "Y" Then flag_AppliedResult = True
            'If dr("TransFlag").ToString() = "Y" Then flag_TransFlag = True
            If dr("IsApprPaper") = "Y" Then flag_IsApprPaper = True
            If DateDiff(DateInterval.Day, CDate(dr("SEnterDate")), CDate(Now())) >= 0 Then
                flag_Entered = True
            End If
            '是否報名時間已過 '報名時間已過，不可修改報名時間 '且 sm.UserInfo.OrgLevel > 1 '委訓 or 縣市政府 '排除產投計畫
            If flag_Entered AndAlso flag_AppliedResult AndAlso flag_IsApprPaper AndAlso sm.UserInfo.OrgLevel > 1 Then
                SEnterDate.Enabled = False
                HR1.Enabled = False
                MM1.Enabled = False
                imgdate1.Visible = False
                FEnterDate.Enabled = False
                HR2.Enabled = False
                MM2.Enabled = False
                imgdate2.Visible = False
                TIMS.Tooltip(SEnterDate, "若要修改,請依規定進行班級變更申請,並告知承辦人.")
                TIMS.Tooltip(FEnterDate, "若要修改,請依規定進行班級變更申請,並告知承辦人.")
                TIMS.Tooltip(HR1, "若要修改,請依規定進行班級變更申請,並告知承辦人.")
                TIMS.Tooltip(HR2, "若要修改,請依規定進行班級變更申請,並告知承辦人.")
                TIMS.Tooltip(MM1, "若要修改,請依規定進行班級變更申請,並告知承辦人.")
                TIMS.Tooltip(MM2, "若要修改,請依規定進行班級變更申請,並告知承辦人.")
            End If
        End If

        '起始日
        Dim s_TmpDate1 As String = Now.ToString("yyyy/MM/dd") & " 00:00"
        TIMS.SET_DateHM(CDate(s_TmpDate1), HR1, MM1)
        SEnterDate.Text = TIMS.Cdate3(dr("SEnterDate"))
        If Convert.ToString(dr("SEnterDate")) <> "" Then TIMS.SET_DateHM(CDate(dr("SEnterDate")), HR1, MM1)

        '結束日
        Dim s_TmpDate2 As String = Now.ToString("yyyy/MM/dd") & " 23:59"
        TIMS.SET_DateHM(CDate(s_TmpDate2), HR2, MM2)
        FEnterDate.Text = TIMS.Cdate3(dr("FEnterDate"))
        If Convert.ToString(dr("FEnterDate")) <> "" Then TIMS.SET_DateHM(CDate(dr("FEnterDate")), HR2, MM2)

        '甄試日期/時段/時間 (時段：01-全天'02-上午'03-下午)
        If Convert.ToString(dr("ExamDate")) <> "" Then
            ExamDate.Text = TIMS.Cdate3(dr("ExamDate")) '01-全天'02-上午'03-下午
            TIMS.SET_DateHM(CDate(dr("ExamDate")), HR6, MM6)
            If Convert.ToString(dr("ExamPeriod")) <> "" Then Common.SetListItem(ExamPeriod, dr("ExamPeriod")) '甄試時段
        Else
            ExamDate.Text = "" 'TIMS.cdate3(dr("ExamDate")) 'ExamPeriod.SelectedIndex = -1
            Common.SetListItem(ExamPeriod, "")
            TIMS.SET_DateHM((Now.ToString("yyyy/MM/dd") & " 00:00"), HR6, MM6)
        End If

        '報到日期 
        CheckInDate.Text = ""
        If Convert.ToString(dr("CheckInDate")) <> "" Then CheckInDate.Text = TIMS.Cdate3(dr("CheckInDate"))  '201608 BY AMU 報到日期 

        ContactName.Text = ""  '聯絡人姓名
        ContactPhone.Text = "" '聯絡人電話
        ContactEmail.Text = "" '聯絡人電子郵件
        MasterEmail.Text = ""  '直屬主管電子郵件
        If Convert.ToString(dr("ContactName")) <> "" Then ContactName.Text = Convert.ToString(dr("ContactName"))
        If Convert.ToString(dr("ContactPhone")) <> "" Then ContactPhone.Text = Convert.ToString(dr("ContactPhone"))
        If Convert.ToString(dr("ContactEmail")) <> "" Then ContactEmail.Text = Convert.ToString(dr("ContactEmail"))
        If Convert.ToString(dr("MasterEmail")) <> "" Then MasterEmail.Text = Convert.ToString(dr("MasterEmail"))
        twiACTNO.Text = "" '訓字保保險證號。
        If Convert.ToString(dr("twiACTNO")) <> "" Then twiACTNO.Text = Convert.ToString(dr("twiACTNO")).ToUpper

        If gflag_ccopy Then
            THours.Text = dr("THours").ToString()
        Else
            Class_Unit.Value = dr("Class_Unit").ToString()
            Common.SetListItem(rblADVANCE, dr("ADVANCE").ToString()) '訓練課程類型 ADVANCE
            TNum.Text = dr("TNum").ToString()
            THours.Text = dr("THours").ToString()
            STDate.Text = TIMS.Cdate3(Convert.ToString(dr("STDate")))
            FDDate.Text = TIMS.Cdate3(Convert.ToString(dr("FDDate")))
            CyclType.Text = TIMS.FmtCyclType(dr("CyclType"))
            ClassCount.Text = If(dr("ClassCount").ToString() = "", "1", dr("ClassCount").ToString())

            ClassCount.Text = TIMS.ClearSQM(ClassCount.Text)
            DefGovCost.Text = dr("DefGovCost").ToString()
            DefUnitCost.Text = dr("DefUnitCost").ToString()
            DefStdCost.Text = dr("DefStdCost").ToString()
            '行政管理費百分比
            If dr("AdmPercent").ToString() <> "" Then Session("AdmGrant") = dr("AdmPercent").ToString()
            '營業稅費用百分比
            If dr("TaxPercent").ToString() <> "" Then Session("TaxGrant") = dr("TaxPercent").ToString()
            Note.Text = dr("Note").ToString()
            ESiteMsg.Text = dr("ESiteMsg").ToString()
        End If

        '已存為正式資料，而且不是要複製計畫，草稿儲存功能不啟用
        If Convert.ToString(dr("IsApprPaper")) = "Y" And Not gflag_ccopy Then Button8.Visible = False

        If Not gflag_ccopy Then
            Select Case Convert.ToString(dr("AppliedResult"))
                Case "Y", "O"
                    '2005/6/20--Melody審核通過or審核後修正者,不可修改班級名稱,期別,開結訓日,課程時數
                    If Convert.ToString(dr("IsApprPaper")) = "Y" Then

                        If Convert.ToString(dr("TransFlag")) = "Y" Then
                            CustomValidator4.Enabled = False
                            Dim v_tip1 As String = "班級轉班上架後，不可修改"

                            STDate.ReadOnly = True
                            FDDate.ReadOnly = True
                            date1.Visible = False
                            date2.Visible = False
                            TIMS.Tooltip(STDate, v_tip1)
                            TIMS.Tooltip(FDDate, v_tip1)

                            ClassName.ReadOnly = True
                            SciHours.ReadOnly = True
                            GenSciHours.ReadOnly = True
                            ProSciHours.ReadOnly = True
                            ProTechHours.ReadOnly = True
                            OtherHours.ReadOnly = True
                            TotalHours.ReadOnly = True
                            THours.ReadOnly = True
                            TIMS.Tooltip(ClassName, v_tip1)
                            TIMS.Tooltip(SciHours, v_tip1)
                            TIMS.Tooltip(GenSciHours, v_tip1)
                            TIMS.Tooltip(ProSciHours, v_tip1)
                            TIMS.Tooltip(ProTechHours, v_tip1)
                            TIMS.Tooltip(OtherHours, v_tip1)
                            TIMS.Tooltip(TotalHours, v_tip1)
                            TIMS.Tooltip(THours, v_tip1)

                            Dim titleMsg1 As String = "班級轉入後，不可修改期別"
                            Dim titleMsg2 As String = "班級轉入後，不可修改上課地址"
                            'CyclType.ReadOnly = True
                            CyclType.Enabled = False
                            TIMS.Tooltip(CyclType, titleMsg1)
                            'TAddress.ReadOnly = True
                            'CTName.ReadOnly = True
                            'CCTName.Disabled = True
                            CCTName.Enabled = False
                            TIMS.Tooltip(CCTName, titleMsg2)
                            TAddressZip.Disabled = True
                            TAddressZIPB3.Disabled = True
                            TAddress.Enabled = False

                            BtnTAddressZip.Disabled = True
                            BtnTAddressZip.Attributes.Remove("onclick")
                            'CTName.Attributes.Remove("onblur")
                            TAddressZip.Attributes.Remove("onblur")
                            TIMS.Tooltip(TAddressZip, titleMsg2)
                            TIMS.Tooltip(TAddressZIPB3, titleMsg2)
                            TIMS.Tooltip(TAddress, titleMsg2)
                            TIMS.Tooltip(BtnTAddressZip, titleMsg2)
                            '20090527 by Jimmy add 依需求於已轉班狀態下不驗證郵遞區號後2碼資料 --begin
                            RequiredFieldValidator10.EnableClientScript = False
                            CheckZIPB3_1.EnableClientScript = False
                            CheckZIPB3_2.EnableClientScript = False
                            '20090527 by Jimmy add 依需求於已轉班狀態下不驗證郵遞區號後2碼資料 --end
                        Else
                            Dim titleMsg1 As String = "班級尚未轉入，可修改"
                            Dim titleMsg2 As String = "班級尚未轉入，可修改上課地址"
                            'CyclType.ReadOnly = False
                            CyclType.Enabled = True
                            TIMS.Tooltip(CyclType, titleMsg1)
                            'TAddress.ReadOnly = False
                            'CTName.ReadOnly = False
                            CCTName.Enabled = False '.Disabled = True
                            TAddressZip.Disabled = False
                            TAddressZIPB3.Disabled = False
                            TAddress.Enabled = True
                            BtnTAddressZip.Disabled = False

                            'Dim BtnTAddressZip_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(TAddressZip, TAddressZIPB3, hidTAddressZIP6W, CCTName, TAddress)
                            'BtnTAddressZip.Attributes.Add("onclick", BtnTAddressZip_Attr_VAL)

                            'TAddressZip.Attributes("onblur") = "getZipName('CTName','TAddressZip', this.value,);"
                            'getzipname(zip,CityTextBox,ZipTextBox)
                            TAddressZip.Attributes("onblur") = "getzipname(this.value, 'CTName','TAddressZip');"
                            TIMS.Tooltip(CCTName, titleMsg2)
                            TIMS.Tooltip(TAddressZip, titleMsg2)
                            TIMS.Tooltip(TAddressZIPB3, titleMsg2)
                            TIMS.Tooltip(TAddress, titleMsg2)
                            TIMS.Tooltip(BtnTAddressZip, titleMsg2)
                            '20090527 by Jimmy add 依需求於未轉班狀態下，則要驗證郵遞區號後2碼資料 --begin
                            RequiredFieldValidator10.EnableClientScript = True
                            CheckZIPB3_1.EnableClientScript = True
                            CheckZIPB3_2.EnableClientScript = True
                            '20090527 by Jimmy add 依需求於未轉班狀態下，則要驗證郵遞區號後2碼資料 --end
                        End If
                    End If

                    blnCanUpdataTestUser = False '非測試使用環境(正式)
                    If TIMS.sUtl_ChkTest() Then blnCanUpdataTestUser = True '是測試使用環境

                    If dr("AppliedResult").ToString() = "Y" Then
                        '已經審核通過
                        If blnCanUpdataTestUser Then
                            '是測試使用環境
                            vMsg = "(已經審核通過) 是測試使用環境，提供該使用者擁有修改權限!!"
                            btnAdd.Enabled = True
                            TIMS.Tooltip(btnAdd, vMsg)
                        Else
                            '非測試使用環境(正式)
                            '自辦與委外計畫、2013年後，開放分署(中心)以上可修改
                            '就服單位協助報名
                            If sm.UserInfo.Years >= 2013 _
                                    AndAlso sm.UserInfo.LID < 2 _
                                    AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                vMsg = "(已經審核通過)，提供該使用者擁有修改權限!!"
                                btnAdd.Enabled = True
                                TIMS.Tooltip(btnAdd, vMsg)
                            Else
                                If sm.UserInfo.RoleID = "0" AndAlso sm.UserInfo.LID = "0" Then
                                    vMsg = "(已經審核通過)，提供該使用者擁有修改權限!!"
                                    btnAdd.Enabled = True
                                    TIMS.Tooltip(btnAdd, vMsg)
                                Else
                                    Select Case iPlanKind
                                        Case 1
                                            If sm.UserInfo.LID = 2 AndAlso dr("TransFlag").ToString() = "N" Then
                                                '計畫種類為1.自辦(內訓) 階層為委訓(縣市政府、一般培訓單位) ，還未轉班
                                                Call Disabled_Items("已經審核通過") '鎖定
                                            End If
                                        Case 2
                                            '計畫種類為2.委外
                                            Call Disabled_Items("已經審核通過") '鎖定
                                    End Select
                                End If
                            End If

                        End If

                        If sm.UserInfo.RoleID = "0" AndAlso sm.UserInfo.LID = "0" Then
                            vMsg = "該使用者擁有修改權限!!"
                            btnAdd.Enabled = True
                            CapMilitary.Enabled = True : TIMS.Tooltip(CapMilitary, vMsg)
                            GetTrain3.Enabled = True : TIMS.Tooltip(GetTrain3, vMsg)
                            GetTrain3Other.ReadOnly = False : TIMS.Tooltip(GetTrain3Other, vMsg)
                            GetTrain4.Enabled = True : TIMS.Tooltip(GetTrain4, vMsg)
                            GetTrain4Other.ReadOnly = False : TIMS.Tooltip(GetTrain4Other, vMsg)
                            File1.Disabled = False : TIMS.Tooltip(File1, vMsg)
                            Btn_TrainDescImport.Enabled = True : TIMS.Tooltip(Btn_TrainDescImport, vMsg)
                            Button29.Enabled = True : TIMS.Tooltip(Button29, vMsg)
                            Button11.Enabled = True : TIMS.Tooltip(Button11, vMsg)
                            Button25.Disabled = False : TIMS.Tooltip(Button25, vMsg)
                            Button25b.Disabled = False : TIMS.Tooltip(Button25b, vMsg)
                            TIMS.Tooltip(btnAdd, vMsg)
                        End If

                    End If
                Case "N"
                Case "M", ""
                    If dr("TransFlag").ToString() = "Y" Then
                        center.Enabled = False
                        Org.Disabled = True
                    End If
            End Select
        End If

        '僅允許檢視資料
        If TIMS.ClearSQM(Request("todo")) = "1" Then '按鈕狀態控制
            Call Disabled_Items("僅允許檢視資料")
            If Not Session("Redirect") Is Nothing Then
                ViewState("Redirect") = Session("Redirect")
                Session("Redirect") = Nothing
            End If
        End If
    End Sub

    ''' <summary>
    ''' 導入網頁時 驗證基本訊息。
    ''' </summary>
    Sub PageLoad_cCreate1()
        '新增
        Label3.Text = sm.UserInfo.Years
        If rqPlanID = "" Then
            Call CREATE_PPINFO() '新增
        Else
            Call SHOW_PLAN_PLANINFO() '修改
        End If
        Call SHOW_PLAN_ABILITYS() '專長能力標籤-ABILITY
        Call CerateTrainDesc()
        Call CreateCostItem()
    End Sub

    '設定年齡文字
    Sub SetlAgeStr(ByVal sAgeType As String)
        Select Case sAgeType
            Case cst_AgeGt2
                l_Age.Text = cst_AgeStr2 '設定年齡文字
                rdoAge1.Checked = True
                rdoAge2.Visible = False
                l_Age2a.Visible = False
                l_Age2b.Visible = False
                txtAge2.Visible = False
            Case cst_AgeGt3
                l_Age.Text = cst_AgeStr3 '設定年齡文字
                rdoAge1.Checked = True
                rdoAge2.Visible = False
                l_Age2a.Visible = False
                l_Age2b.Visible = False
                txtAge2.Visible = False
        End Select
    End Sub

    '檢核年齡顯示文字
    Sub ChklAgeType()
        HidAgeType.Value = cst_AgeGt1
        If cst_tplanid47age2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            HidAgeType.Value = cst_AgeGt2
        End If
        If cst_tplanid58age3.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            HidAgeType.Value = cst_AgeGt3
        End If
        Call SetlAgeStr(HidAgeType.Value) '設定年齡文字
    End Sub

    '建立下拉選單物件
    Sub cCreateItem()
        Dim sql As String = ""
        Hid_MaxTNum.Value = ""
        If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Const cst_MaxTNum As String = "20" '"計畫人數上限為20人"
            Hid_MaxTNum.Value = cst_MaxTNum
        End If

        '取得鍵值-訓練時段
        TPeriodList = TIMS.GET_HOURRAN(TPeriodList, objconn, sm)
        '取出鍵詞-訓練期限代碼'
        Call TIMS.GET_TRAINEXP(TDeadline_List, objconn, sm)
        '檢核年齡顯示文字
        Call ChklAgeType()
        '訓練費用-費用列表
        Dim dtCOSTITEM As DataTable = TIMS.GET_KEY_COSTITEMdt1(sm, objconn)
        CostID = TIMS.GET_COSTITEM(CostID, dtCOSTITEM) 'SELECT * FROM KEY_COSTITEM ORDER BY SORT
        Degree = TIMS.Get_Degree(Degree, 2, objconn)
        CostID4 = TIMS.GET_COSTITEM(CostID4, dtCOSTITEM) 'SELECT * FROM KEY_COSTITEM ORDER BY SORT

        'ddlJGID = TIMS.Get_JobGroup(ddlJGID) '職群代碼 SELECT JGID,JGNAME,SORT FROM Key_JobGroup
        'ddlEnterPoint = TIMS.Get_EnterPoint(ddlEnterPoint) '錄訓百分比代碼 SELECT KID,KNAME  FROM Key_EnterPoint
        Call TIMS.SUB_SET_HR_MI(HR1, MM1)
        Call TIMS.SUB_SET_HR_MI(HR2, MM2)
        'Dim s_TmpDate2 As String = Now.ToString("yyyy/MM/dd") & " 23:59"
        'TIMS.SET_DateHM(Now.ToString("yyyy/MM/dd") & " 23:59", HR2, MM2)
        'TIMS.SET_DateHM(String.Format("{0:yyyy/MM/dd} 23:59", Now), HR2, MM2)
        TIMS.SET_DateHM($"{Now:yyyy/MM/dd} 23:59", HR2, MM2)

        Call TIMS.SUB_SET_HR_MI(HR6, MM6)
        '01-全天'02-上午'03-下午
        ExamPeriod = TIMS.GET_ExamPeriod(ExamPeriod, objconn)

        btnExamC2.Disabled = True
        btnExamC3.Disabled = True
        '新增「是否為輔導考照班」
        '選擇「是」時，須填寫「完訓後可參加之全國技術士技能檢定職類與考試級別」
        RBCOACHING_Y.Attributes("onclick") = "show_COACHING1()"
        RBCOACHING_N.Attributes("onclick") = "show_COACHING1()"
        RBCOACHING_Y.Attributes("onchange") = "show_COACHING1()"
        RBCOACHING_N.Attributes("onchange") = "show_COACHING1()"
        TIMS.RegisterStartupScript(Me, TIMS.xBlockName(), "<script>show_COACHING1();</script>")

        'trKID20.Visible
        '2019 (政府政策性產業)
        'sql = " SELECT KID, KNAME FROM KEY_BUSINESS WHERE DEPID='20' ORDER BY KID"
        'Dim dtKID_N20 As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim dtKID_N20 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "20")
        Call TIMS.GET_CBL_KID20(CBLKID20_1, dtKID_N20, 1)
        Call TIMS.GET_CBL_KID20(CBLKID20_2, dtKID_N20, 2)
        Call TIMS.GET_CBL_KID20(CBLKID20_3, dtKID_N20, 3)
        Call TIMS.GET_CBL_KID20(CBLKID20_4, dtKID_N20, 4)
        Call TIMS.GET_CBL_KID20(CBLKID20_5, dtKID_N20, 5)
        Call TIMS.GET_CBL_KID20(CBLKID20_6, dtKID_N20, 6)
        'Dim dtKID_N22 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, 22) 'Call TIMS.GET_CBL_KID22(CBLKID22, dtKID_N22)
        'CheckBoxList 選項設定-政府政策性產業 2025-2026
        Dim dtKID_N25 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "25")
        Call TIMS.GET_CBL_KID25(CBLKID25_1, dtKID_N25, 1)
        Call TIMS.GET_CBL_KID25(CBLKID25_2, dtKID_N25, 2)
        Call TIMS.GET_CBL_KID25(CBLKID25_3, dtKID_N25, 3)
        Call TIMS.GET_CBL_KID25(CBLKID25_4, dtKID_N25, 4)
        Call TIMS.GET_CBL_KID25(CBLKID25_5, dtKID_N25, 5)
        Call TIMS.GET_CBL_KID25(CBLKID25_6, dtKID_N25, 6)

        '產業別(管考)
        If fg_USE_CBLKID60_TP06 Then
            CBLKID60.Items.Clear()
            sql = "SELECT KID,KNAME FROM VIEW_DEPOT60 ORDER BY KID"
            DbAccess.MakeListItem(CBLKID60, sql, objconn)
        End If

        'Dim v_GetTrain3 As String = TIMS.GetListValue(GetTrain3)
        'iType　1:自辦在職 / 2:70:區域產業據點職業訓練計畫(在職) 
        GetTrain3 = TIMS.Get_CBL_GetTrain3(GetTrain3, If(flag_TPlanID70_1, 2, 1))

        'If v_GetTrain3 <> "" Then TIMS.SetCblValue(GetTrain3, v_GetTrain3)
        'TRNUNITCHO-委訓單位類型
        '1:政府機關/ 2:公民營事業機構/ 3:學校/ 4:團體/ 9:其他(請說明)
        TRNUNITCHO = TIMS.GET_CBLCODE1("TRNUNITCHO", TRNUNITCHO, objconn)
    End Sub

    '取出訓練內容簡介/copy /insert /updata
    Sub CerateTrainDesc()
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        'Dim Years As String = ""
        'Const cst_pkname As String = "PTDID"
        If Session(Hid_TrainDesc_GUID1.Value) Is Nothing Then
            sql = "SELECT * FROM PLAN_TRAINDESC WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, objconn)
            dt.Columns(cst_PlanTrainDescPKName).AutoIncrement = True
            dt.Columns(cst_PlanTrainDescPKName).AutoIncrementSeed = -1
            dt.Columns(cst_PlanTrainDescPKName).AutoIncrementStep = -1
            If rqPlanID <> "" AndAlso rqComIDNO <> "" AndAlso rqSeqNO <> "" Then
                If gflag_ccopy Then
                    '2006年後,若是用copy方式，則試著取得舊資料來新增
                    'sql = "SELECT YEARS FROM ID_PLAN WHERE PlanID='" & rqPlanID & "'"
                    'Years = DbAccess.ExecuteScalar(sql, objconn)
                    sql = " SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & rqPlanID & "' and ComIDNO='" & rqComIDNO & "' and SeqNo='" & rqSeqNO & "' ORDER BY PTDID "
                    Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)
                    For Each dr1 As DataRow In dt1.Rows
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("PName") = dr1("PName")
                        dr("PHour") = dr1("PHour")
                        dr("PCont") = TIMS.ClearSQM(Convert.ToString(dr1("PCont")))
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    Next
                Else
                    sql = " SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & rqPlanID & "' and ComIDNO='" & rqComIDNO & "' and SeqNo='" & rqSeqNO & "' ORDER BY PTDID "
                    dt = DbAccess.GetDataTable(sql, objconn)
                    dt.Columns(cst_PlanTrainDescPKName).AutoIncrement = True
                    dt.Columns(cst_PlanTrainDescPKName).AutoIncrementSeed = -1
                    dt.Columns(cst_PlanTrainDescPKName).AutoIncrementStep = -1
                End If
            End If
        Else
            dt = Session(Hid_TrainDesc_GUID1.Value)
        End If
        Session(Hid_TrainDesc_GUID1.Value) = dt
        Call ShowTrainDesc() '顯示 訓練內容簡介
    End Sub

    '計畫經費項目檔(PLAN_COSTITEM)
    Sub CreateCostItem()
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        'Dim dr As DataRow
        Dim i_Total As Double = 0 '總費用
        Dim i_AdmTotal As Double = 0 '行政管理費 '(行政管理費百分比)
        Dim i_TaxTotal As Double = 0 '營業稅
        Dim i_cost_02_08 As Double = 0 '材料費
        Dim strAdmCostTxt As String = ""
        Dim strTaxCostTxt As String = ""

        If Session(Hid_CostItem_GUID1.Value) Is Nothing Then
            sql = " SELECT * FROM PLAN_COSTITEM WHERE 1<>1 "
            dt = DbAccess.GetDataTable(sql, objconn)
            dt.Columns("PCID").AutoIncrement = True
            dt.Columns("PCID").AutoIncrementSeed = -1
            dt.Columns("PCID").AutoIncrementStep = -1
            If rqPlanID <> "" AndAlso rqComIDNO <> "" AndAlso rqSeqNO <> "" Then
                If gflag_ccopy Then
                    sql = " SELECT * FROM PLAN_COSTITEM WHERE PlanID='" & rqPlanID & "' and ComIDNO='" & rqComIDNO & "' and SeqNO='" & rqSeqNO & "' ORDER BY PCID"
                    Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)
                    For Each dr1 As DataRow In dt1.Rows
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        'dr("PCID") = dr1("PCID")
                        dr("CostMode") = dr1("CostMode")
                        dr("CostID") = dr1("CostID")
                        dr("ItemOther") = dr1("ItemOther")
                        dr("OPrice") = dr1("OPrice")
                        dr("Itemage") = dr1("Itemage")
                        dr("ItemCost") = dr1("ItemCost")
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    Next
                Else
                    sql = "SELECT * FROM PLAN_COSTITEM WHERE PlanID='" & rqPlanID & "' AND ComIDNO='" & rqComIDNO & "' AND SeqNO='" & rqSeqNO & "' ORDER BY PCID "
                    dt = DbAccess.GetDataTable(sql, objconn)
                    dt.Columns("PCID").AutoIncrement = True
                    dt.Columns("PCID").AutoIncrementSeed = -1
                    dt.Columns("PCID").AutoIncrementStep = -1
                End If
            End If
            Session(Hid_CostItem_GUID1.Value) = dt
        Else
            dt = Session(Hid_CostItem_GUID1.Value)
        End If

        Dim dv As New DataView
        dt.TableName = "CostItem"
        dv.Table = dt

        If iPlanKind = 1 Then
            'PlanKind (1) 1:自辦 2:委外
            DataGrid1Table.Style.Item("display") = ""
            If dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows).Length = 0 Then
                DataGrid1Table.Style.Item("display") = "none"
            End If
            With DataGrid1
                .DataSource = dt
                .DataKeyField = "PCID"
                .DataBind()
            End With

            '行政管理費 '(行政管理費百分比)
            If Not Session("AdmGrant") Is Nothing Then
                Dim Flag_tmp1 As Boolean = False
                For Each dr As DataRow In dt.Rows
                    If Convert.ToString(dr("AdmFlag")) = "Y" Then
                        Flag_tmp1 = True
                        Exit For
                    End If
                Next

                AdmGrantTR.Visible = False
                If Flag_tmp1 And Session("AdmGrant") <> "" Then
                    i_AdmTotal = 0
                    AdmCost.Text = ""
                    AdmGrantTR.Visible = True
                    strAdmCostTxt = ""
                    For Each dr As DataRow In dt.Select("AdmFlag='Y'")
                        i_AdmTotal += CDbl(dr("OPrice")) * CDbl(dr("Itemage")) * CDbl(dr("ItemCost"))
                        Dim strCostName As String = Get_xCostName2(dtKEYCOSTITEM, dr)
                        If strAdmCostTxt <> "" Then strAdmCostTxt &= "+"
                        strAdmCostTxt &= strCostName
                    Next
                    If strAdmCostTxt <> "" Then
                        AdmCost.Text = "(" & strAdmCostTxt & ")"
                        AdmCost.Text &= "*" & Int(Session("AdmGrant")) & "%=" & TIMS.ROUND(Int(Session("AdmGrant")) * i_AdmTotal / 100)
                    End If
                End If
            End If

            '營業稅 '(營業稅費用百分比)
            If Not Session("TaxGrant") Is Nothing Then
                Dim Flag_tmp2 As Boolean = False
                For Each dr As DataRow In dt.Rows
                    If Convert.ToString(dr("TaxFlag")) = "Y" Then
                        Flag_tmp2 = True
                        Exit For
                    End If
                Next
                TaxGrantTR.Visible = False
                If Flag_tmp2 And Session("TaxGrant") <> "" Then
                    i_TaxTotal = 0
                    TaxCost.Text = ""
                    TaxGrantTR.Visible = True

                    strTaxCostTxt = ""
                    For Each dr As DataRow In dt.Select("TaxFlag='Y'")
                        i_TaxTotal += CDbl(dr("OPrice")) * CDbl(dr("Itemage")) * CDbl(dr("ItemCost"))
                        Dim strCostName As String = Get_xCostName2(dtKEYCOSTITEM, dr)
                        If strTaxCostTxt <> "" Then strTaxCostTxt &= "+"
                        strTaxCostTxt &= strCostName
                    Next
                    If strTaxCostTxt <> "" Then
                        TaxCost.Text = "(" & strTaxCostTxt & ")"
                        TaxCost.Text &= "*" & Int(Session("TaxGrant")) & "%=" & TIMS.ROUND(Int(Session("TaxGrant")) * i_TaxTotal / 100)
                    End If
                End If
            End If

            i_Total = 0
            For Each item As DataGridItem In DataGrid1.Items
                If Not IsNumeric(item.Cells(4).Text) Then item.Cells(4).Text = "0"
                i_Total += CDbl(item.Cells(4).Text)
            Next
            If AdmGrantTR.Visible = True Then   '行政管理費 '(行政管理費百分比)
                i_Total += CDbl(TIMS.ROUND(Int(Session("AdmGrant")) * i_AdmTotal / 100))
            End If
            If TaxGrantTR.Visible = True Then '營業稅 '(營業稅費用百分比)
                i_Total += CDbl(TIMS.ROUND(Int(Session("TaxGrant")) * i_TaxTotal / 100))
            End If
            TotalCost1.Text = TIMS.ROUND(i_Total)
            TableCost1.Style("display") = ""
        Else
            'PlanKind (2) 1:自辦 2:委外
            'If Request(cst_ccopy) = 1 Then CostMode = GetCostMode(1)
            If TIMS.dtNODATA(dt) Then
                DataGrid2Table.Style.Item("display") = "none"
                DataGrid3Table.Style.Item("display") = "none"
                DataGrid4Table.Style.Item("display") = "none"

                i_CostMode = GetCostMode(1)
                If i_CostMode = 0 Then Return 'Exit Sub '異常離開
            Else
                If Not IsPostBack Then '計價方案以第1次登入為準
                    Dim dr As DataRow = dt.Rows(0)
                    Select Case dr("CostMode")
                        Case 2
                            TableCost2.Style("display") = ""
                        Case 3
                            TableCost3.Style("display") = ""
                        Case 4
                            TableCost4.Style("display") = ""
                    End Select
                End If

                '每人每時單價計價法- -Start
                i_Total = 0
                dv.RowFilter = "CostMode=2"
                DataGrid2Table.Style.Item("display") = If(dv.Count > 0, "", "none")
                With DataGrid2
                    .DataSource = dv
                    .DataKeyField = "PCID"
                    .DataBind()
                End With

                For Each item As DataGridItem In DataGrid2.Items
                    i_Total += CDbl(item.Cells(3).Text)
                Next
                TotalCost2.Text = TIMS.ROUND(i_Total)
                '每人每時單價計價法- -End

                '每人輔助單價計價法- -Start
                i_Total = 0
                dv.RowFilter = "CostMode=3"
                DataGrid3Table.Style.Item("display") = If(dv.Count > 0, "", "none")
                With DataGrid3
                    .DataSource = dv
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
                For Each item As DataGridItem In DataGrid3.Items
                    i_Total += CDbl(item.Cells(2).Text)
                Next
                TotalCost3.Text = TIMS.ROUND(i_Total)
                '每人輔助單價計價法- -End

                '個人單價計價法- -Start
                i_Total = 0
                dv.RowFilter = "CostMode=4"
                DataGrid4Table.Style.Item("display") = If(dv.Count > 0, "", "none")
                With DataGrid4
                    .DataSource = dv
                    .DataKeyField = "PCID"
                    .DataBind()
                End With

                '行政管理費
                AdmTR4.Visible = False
                If Not Session("AdmGrant") Is Nothing Then
                    Dim Flag As Boolean = False
                    '檢查選項中是否有勾選的行政管理費
                    For Each dr As DataRow In dt.Select("CostMode=4")
                        If dr("AdmFlag").ToString() = "Y" Then
                            Flag = True
                            Exit For
                        End If
                    Next
                    'AdmTR4.Visible = False
                    If Flag And Session("AdmGrant") <> "" Then
                        i_AdmTotal = 0
                        AdmCost4.Text = ""
                        AdmTR4.Visible = True
                        strAdmCostTxt = ""
                        For Each dr As DataRow In dt.Select("AdmFlag='Y' and CostMode='4'")
                            i_AdmTotal += CDbl(dr("OPrice")) * CDbl(dr("Itemage"))
                            Dim strCostName As String = Get_xCostName2(dtKEYCOSTITEM, dr)
                            If strAdmCostTxt <> "" Then strAdmCostTxt &= "+"
                            strAdmCostTxt &= strCostName
                        Next
                        If strAdmCostTxt <> "" Then
                            AdmCost.Text = "(" & strAdmCostTxt & ")"
                            AdmCost.Text &= "*" & Int(Session("AdmGrant")) & "%=" & TIMS.ROUND(Int(Session("AdmGrant")) * i_AdmTotal / 100)
                        End If
                    End If
                End If

                '營業稅 '(營業稅費用百分比)
                TaxTR4.Visible = False
                If Not Session("TaxGrant") Is Nothing Then
                    Dim Flag As Boolean = False             '檢查選項中是否有勾選的營業稅
                    For Each dr As DataRow In dt.Select("CostMode=4")
                        If dr("TaxFlag").ToString() = "Y" Then
                            Flag = True
                            Exit For
                        End If
                    Next
                    If Flag And Session("TaxGrant") <> "" Then
                        i_TaxTotal = 0
                        TaxCost4.Text = ""
                        TaxTR4.Visible = True
                        strTaxCostTxt = ""
                        For Each dr As DataRow In dt.Select("TaxFlag='Y' and CostMode='4'") '1.自辦 2.每人每時 3.每人輔助 4.個人單價, 5.產學訓專用
                            i_TaxTotal += CDbl(dr("OPrice")) * CDbl(dr("Itemage"))
                            Dim strCostName As String = Get_xCostName2(dtKEYCOSTITEM, dr)
                            If strTaxCostTxt <> "" Then strTaxCostTxt &= "+"
                            strTaxCostTxt &= strCostName
                        Next
                        If strTaxCostTxt <> "" Then
                            TaxCost.Text = "(" & strTaxCostTxt & ")"
                            TaxCost.Text &= "*" & Int(Session("TaxGrant")) & "%=" & TIMS.ROUND(Int(Session("TaxGrant")) * i_TaxTotal / 100)
                        End If
                    End If
                End If

                For Each item As DataGridItem In DataGrid4.Items
                    i_Total += CDbl(item.Cells(3).Text)
                Next
                If AdmTR4.Visible = True Then i_Total += CDbl(TIMS.ROUND(Int(Session("AdmGrant")) * i_AdmTotal / 100)) '行政管理費 '(行政管理費百分比)

                ViewState("hidTaxCost4") = 0 '營業稅 hidden設定
                If TaxTR4.Visible = True Then '營業稅 '(營業稅費用百分比)
                    i_Total += CDbl(TIMS.ROUND(Int(Session("TaxGrant")) * i_TaxTotal / 100))
                    ViewState("hidTaxCost4") = CDbl(TIMS.ROUND(Int(Session("TaxGrant")) * i_TaxTotal / 100))
                End If

                TotalCost4.Text = TIMS.ROUND(i_Total)
                '營業稅 hidden設定
                If IsNumeric(ViewState("hidTaxCost4")) Then
                    If CInt(ViewState("hidTaxCost4")) > 0 Then
                        ViewState("hidTaxCost4") = CInt(ViewState("hidTaxCost4"))
                    Else
                        ViewState("hidTaxCost4") = 0
                    End If
                Else
                    ViewState("hidTaxCost4") = 0
                End If

                Const cst_err_msg_1 As String = "尚未設定人數或人數並非為數字"
                PerCost.Text = cst_err_msg_1 '"尚未設定人數或人數並非為數字"
                If IsNumeric(TNum.Text) Then
                    Try
                        TNum.Text = CInt(TNum.Text)
                        hidTaxCost4.Value = ViewState("hidTaxCost4")
                        PerCost.Text = TIMS.ROUND((Int(TotalCost4.Text) - CInt(hidTaxCost4.Value)) / Int(TNum.Text))
                    Catch ex As Exception
                        PerCost.Text = cst_err_msg_1 '"尚未設定人數或人數並非為數字"
                    End Try
                End If

                'Hid_TPlanID.Value 
                If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Hid_cost_02_08.Value = TIMS.ROUND(i_cost_02_08)
                    Dim v_ACTHUMCOST As Double = TIMS.ROUND(Val(TotalCost4.Text) / Val(TotalHours.Text) / Val(TNum.Text), 2)
                    Dim v_METCOSTPER As Double = TIMS.ROUND(Val(Hid_cost_02_08.Value) / Val(TotalCost4.Text), 2)
                    ACTHUMCOST.Text = v_ACTHUMCOST
                    METCOSTPER.Text = v_METCOSTPER
                End If

                ' If (Hid_TPlanID && Hid_TPlanID.Value == '70' && ACTHUMCOST && METCOSTPER && v_TotalHours > 0) {
                ' // 總計 / 訓練時數 / 訓練人數
                ' //總計金額-TotalCost4/訓練時數-TotalHours/訓練人數-v_TNum
                ' ACTHUMCOST.innerHTML = toDecimal(toDecimal(toDecimal(v_TotalCost4) / v_TotalHours) / v_TNum);
                ' //parseInt(TotalHours.value , 10)
                ' //材料費小計/總計-TotalCost4
                ' METCOSTPER.innerHTML = toDecimal(toDecimal(v_Hid_cost04) / v_TotalHours);

                '個人單價計價法- -End
            End If
        End If
    End Sub

    '顯示 訓練內容簡介
    Sub ShowTrainDesc()
        HPHour.Value = 0
        DataGrid5Table.Visible = False
        If Session(Hid_TrainDesc_GUID1.Value) IsNot Nothing Then
            Dim dt As DataTable = Session(Hid_TrainDesc_GUID1.Value)
            'Me.HPHour.Value = 0
            'DataGrid5Table.Visible = False
            If TIMS.dtHaveDATA(dt) Then
                'Me.HPHour.Value = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then
                        If dr("Phour").ToString() <> "" Then HPHour.Value += CInt(dr("Phour").ToString())
                    End If
                Next
                DataGrid5Table.Visible = True
                DataGrid5.DataSource = dt
                DataGrid5.DataBind()
            End If
        End If
    End Sub

    '取出計價模式(PLAN_COSTCATE)
    Function GetCostMode(ByVal i_showMsg As Integer) As Integer
        'showMsg: 0:不秀訊息 1:秀訊息
        Dim iRst As Integer = 0 '(0:異常)
        '1.成本加工費法(限定自辦使用) '2.每人每時單價計價法 '3.每人輔助單價計價法 '4.個人單價計價法
        Dim dt As New DataTable
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = " SELECT * FROM PLAN_COSTCATE WHERE TPlanID=@TPlanID "
        Using sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
                dt.Load(.ExecuteReader())
            End With
        End Using
        Dim dr As DataRow = Nothing
        If TIMS.dtHaveDATA(dt) Then dr = dt.Rows(0)
        'gda = TIMS.GetOneDA(objconn)
        'gda.SelectCommand.CommandText = sql
        'gda.SelectCommand.Parameters.Clear()
        'gda.SelectCommand.Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
        'Dim dr As DataRow = TIMS.GetOneRow(gda)
        If dr Is Nothing Then
            If i_showMsg = 1 Then
                vMsg = "尚未設定此訓練計畫的計價種類!(請洽系統管理者協助)"
                Common.MessageBox(Me, vMsg)
                Call TIMS.Utl_RespWriteEnd(Me, objconn, vMsg)
                Return 0
            End If
            Button8.Enabled = False
            btnAdd.Enabled = False
            TIMS.Tooltip(Button8, vMsg)
            TIMS.Tooltip(btnAdd, vMsg)
            Return 0
        End If

        iRst = dr("CateNo")
        Select Case iRst
            Case 2 '2.每人每時
                TableCost2.Style("display") = ""
            Case 3 '3.每人輔助
                TableCost3.Style("display") = ""
            Case 4 '4.個人單價
                TableCost4.Style("display") = ""
            Case 1
                If i_showMsg = 1 Then
                    vMsg = "此訓練計畫採成本加工費法(計價種類限定使用:自辦 非委外)!!"
                    Common.MessageBox(Me, vMsg)
                End If
                Button8.Enabled = False
                btnAdd.Enabled = False
                TIMS.Tooltip(Button8, vMsg)
                TIMS.Tooltip(btnAdd, vMsg)
                iRst = 0
            Case Else
                iRst = 0
        End Select
        Return iRst
    End Function

    '取得SeqNO 依 Trans
    Public Shared Function GetMaxSeqNum(ByVal Trans As SqlTransaction, ByRef PlanID As String, ByRef ComIDNO As String) As Integer
        Dim iRst As Integer = 1
        '取得SeqNO
        Dim sql As String = ""
        sql &= " SELECT PLANID ,COMIDNO ,SEQNO"
        sql &= " FROM PLAN_PLANINFO "
        sql &= " WHERE PLANID='" & PlanID & "' AND COMIDNO='" & ComIDNO & "'"
        sql &= " ORDER BY SEQNO DESC "
        Dim dr As DataRow = DbAccess.GetOneRow(sql, Trans)
        If dr IsNot Nothing Then iRst = dr("SeqNO") + 1
        Return iRst
    End Function

    ''' <summary>檢查匯入資料</summary>
    ''' <param name="colArray"></param>
    ''' <returns></returns>
    Function CheckImportData(ByVal colArray As Array) As String
        Dim Reason As String = ""
        Const cst_filedNum As Integer = 3
        Const cst_必須填寫 As String = "必須填寫"
        'Dim sql As String
        'Dim dr As DataRow
        If colArray.Length <> cst_filedNum Then
            'Reason += "欄位數量不正確(應該為" & cst_filedNum & "個欄位)" & vbCrLf
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
        Else
            Dim PName As String = colArray(0).ToString() '單元名稱
            Dim PHour As String = colArray(1).ToString() '時數
            Dim PCont As String = colArray(2).ToString() '課程大綱
            PCont = TIMS.ClearSQM(PCont)
            If PName = "" Then Reason += cst_必須填寫 & "單元名稱" & vbCrLf
            If PHour = "" Then
                Reason += cst_必須填寫 & "時數" & vbCrLf
            Else
                If Not IsNumeric(PHour) Then
                    Reason += "時數" & "必須為數字" & vbCrLf
                Else
                    If Not TIMS.IsNumeric2(PHour) Then Reason += "時數" & "必須為數字" & vbCrLf
                End If
            End If
            If PCont = "" Then Reason += cst_必須填寫 & "課程大綱" & vbCrLf
        End If
        Return Reason
    End Function

    '鎖定
    Sub Disabled_Items(ByVal titlemsg As String)
        EMail.ReadOnly = True : TIMS.Tooltip(EMail, titlemsg)
        trainValue.Disabled = True : TIMS.Tooltip(trainValue, titlemsg)
        cjobValue.Disabled = True : TIMS.Tooltip(cjobValue, titlemsg)
        PlanCause.ReadOnly = True : TIMS.Tooltip(PlanCause, titlemsg)
        PurScience.ReadOnly = True : TIMS.Tooltip(PurScience, titlemsg)
        PurTech.ReadOnly = True : TIMS.Tooltip(PurTech, titlemsg)
        PurMoral.ReadOnly = True : TIMS.Tooltip(PurMoral, titlemsg)
        Degree.Enabled = False : TIMS.Tooltip(Degree, titlemsg)
        'Me.rblAge.Enabled = False : TIMS.Tooltip(rblAge, titlemsg)
        'Me.Age_l.ReadOnly = True : TIMS.Tooltip(Age_l, titlemsg)
        'Me.Age_u.ReadOnly = True : TIMS.Tooltip(Age_u, titlemsg)
        rdoAge1.Enabled = False : TIMS.Tooltip(rdoAge1, titlemsg)
        rdoAge2.Enabled = False : TIMS.Tooltip(rdoAge2, titlemsg)
        txtAge2.Enabled = False : TIMS.Tooltip(txtAge2, titlemsg)
        'Me.Sex.Enabled = False : TIMS.Tooltip(Sex, titlemsg)
        'Me.Solder.Enabled = False
        CapMilitary.Enabled = False : TIMS.Tooltip(CapMilitary, titlemsg)
        Other1.ReadOnly = True : TIMS.Tooltip(Other1, titlemsg)
        Other2.ReadOnly = True : TIMS.Tooltip(Other2, titlemsg)
        Other3.ReadOnly = True : TIMS.Tooltip(Other3, titlemsg)

        TMScience.ReadOnly = True : TIMS.Tooltip(TMScience, titlemsg)
        TMTech.ReadOnly = True : TIMS.Tooltip(TMTech, titlemsg)
        GetTrain1.Enabled = False : TIMS.Tooltip(GetTrain1, titlemsg)
        GetTrain2.ReadOnly = True : TIMS.Tooltip(GetTrain2, titlemsg)
        GetTrain3.Enabled = False : TIMS.Tooltip(GetTrain3, titlemsg)
        GetTrain3Other.ReadOnly = True : TIMS.Tooltip(GetTrain3Other, titlemsg)
        GetTrain4.Enabled = False : TIMS.Tooltip(GetTrain4, titlemsg)
        GetTrain4Other.ReadOnly = True : TIMS.Tooltip(GetTrain4Other, titlemsg)
        SciHours.ReadOnly = True : TIMS.Tooltip(SciHours, titlemsg)
        GenSciHours.ReadOnly = True : TIMS.Tooltip(GenSciHours, titlemsg)
        ProSciHours.ReadOnly = True : TIMS.Tooltip(ProSciHours, titlemsg)
        ProTechHours.ReadOnly = True : TIMS.Tooltip(ProTechHours, titlemsg)
        TotalHours.ReadOnly = True : TIMS.Tooltip(TotalHours, titlemsg)
        ClassName.ReadOnly = True : TIMS.Tooltip(ClassName, titlemsg)
        rblADVANCE.Enabled = False : TIMS.Tooltip(rblADVANCE, titlemsg) '訓練課程類型
        TNum.ReadOnly = True : TIMS.Tooltip(TNum, titlemsg)
        THours.ReadOnly = True : TIMS.Tooltip(THours, titlemsg)
        STDate.ReadOnly = True : TIMS.Tooltip(STDate, titlemsg)
        FDDate.ReadOnly = True : TIMS.Tooltip(FDDate, titlemsg)
        CyclType.ReadOnly = True : TIMS.Tooltip(CyclType, titlemsg)
        CustomValidator4.Enabled = False : TIMS.Tooltip(CustomValidator4, titlemsg)
        ClassCount.ReadOnly = True : TIMS.Tooltip(ClassCount, titlemsg)
        'CCTName.Disabled = True : TIMS.Tooltip(CCTName, titlemsg)
        CCTName.Enabled = False : TIMS.Tooltip(CCTName, titlemsg)
        TAddressZip.Disabled = True : TIMS.Tooltip(TAddressZip, titlemsg)
        TAddressZIPB3.Disabled = True : TIMS.Tooltip(TAddressZIPB3, titlemsg)
        'Me.TAddress.ReadOnly = True : TIMS.Tooltip(TAddress, titlemsg)
        TAddress.Enabled = False : TIMS.Tooltip(TAddress, titlemsg)
        'Me.ECTName.Disabled = True : TIMS.Tooltip(ECTName, titlemsg)
        ECTName.Enabled = False : TIMS.Tooltip(ECTName, titlemsg)
        EAddressZip.Disabled = True : TIMS.Tooltip(EAddressZip, titlemsg)
        EAddressZIPB3.Disabled = True : TIMS.Tooltip(EAddressZIPB3, titlemsg)
        'Me.TAddress.ReadOnly = True : TIMS.Tooltip(TAddress, titlemsg)
        EAddress.Enabled = False : TIMS.Tooltip(EAddress, titlemsg)
        SEnterDate.Enabled = False : TIMS.Tooltip(SEnterDate, titlemsg)
        HR1.Enabled = False : TIMS.Tooltip(HR1, titlemsg)
        MM1.Enabled = False : TIMS.Tooltip(MM1, titlemsg)
        FEnterDate.Enabled = False : TIMS.Tooltip(FEnterDate, titlemsg)
        HR2.Enabled = False : TIMS.Tooltip(HR2, titlemsg)
        MM2.Enabled = False : TIMS.Tooltip(MM2, titlemsg)
        ExamDate.Enabled = False : TIMS.Tooltip(ExamDate, titlemsg)
        ExamPeriod.Enabled = False : TIMS.Tooltip(ExamPeriod, titlemsg)
        CheckInDate.Enabled = False : TIMS.Tooltip(CheckInDate, titlemsg) '報到日期
        ContactName.ReadOnly = True : TIMS.Tooltip(ContactName, titlemsg)
        ContactPhone.ReadOnly = True : TIMS.Tooltip(ContactPhone, titlemsg)
        ContactEmail.ReadOnly = True : TIMS.Tooltip(ContactEmail, titlemsg)
        MasterEmail.ReadOnly = True : TIMS.Tooltip(MasterEmail, titlemsg)
        twiACTNO.ReadOnly = True : TIMS.Tooltip(twiACTNO, titlemsg) '訓字保保險證號
        DefGovCost.ReadOnly = True : TIMS.Tooltip(DefGovCost, titlemsg)
        DefUnitCost.ReadOnly = True : TIMS.Tooltip(DefUnitCost, titlemsg)
        DefStdCost.ReadOnly = True : TIMS.Tooltip(DefStdCost, titlemsg)
        Note.ReadOnly = True : TIMS.Tooltip(Note, titlemsg)
        ESiteMsg.ReadOnly = True : TIMS.Tooltip(ESiteMsg, titlemsg)
        center.Enabled = False : TIMS.Tooltip(center, titlemsg)
        Org.Disabled = True : TIMS.Tooltip(Org, titlemsg)
        'btnAdd.Visible = False : TIMS.Tooltip(btnAdd, titlemsg)
        'Button8.Visible = False : TIMS.Tooltip(Button8, titlemsg)
        btnAdd.Enabled = False : TIMS.Tooltip(btnAdd, titlemsg)
        Button8.Enabled = False : TIMS.Tooltip(Button8, titlemsg)
        Button2.Enabled = False : TIMS.Tooltip(Button2, titlemsg)
        Button3.Disabled = True : TIMS.Tooltip(Button3, titlemsg)
        Button3b.Disabled = True : TIMS.Tooltip(Button3b, titlemsg)
        File1.Disabled = True : TIMS.Tooltip(File1, titlemsg)
        Btn_TrainDescImport.Enabled = False : TIMS.Tooltip(Btn_TrainDescImport, titlemsg)
        Button29.Enabled = False : TIMS.Tooltip(Button29, titlemsg)
        Button26.Disabled = True : TIMS.Tooltip(Button26, titlemsg)
        Button34.Disabled = True : TIMS.Tooltip(Button34, titlemsg)
        Button9.Enabled = False : TIMS.Tooltip(Button9, titlemsg)
        Button10.Enabled = False : TIMS.Tooltip(Button10, titlemsg)
        Button11.Enabled = False : TIMS.Tooltip(Button11, titlemsg)
        Button25.Disabled = True : TIMS.Tooltip(Button25, titlemsg)
        Button25b.Disabled = True : TIMS.Tooltip(Button25b, titlemsg)
        btu_sel.Disabled = True : TIMS.Tooltip(btu_sel, titlemsg)
        btu_sel2.Disabled = True : TIMS.Tooltip(btu_sel2, titlemsg)
        BtnTAddressZip.Disabled = True : TIMS.Tooltip(BtnTAddressZip, titlemsg)
        BtnEAddressZip.Disabled = True : TIMS.Tooltip(BtnTAddressZip, titlemsg)
    End Sub

    ''' <summary>要刪除 - PLAN_VERRECORD </summary>
    ''' <param name="pParms"></param>
    ''' <param name="oTrans"></param>
    Public Shared Sub DEL_PLANVERRECORD(ByRef pParms As Hashtable, ByRef oTrans As SqlTransaction)
        '過濾可疑字 'Dim ssPlanID As String = sm.UserInfo.PlanID 
        'ComidValue.Value = TIMS.ClearSQM(ComidValue.Value) 'ssPlanID = TIMS.ClearSQM(ssPlanID)
        Dim v_PlanID As String = TIMS.GetMyValue2(pParms, "PlanID")
        Dim v_ComidValue As String = TIMS.GetMyValue2(pParms, "ComidValue")
        Dim v_SeqNo As String = TIMS.GetMyValue2(pParms, "SeqNo")
        Dim d_sql As String = " DELETE PLAN_VERRECORD WHERE PlanID='" & v_PlanID & "' AND ComIDNO='" & v_ComidValue & "' AND SeqNo='" & v_SeqNo & "'"
        DbAccess.ExecuteNonQuery(d_sql, oTrans)
    End Sub

    ''' <summary> '(SAVE) INSERT PLAN_DEPOT (SAVE) </summary>
    ''' <param name="sSearchW"></param>
    ''' <param name="oTrans"></param>
    Sub SAVE_PLAN_DEPOT(ByVal sSearchW As String, ByVal oTrans As SqlTransaction)
        '確認
        Dim PlanID As String = TIMS.GetMyValue(sSearchW, "PlanID")
        Dim ComIDNO As String = TIMS.GetMyValue(sSearchW, "ComIDNO")
        Dim SeqNo As String = TIMS.GetMyValue(sSearchW, "SeqNo")

        Dim KID20 As String = TIMS.GetMyValue(sSearchW, "KID20")
        Dim KID25 As String = TIMS.GetMyValue(sSearchW, "KID25")
        Dim KID60 As String = TIMS.GetMyValue(sSearchW, "KID60")

        If PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "" Then Return 'Exit Sub

        Dim s_Parms As New Hashtable From {{"PLANID", PlanID}, {"COMIDNO", ComIDNO}, {"SEQNO", SeqNo}}
        Dim sql As String = " SELECT 'X' FROM dbo.PLAN_DEPOT WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, oTrans, s_Parms)

        If TIMS.dtNODATA(dt1) Then
            'INSERT
            Dim iCmd_Parms As New Hashtable From {
                {"PLANID", PlanID},
                {"COMIDNO", ComIDNO},
                {"SEQNO", SeqNo},
                {"KID20", If(KID20 <> "", KID20, Convert.DBNull)},
                {"KID25", If(KID25 <> "", KID25, Convert.DBNull)},
                {"KID60", If(KID60 <> "", KID60, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim i_sql As String = ""
            i_sql &= " INSERT INTO PLAN_DEPOT(PLANID ,COMIDNO ,SEQNO ,KID20,KID25,KID60 ,APPRESULT ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
            i_sql &= " VALUES (@PLANID ,@COMIDNO ,@SEQNO ,@KID20,@KID25,@KID60 ,'Y' ,@MODIFYACCT ,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(i_sql, oTrans, iCmd_Parms)
        Else
            'UPDATE
            Dim uCmd_Parms As New Hashtable From {
                {"KID20", If(KID20 <> "", KID20, Convert.DBNull)},
                {"KID25", If(KID25 <> "", KID25, Convert.DBNull)},
                {"KID60", If(KID60 <> "", KID60, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID},
                {"PLANID", PlanID},
                {"COMIDNO", ComIDNO},
                {"SEQNO", SeqNo}
            }
            Dim u_sql As String = ""
            u_sql &= " UPDATE PLAN_DEPOT SET KID20=@KID20,KID25=@KID25,KID60=@KID60" & vbCrLf
            u_sql &= " ,APPRESULT='Y',MODIFYACCT=@MODIFYACCT ,MODIFYDATE = GETDATE()" & vbCrLf
            u_sql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
            DbAccess.ExecuteNonQuery(u_sql, oTrans, uCmd_Parms)
        End If
    End Sub

    ''' <summary>
    ''' 年齡上下限制
    ''' </summary>
    ''' <param name="o_rdoAge2"></param>
    ''' <param name="AgeTypeValue"></param>
    ''' <param name="Age2Val"></param>
    ''' <param name="iType"></param>
    ''' <returns></returns>
    Public Shared Function Get_CapAge_Val(ByRef o_rdoAge2 As RadioButton, ByVal AgeTypeValue As String, ByVal Age2Val As String, ByVal iType As Integer) As Integer
        'If rdoAge1.Checked Then iAgeType1 = 1
        Dim iAgeType1 As Integer = 1
        If o_rdoAge2.Checked Then iAgeType1 = 2
        Dim i_CapAge1 As Integer = -1
        Dim i_CapAge2 As Integer = -1
        Select Case iAgeType1
            Case 1
                Select Case AgeTypeValue'HidAgeType.Value
                    Case cst_AgeGt1
                        i_CapAge1 = 15 '年滿15歲以上
                        i_CapAge2 = -1'Convert.DBNull
                    Case cst_AgeGt2
                        i_CapAge1 = 16 '年滿16歲以上
                        i_CapAge2 = -1'Convert.DBNull
                    Case cst_AgeGt3
                        i_CapAge1 = 20 '年滿20歲以上
                        i_CapAge2 = -1 'Convert.DBNull
                End Select
            Case Else
                '有上限，年滿15歲~
                i_CapAge1 = 15
                i_CapAge2 = Val(Age2Val) 'Val(txtAge2.Text)
        End Select
        'dr("CapAge1") = If(i_CapAge1 > 0, i_CapAge1, Convert.DBNull)
        'dr("CapAge2") = If(i_CapAge2 > 0, i_CapAge2, Convert.DBNull)
        If iType = 1 Then Return i_CapAge1
        If iType = 2 Then Return i_CapAge2
        Return -1
    End Function

    ''' <summary>
    ''' 兵役改複選功能(為配合舊資料需限制存取條件)
    ''' </summary>
    ''' <param name="oCapMilitary"></param>
    ''' <returns></returns>
    Public Shared Function Get_CapMilitaryVal(ByRef oCapMilitary As CheckBoxList) As String
        'dr("CapMilitary") = If(Solder.SelectedIndex = -1, Convert.DBNull, Solder.SelectedValue)
        Dim CapMilitaryVal As String = ""
        'CapMilitaryVal = ""
        For Each item As ListItem In oCapMilitary.Items
            If item.Selected Then
                Select Case item.Value
                    Case "00" '全選
                        CapMilitaryVal = item.Value
                        Exit For
                    Case Else
                        '04010302 '等同全選
                        CapMilitaryVal += item.Value
                End Select
            End If
        Next
        Dim v_CapMilitary As String = ""
        If CapMilitaryVal = "04010302" OrElse CapMilitaryVal.Length > 6 Then
            '00:等同全選
            v_CapMilitary = "00"
        Else
            '00:等同全選
            v_CapMilitary = If(CapMilitaryVal <> "", CapMilitaryVal, "00") '2013年應該不太可能有此值
        End If
        'dr("CapMilitary") = If(v_CapMilitary <> "", v_CapMilitary, Convert.DBNull)
        Return v_CapMilitary
    End Function

    ''' <summary>正式儲存 PLAN_PLANINFO</summary>
    ''' <param name="iNum">iNum: 1是正式 '2是草稿</param>
    Sub SAVE_PLAN_PLANINFO(ByVal iNum As Integer)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'iNum: 1是正式 '2是草稿
        Dim sql As String = ""
        Dim iSeqNO As Integer = 0
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing

        If Not (iPlanKind = 1) Then
            i_CostMode = GetCostMode(0)
            If i_CostMode = 0 Then Return 'Exit Sub '異常離開
        End If

        HidOrgID.Value = TIMS.Get_OrgID(sm.UserInfo.RID, objconn)
        If ComidValue.Value = "" Then ComidValue.Value = TIMS.ClearSQM(TIMS.Get_ComIDNOforOrgID(HidOrgID.Value, objconn))
        If ComidValue.Value = "" Then Return 'Exit Sub '至少要生個統編
        'i_PlanID = 0 'Val(sm.UserInfo.PlanID) 's_ComIDNO = "" 'ComidValue.Value 'i_SeqNO = 0

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim i_PlanID As Integer = sm.UserInfo.PlanID
        If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 AndAlso sm.UserInfo.LID = 0 Then
            Dim s_PLANID As String = TIMS.GET_RIDPLANID(Nothing, objconn, RIDValue.Value)
            If s_PLANID <> "" AndAlso s_PLANID <> Convert.ToString(i_PlanID) Then i_PlanID = Val(s_PLANID)
        End If

        Dim dr As DataRow = Nothing
        Dim flag_insertNew1 As Boolean = False 'true:新增一筆／false:修改一筆資料
        If (rqPlanID = "" OrElse gflag_ccopy) Then flag_insertNew1 = True
        Dim s_RedirectUrlOne As String = ""

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                If flag_insertNew1 Then
                    iSeqNO = GetMaxSeqNum(Trans, i_PlanID, ComidValue.Value) 'sm.UserInfo.PlanID
                    '準備儲存資料
                    sql = "SELECT * FROM PLAN_PLANINFO WHERE 1<>1"
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("PlanID") = i_PlanID 'sm.UserInfo.PlanID
                    dr("ComIDNO") = ComidValue.Value
                    dr("SeqNO") = iSeqNO
                    '預防新增時選擇草稿儲存
                    '導致因為停留員畫面，再儲存時會第二次重複儲存。
                    Org.Disabled = True
                    PlanID_value = i_PlanID 'Val(sm.UserInfo.PlanID)
                    ComIDNO_value = ComidValue.Value
                    SeqNO_value = iSeqNO
                Else
                    iSeqNO = rqSeqNO
                    sql = "SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rqPlanID & "' and ComIDNO='" & rqComIDNO & "' and SeqNO='" & rqSeqNO & "'"
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    dr = dt.Rows(0)
                    PlanID_value = Val(rqPlanID)
                    ComIDNO_value = rqComIDNO 'ComidValue.Value
                    SeqNO_value = Val(rqSeqNO)
                End If

                dr("RID") = RIDValue.Value
                dr("PlanYear") = Label3.Text
                dr("TPlanID") = sTPlanID
                dr("TMID") = If(trainValue.Value <> "", trainValue.Value, Convert.DBNull)
                dr("CJOB_UNKEY") = If(cjobValue.Value <> "", cjobValue.Value, Convert.DBNull)
                dr("PlanCause") = If(PlanCause.Text <> "", PlanCause.Text, Convert.DBNull)
                dr("PurScience") = If(PurScience.Text <> "", PurScience.Text, Convert.DBNull)
                dr("PurTech") = If(PurTech.Text <> "", PurTech.Text, Convert.DBNull)
                dr("PurMoral") = If(PurMoral.Text <> "", PurMoral.Text, Convert.DBNull)
                Dim v_Degree As String = TIMS.GetListValue(Degree)
                dr("CapDegree") = If(v_Degree <> "", v_Degree, "00")

                Dim i_CapAge1 As Integer = Get_CapAge_Val(rdoAge2, HidAgeType.Value, txtAge2.Text, 1)
                Dim i_CapAge2 As Integer = Get_CapAge_Val(rdoAge2, HidAgeType.Value, txtAge2.Text, 2)
                dr("CapAge1") = If(i_CapAge1 > 0, i_CapAge1, Convert.DBNull)
                dr("CapAge2") = If(i_CapAge2 > 0, i_CapAge2, Convert.DBNull)

                '性別-不區分了
                dr("CapSex") = "0"

                '兵役改複選功能(為配合舊資料需限制存取條件)
                Dim v_CapMilitary As String = Get_CapMilitaryVal(CapMilitary)
                dr("CapMilitary") = If(v_CapMilitary <> "", v_CapMilitary, Convert.DBNull)

                TRNUNITNAME.Text = TIMS.ClearSQM(TRNUNITNAME.Text)
                'TRNUNITCHO-委訓單位類型
                '1:政府機關/ 2:公民營事業機構/ 3:學校/ 4:團體/ 9:其他(請說明)
                Dim v_TRNUNITCHO As String = TIMS.GetListValue(TRNUNITCHO)
                TRNUNITTYPE.Text = TIMS.ClearSQM(TRNUNITTYPE.Text)
                TRNUNITEE.Text = TIMS.ClearSQM(TRNUNITEE.Text)
                dr("TRNUNITNAME") = If(TRNUNITNAME.Text <> "", TRNUNITNAME.Text, Convert.DBNull)
                'TRNUNITCHO-委訓單位類型
                '1:政府機關/ 2:公民營事業機構/ 3:學校/ 4:團體/ 9:其他(請說明)
                dr("TRNUNITCHO") = If(v_TRNUNITCHO <> "", v_TRNUNITCHO, Convert.DBNull)
                dr("TRNUNITTYPE") = If(TRNUNITTYPE.Text <> "", TRNUNITTYPE.Text, Convert.DBNull)
                dr("TRNUNITEE") = If(TRNUNITEE.Text <> "", TRNUNITEE.Text, Convert.DBNull)

                Other1.Text = TIMS.ClearSQM(Other1.Text)
                Other2.Text = TIMS.ClearSQM(Other2.Text)
                Other3.Text = TIMS.ClearSQM(Other3.Text)
                TMScience.Text = TIMS.ClearSQM(TMScience.Text)
                TMTech.Text = TIMS.ClearSQM(TMTech.Text)
                dr("CapOther1") = If(Other1.Text <> "", Other1.Text, Convert.DBNull)
                dr("CapOther2") = If(Other2.Text <> "", Other2.Text, Convert.DBNull)
                dr("CapOther3") = If(Other3.Text <> "", Other3.Text, Convert.DBNull)
                dr("TMScience") = If(TMScience.Text <> "", TMScience.Text, Convert.DBNull)
                dr("TMTech") = If(TMTech.Text <> "", TMTech.Text, Convert.DBNull)

                '持推介單報參訓之適用條件
                Dim v_GetTrain1 As String = TIMS.GetListValue(GetTrain1)
                If tr_GetTrain1.Visible = False AndAlso v_GetTrain1 = "" Then v_GetTrain1 = "3" '3.不適用推介機制
                dr("GetTrain1") = If(v_GetTrain1 <> "", v_GetTrain1, Convert.DBNull)
                '自行報名參訓者錄訓規定
                dr("GetTrain2") = If(GetTrain2.Text <> "", GetTrain2.Text, Convert.DBNull)

                Dim v_GetTrain3 As String = TIMS.GetCblValue(GetTrain3) '甄試方式
                dr("GetTrain3") = If(v_GetTrain3 <> "", v_GetTrain3, Convert.DBNull)
                dr("GetTrain3Other") = If(GetTrain3Other.Text <> "", GetTrain3Other.Text, Convert.DBNull)
                '其他
                Dim v_GetTrain4 As String = TIMS.GetCblValue(GetTrain4)
                dr("GetTrain4") = If(v_GetTrain4 <> "", v_GetTrain4, Convert.DBNull)
                dr("GetTrain4Other") = If(GetTrain4Other.Text <> "", GetTrain4Other.Text, Convert.DBNull)
                '課程編配「一般學科」為必填欄位
                dr("GenSciHours") = If(GenSciHours.Text <> "", Val(GenSciHours.Text), Convert.DBNull)
                '課程編配「專業學科」為必填欄位
                dr("ProSciHours") = If(ProSciHours.Text <> "", Val(ProSciHours.Text), Convert.DBNull)
                '課程編配「術科」為必填欄位
                dr("ProTechHours") = If(ProTechHours.Text <> "", Val(ProTechHours.Text), Convert.DBNull)

                OtherHours.Text = TIMS.ClearSQM(OtherHours.Text)
                dr("OtherHours") = If(OtherHours.Text <> "", Val(OtherHours.Text), Convert.DBNull)
                TotalHours.Text = TIMS.ClearSQM(TotalHours.Text)
                dr("TotalHours") = If(TotalHours.Text <> "", Val(TotalHours.Text), Convert.DBNull)
                DefGovCost.Text = TIMS.ClearSQM(DefGovCost.Text)
                dr("DefGovCost") = If(DefGovCost.Text <> "", Val(DefGovCost.Text), Convert.DBNull)
                DefUnitCost.Text = TIMS.ClearSQM(DefUnitCost.Text)
                dr("DefUnitCost") = If(DefUnitCost.Text <> "", Val(DefUnitCost.Text), Convert.DBNull)
                DefStdCost.Text = TIMS.ClearSQM(DefStdCost.Text)
                dr("DefStdCost") = If(DefStdCost.Text <> "", Val(DefStdCost.Text), Convert.DBNull)

                ClassName.Text = Replace(TIMS.ClearSQM(ClassName.Text), "&", "＆")
                dr("ClassName") = If(ClassName.Text <> "", ClassName.Text, Convert.DBNull)
                dr("Class_Unit") = If(Class_Unit.Value <> "", Class_Unit.Value, Convert.DBNull)
                Dim v_rblADVANCE As String = TIMS.GetListValue(rblADVANCE) '訓練課程類型
                dr("ADVANCE") = If(v_rblADVANCE <> "", v_rblADVANCE, Convert.DBNull) '訓練課程類型
                dr("TNum") = If(TNum.Text <> "", TNum.Text, Convert.DBNull)
                dr("THours") = If(THours.Text <> "", THours.Text, Convert.DBNull)
                dr("STDate") = If(STDate.Text <> "", STDate.Text, Convert.DBNull)
                dr("FDDate") = If(FDDate.Text <> "", FDDate.Text, Convert.DBNull)

                dr("TAddressZip") = If(TAddressZip.Value <> "", TAddressZip.Value, Convert.DBNull)
                hidTAddressZIP6W.Value = TIMS.GetZIPCODE6W(TAddressZip.Value, TAddressZIPB3.Value)
                dr("TAddressZIP6W") = If(hidTAddressZIP6W.Value <> "", hidTAddressZIP6W.Value, Convert.DBNull) '2009-05-20 fix
                dr("TAddress") = If(TAddress.Text <> "", TAddress.Text, Convert.DBNull)

                '報名地點/甄試地點
                dr("EAddressZip") = If(EAddressZip.Value <> "", EAddressZip.Value, Convert.DBNull)
                hidEAddressZIP6W.Value = TIMS.GetZIPCODE6W(EAddressZip.Value, EAddressZIPB3.Value)
                dr("EAddressZIP6W") = If(hidEAddressZIP6W.Value <> "", hidEAddressZIP6W.Value, Convert.DBNull) '2009-05-20 fix
                dr("EAddress") = If(EAddress.Text <> "", EAddress.Text, Convert.DBNull)

                '報名開始日期
                Dim s_SEnterDate As String = TIMS.GET_DateHM(SEnterDate, HR1, MM1)
                dr("SEnterDate") = If(s_SEnterDate <> "", CDate(s_SEnterDate), Convert.DBNull)
                '報名結束日期
                Dim s_FEnterDate As String = TIMS.GET_DateHM(FEnterDate, HR2, MM2)
                dr("FEnterDate") = If(s_FEnterDate <> "", CDate(s_FEnterDate), Convert.DBNull)

                '甄試日期
                ExamDate.Text = TIMS.Cdate3(TIMS.ClearSQM(ExamDate.Text))
                '20100329 andy add 甄試日期(時段)'01-全天'02-上午'03-下午
                Dim v_ExamPeriod As String = TIMS.GetListValue(ExamPeriod)
                '區域據點-調整甄試日期欄位相關卡控
                If (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1) AndAlso v_GetTrain3 = "" Then
                    ExamDate.Text = ""
                    v_ExamPeriod = ""
                    Common.SetListItem(ExamPeriod, v_ExamPeriod)
                End If
                '甄試日期/時間
                'If TIMS.GFG_OJT_25050801_NoUse_ExamDateTime Then dr("ExamDate") = If(ExamDate.Text <> "", TIMS.Cdate2(ExamDate.Text), Convert.DBNull)
                Dim s_ExamDate As String = TIMS.GET_DateHM(ExamDate, HR6, MM6)
                dr("ExamDate") = If(s_ExamDate <> "", CDate(s_ExamDate), Convert.DBNull)
                dr("ExamPeriod") = If(ExamDate.Text <> "" AndAlso v_ExamPeriod <> "", v_ExamPeriod, Convert.DBNull)
                '報到日期 
                dr("CheckInDate") = If(CheckInDate.Text <> "", TIMS.Cdate2(CheckInDate.Text), Convert.DBNull)
                '聯絡人姓名
                dr("ContactName") = If(ContactName.Text <> "", ContactName.Text, Convert.DBNull)
                '聯絡人電話
                dr("ContactPhone") = If(ContactPhone.Text <> "", ContactPhone.Text, Convert.DBNull)
                '聯絡人電子郵件
                dr("ContactEmail") = If(ContactEmail.Text <> "", ContactEmail.Text, Convert.DBNull)
                '直屬主管電子郵件
                dr("MasterEmail") = If(MasterEmail.Text <> "", MasterEmail.Text, Convert.DBNull)
                '訓字保保險證號。
                If twiACTNO.Text <> "" Then twiACTNO.Text = TIMS.ChangeIDNO(twiACTNO.Text)
                dr("twiACTNO") = If(twiACTNO.Text <> "", twiACTNO.Text, Convert.DBNull)

                CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
                dr("CyclType") = If(CyclType.Text <> "", CyclType.Text, Convert.DBNull)

                dr("ClassCount") = If(ClassCount.Text <> "", ClassCount.Text, Convert.DBNull)
                If iPlanKind = 1 Then
                    dr("TotalCost") = If(TotalCost1.Text <> "", TotalCost1.Text, Convert.DBNull)
                Else
                    Select Case i_CostMode
                        Case 2
                            dr("TotalCost") = If(TotalCost2.Text <> "", TotalCost2.Text, Convert.DBNull)
                        Case 3
                            dr("TotalCost") = If(TotalCost3.Text <> "", TotalCost3.Text, Convert.DBNull)
                        Case 4
                            dr("TotalCost") = If(TotalCost4.Text <> "", TotalCost4.Text, Convert.DBNull)
                    End Select
                End If

                'AdmGrant 'AdmPercent '已採用Plan_CostItem來代替 行政管理費用尚未代替
                If Not Session("AdmGrant") Is Nothing Then
                    If Session("AdmGrant") <> "" Then dr("AdmPercent") = Session("AdmGrant")
                End If

                If Not Session("TaxGrant") Is Nothing Then '營業稅費用使用舊方案
                    If Session("TaxGrant") <> "" Then dr("TaxPercent") = Session("TaxGrant")
                End If

                dr("Note") = If(Note.Text <> "", Note.Text, Convert.DBNull)
                dr("ESiteMsg") = If(ESiteMsg.Text <> "", ESiteMsg.Text, Convert.DBNull)
                'iNum: 1是正式 '2是草稿
                If iNum = 1 AndAlso dr("AppliedDate").ToString() = "" Then dr("AppliedDate") = Now.Date
                dr("AppliedOrigin") = 1
                dr("PlanEMail") = If(EMail.Text = "", Convert.DBNull, EMail.Text)

                '班級英文名稱
                dr("CLASSENGNAME") = If(ClassEngName.Text <> "", Trim(ClassEngName.Text), Convert.DBNull)
                '訓練時段'取得鍵值-訓練時段
                Dim v_TPeriodList As String = TIMS.GetListValue(TPeriodList)
                dr("TPERIOD") = If(v_TPeriodList <> "", v_TPeriodList, Convert.DBNull)
                '訓練時段2
                dr("NOTE3") = If(TB_NOTE3.Text <> "", Trim(TB_NOTE3.Text), Convert.DBNull)
                '「訓練期限」
                Dim v_TDeadline_List As String = TIMS.GetListValue(TDeadline_List)
                dr("TDEADLINE") = If(v_TDeadline_List <> "", v_TDeadline_List, Convert.DBNull)
                '導師名稱
                dr("CTName") = TIMS.Get_CTNAME1(CTName.Text) 'CTName.Text
                '是否為輔導考照班
                dr("COACHING") = If(RBCOACHING_Y.Checked, "Y", If(RBCOACHING_N.Checked, "N", Convert.DBNull))
                '檢定職類代碼1/2/3 與考試級別1/2/3
                'txtGP1', 'txtXM1', 'txtLV1', 'EXAM1val', 'EXLV1val' 
                If RBCOACHING_Y.Checked Then
                    dr("EXAMIDS1") = If(EXAM1val.Value <> "", EXAM1val.Value, Convert.DBNull)
                    dr("EXAMIDS2") = If(EXAM2val.Value <> "", EXAM2val.Value, Convert.DBNull)
                    dr("EXAMIDS3") = If(EXAM3val.Value <> "", EXAM3val.Value, Convert.DBNull)
                    dr("EXAMLVID1") = If(EXLV1val.Value <> "", EXLV1val.Value, Convert.DBNull)
                    dr("EXAMLVID2") = If(EXLV2val.Value <> "", EXLV2val.Value, Convert.DBNull)
                    dr("EXAMLVID3") = If(EXLV3val.Value <> "", EXLV3val.Value, Convert.DBNull)
                Else
                    dr("EXAMIDS1") = Convert.DBNull
                    dr("EXAMIDS2") = Convert.DBNull
                    dr("EXAMIDS3") = Convert.DBNull
                    dr("EXAMLVID1") = Convert.DBNull
                    dr("EXAMLVID2") = Convert.DBNull
                    dr("EXAMLVID3") = Convert.DBNull
                End If

                Dim flag_can_DEL_VERRECORD As Boolean = False '確認是否要刪除-PLAN_VERRECORD
                Dim s_AppliedResult As String = Convert.ToString(dr("AppliedResult"))
                Dim s_TransFlag As String = Convert.ToString(dr("TransFlag"))
                Dim s_IsApprPaper As String = Convert.ToString(dr("IsApprPaper"))
                'iNum: 1是正式 '2是草稿
                If iNum = 1 Then '計畫為正式
                    If flag_insertNew1 Then
                        '新增的狀況 '分署內訓計畫為審核通過
                        s_AppliedResult = If(iPlanKind = 1, "Y", "")
                    Else
                        '修改的狀況
                        If iPlanKind = 1 Then
                            s_AppliedResult = "Y" '分署內訓計畫為審核通過
                        Else
                            Select Case s_AppliedResult'dr("AppliedResult").ToString()
                                Case "Y", "O"
                                Case Else
                                    s_AppliedResult = ""
                                    flag_can_DEL_VERRECORD = True
                            End Select
                        End If
                    End If
                    'dr("AppliedResult") = If(s_AppliedResult <> "", s_AppliedResult, Convert.DBNull)
                    '空白的話自動轉為N
                    s_TransFlag = If(s_TransFlag <> "", s_TransFlag, "N")
                    s_IsApprPaper = "Y" '(轉正式)
                Else
                    '草稿儲存 清空  AppliedResult
                    s_AppliedResult = "" 'Convert.DBNull
                End If
                '審核通過
                dr("AppliedResult") = If(s_AppliedResult <> "", s_AppliedResult, Convert.DBNull)
                '空白的話自動轉為N
                dr("TransFlag") = If(s_TransFlag <> "", s_TransFlag, Convert.DBNull)
                dr("IsApprPaper") = If(s_IsApprPaper <> "", s_IsApprPaper, Convert.DBNull) '"Y" '(轉正式)
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
                'SAVE PLAN_PLANINFO
                DbAccess.UpdateDataTable(dt, da, Trans)

                'i_PlanID = Val(sm.UserInfo.PlanID) 's_ComIDNO = ComidValue.Value 'i_SeqNO = iSeqNO
                'Dim v_CBLKID22 As String = TIMS.GetListValue(CBLKID22)
                Dim cvCBLKID60 As String = TIMS.GetCblValue(CBLKID60)
                Dim s_cmdArg As String = ""
                s_cmdArg &= "&PlanID=" & PlanID_value 'hidPlanID.Value
                s_cmdArg &= "&ComIDNO=" & ComIDNO_value 'hidComIDNO.Value
                s_cmdArg &= "&SeqNo=" & SeqNO_value 'hidSeqNO.Value
                '2019年啟用 work2019x01:2019 政府政策性產業
                s_cmdArg &= "&KID20=" & GET_KID20_VAL()
                s_cmdArg &= "&KID25=" & GET_KID25_VAL()
                s_cmdArg &= "&KID60=" & cvCBLKID60
                's_cmdArg &= "&KID22=" & v_CBLKID22
                Call SAVE_PLAN_DEPOT(s_cmdArg, Trans)

                If flag_can_DEL_VERRECORD Then
                    '要刪除 - PLAN_VERRECORD
                    'Dim v_PlanID As String = TIMS.GetMyValue2(pParms, "PlanID")
                    'Dim v_ComidValue As String = TIMS.GetMyValue2(pParms, "ComidValue")
                    'Dim v_SeqNo As String = TIMS.GetMyValue2(pParms, "SeqNo") ' sm.UserInfo.PlanID)
                    Dim pParmsD1 As New Hashtable From {{"PlanID", i_PlanID}, {"ComidValue", TIMS.ClearSQM(ComidValue.Value)}, {"SeqNo", TIMS.ClearSQM(rqSeqNO)}}
                    Call DEL_PLANVERRECORD(pParmsD1, Trans)
                End If

                Dim dtTemp As DataTable = Nothing
                '儲存計畫可能失敗 'update 計畫經費項目檔(Plan_CostItem)
                Call SAVE_PLAN_COSTITEM_OTH(dtTemp, dt, da, Trans, iSeqNO, i_PlanID)
                '儲存計畫可能失敗  '更新訓練課程內容簡介(Plan_TrainDesc)
                Call SAVE_PLAN_TRAINDESC_OTH(dtTemp, dt, da, Trans, iSeqNO, i_PlanID)
                DbAccess.CommitTrans(Trans)
                'Call TIMS.CloseDbConn(Conn)

            Catch ex As Exception
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rqPlanID, rqComIDNO, rqSeqNO) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)

                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Session(Hid_TrainDesc_GUID1.Value) = Nothing
                Session(Hid_CostItem_GUID1.Value) = Nothing
                Session("AdmGrant") = Nothing '行政管理費百分比
                Session("TaxGrant") = Nothing '營業稅費用百分比
                If iNum = 1 Then
                    If rqPlanID = "" Then
                        Common.MessageBox(Page, String.Concat("計畫申請失敗!!", ex.Message))
                    Else
                        Common.MessageBox(Page, String.Concat("計畫儲存失敗!!", ex.Message))
                    End If
                Else
                    Common.MessageBox(Page, String.Concat("草稿儲存失敗!!", ex.Message))
                End If
                If (LayerState.Value = "") Then LayerState.Value = "1"
                Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
                Page.RegisterStartupScript("_onload", s_js1)
                Return 'Throw ex
            End Try
            Call TIMS.CloseDbConn(TransConn)
        End Using

        '儲存計畫可能失敗 '更新 CLASS_CLASSINFO.FENTERDATE2
        Call SAVE_CLASSINFO_OTH(objconn, iSeqNO, i_PlanID)
        '更新 CLASS_CLASSINFO
        Call UPDATE_CLASS_CLASSINFO(PlanID_value, ComIDNO_value, SeqNO_value, objconn)
        '專長能力標籤-ABILITY
        Call SAVE_PLAN_ABILITY()

        Session(Hid_TrainDesc_GUID1.Value) = Nothing
        Session(Hid_CostItem_GUID1.Value) = Nothing
        Session("AdmGrant") = Nothing '行政管理費百分比
        Session("TaxGrant") = Nothing '營業稅費用百分比
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        If rqPlanID = "" Then
            '新增時儲存
            If iNum = 1 Then
                '正式送出
                'Common.RespWrite(Me, "<script>alert('計畫申請成功!!');location.href='TC_03_001.aspx?ID=" & rqMID & "'</script>")
                Common.MessageBox(Page, "計畫申請成功!!")
                s_RedirectUrlOne = String.Concat("../03/TC_03_001.aspx?ID=", rqMID)
            Else
                '草稿儲存
                'Common.RespWrite(Me, "<script>alert('草稿儲存成功!!');")
                'Common.RespWrite(Me, "location.href='../03/TC_03_001.aspx?ID=" & rqMID & "&PlanID=" & sm.UserInfo.PlanID & "&ComIDNO=" & ComidValue.Value & "&SeqNo=" & iSeqNO & "'</script>")
                Dim MyValue1 As String = ""
                MyValue1 &= "&PlanID=" & i_PlanID ' sm.UserInfo.PlanID
                MyValue1 &= "&ComIDNO=" & ComidValue.Value
                MyValue1 &= "&SeqNo=" & iSeqNO
                Common.MessageBox(Page, "草稿儲存成功!!")
                s_RedirectUrlOne = String.Concat("../03/TC_03_001.aspx?ID=", rqMID, MyValue1)
            End If
        Else
            '查詢修改 或複制
            If iNum = 1 Then
                '正式送出
                'Common.RespWrite(Me, "<script>alert('計畫儲存成功!!');</script>")
                If gflag_ccopy Then
                    '班級複製作業 回查詢
                    'Common.RespWrite(Me, "<script>location.href='../03/TC_03_002.aspx?ID=" & rqMID & "'</script>")
                    Common.MessageBox(Page, "計畫儲存成功!!")
                    s_RedirectUrlOne = String.Concat("../03/TC_03_002.aspx?ID=", rqMID)
                Else
                    '班級查詢作業 回查詢
                    'Common.RespWrite(Me, "<script>location.href='../02/TC_02_001.aspx?ID=" & rqMID & "'</script>")
                    Common.MessageBox(Page, "計畫儲存成功!!")
                    s_RedirectUrlOne = String.Concat("../02/TC_02_001.aspx?ID=", rqMID)
                End If
            Else
                '草稿儲存
                'Common.RespWrite(Me, "<script>alert('草稿儲存成功!!');") '</script>
                'Common.RespWrite(Me, "location.href='../03/TC_03_001.aspx?ID=" & rqMID & "&PlanID=" & sm.UserInfo.PlanID & "&ComIDNO=" & ComidValue.Value & "&SeqNo=" & iSeqNO & "'</script>")                   
                Dim MyValue1 As String = ""
                MyValue1 &= "&PlanID=" & i_PlanID ' sm.UserInfo.PlanID
                MyValue1 &= "&ComIDNO=" & ComidValue.Value
                MyValue1 &= "&SeqNo=" & iSeqNO
                Common.MessageBox(Page, "草稿儲存成功!!")
                s_RedirectUrlOne = String.Concat("../03/TC_03_001.aspx?ID=", rqMID, MyValue1)
            End If
        End If


        If s_RedirectUrlOne <> "" Then TIMS.Utl_Redirect(Me, objconn, s_RedirectUrlOne)
    End Sub

    ''' <summary>
    ''' 更新 CLASS_CLASSINFO
    ''' </summary>
    ''' <param name="PlanID"></param>
    ''' <param name="ComIDNO"></param>
    ''' <param name="SeqNo"></param>
    ''' <param name="conn"></param>
    Sub UPDATE_CLASS_CLASSINFO(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String, ByRef conn As SqlConnection)
        ', ByRef Trans As SqlTransaction
        If PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "" Then Return 'rst Exit Sub

        Dim parms As New Hashtable From {{"PlanID", PlanID}, {"ComIDNO", ComIDNO}, {"SeqNo", SeqNo}}
        Dim sql As String = "SELECT OCID FROM CLASS_CLASSINFO WHERE PlanID=@PlanID and ComIDNO=@ComIDNO and SeqNo=@SeqNo"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, conn, parms)
        If TIMS.dtNODATA(dt) Then Exit Sub
        Dim dr As DataRow = dt.Rows(0)
        Dim s_OCID As String = Convert.ToString(dr("OCID"))

        'Dim sql As String = ""
        Dim U_sql As String = ""
        U_sql &= " UPDATE CLASS_CLASSINFO" & vbCrLf
        U_sql &= " SET TNUM=@TNUM ,THOURS=@THOURS ,TMID=@TMID ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        U_sql &= " WHERE OCID=@OCID AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo" & vbCrLf
        Dim U_PARMS As New Hashtable From {
            {"TNUM", If(TNum.Text <> "", Val(TNum.Text), Convert.DBNull)},
            {"THOURS", If(THours.Text <> "", Val(THours.Text), Convert.DBNull)},
            {"TMID", If(trainValue.Value <> "", Val(trainValue.Value), Convert.DBNull)},
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"OCID", s_OCID},
            {"PlanID", PlanID},
            {"ComIDNO", ComIDNO},
            {"SeqNo", SeqNo}
        }
        DbAccess.ExecuteNonQuery(U_sql, conn, U_PARMS)
    End Sub

    '檢查session 資料是否正常，若不正常停止儲存。預設為 false:異常 true:正常
    Function Chk_ssCOSTITEM(ByVal SeqNO As String, ByVal PCID As Integer, i_PlanID As Integer) As Boolean
        Dim rst As Boolean = False 'false:異常 (新增一筆)  true:正常 (覆蓋)

        If Not Session(Hid_CostItem_GUID1.Value) Is Nothing Then
            Dim dtTemp As DataTable = Session(Hid_CostItem_GUID1.Value)
            Dim ff As String = "PCID=" & PCID
            For Each dr As DataRow In dtTemp.Select(ff, Nothing, DataViewRowState.CurrentRows)
                '檢查所有計畫，是否有異常
                If rqPlanID = "" OrElse gflag_ccopy Then
                    '新增'或COPY
                    If Convert.ToString(dr("PlanID")) <> Convert.ToString(i_PlanID) OrElse
                        Convert.ToString(dr("ComIDNO")) <> ComidValue.Value OrElse
                        Convert.ToString(dr("SeqNO")) <> SeqNO Then
                        '任何異常 離開
                        Return rst ' False
                    End If
                Else
                    '非:新增'或COPY '是:修改
                    If Convert.ToString(dr("PlanID")) <> rqPlanID OrElse
                        Convert.ToString(dr("ComIDNO")) <> rqComIDNO OrElse
                        Convert.ToString(dr("SeqNO")) <> rqSeqNO Then
                        '任何異常 離開
                        Return rst ' False
                    End If
                End If
            Next
        End If
        rst = True
        Return rst
    End Function

    '檢查session 資料是否正常，若不正常停止儲存。預設為 false:異常 true:正常
    Function Chk_ssTRAINDESC(ByVal SeqNO As String, ByVal PTDID As Integer, i_PlanID As Integer) As Boolean
        Dim rst As Boolean = False 'false:異常 (新增一筆)  true:正常 (覆蓋)
        If Not Session(Hid_TrainDesc_GUID1.Value) Is Nothing Then
            Dim dtTemp As DataTable = Session(Hid_TrainDesc_GUID1.Value)
            Dim ff As String = "PTDID=" & PTDID
            For Each dr As DataRow In dtTemp.Select(ff, Nothing, DataViewRowState.CurrentRows)
                '檢查所有計畫，是否有異常
                If rqPlanID = "" OrElse gflag_ccopy Then
                    '新增'或COPY
                    If Convert.ToString(dr("PlanID")) <> Convert.ToString(i_PlanID) OrElse
                        Convert.ToString(dr("ComIDNO")) <> ComidValue.Value OrElse
                        Convert.ToString(dr("SeqNO")) <> SeqNO Then
                        '任何異常 離開
                        Return rst ' False
                    End If
                Else
                    '非:新增'或COPY '是:修改
                    If Convert.ToString(dr("PlanID")) <> rqPlanID OrElse
                        Convert.ToString(dr("ComIDNO")) <> rqComIDNO OrElse
                        Convert.ToString(dr("SeqNO")) <> rqSeqNO Then
                        '任何異常 離開
                        Return rst ' False
                    End If
                End If
            Next
        End If
        rst = True
        Return rst
    End Function

    '儲存計畫可能失敗 'update 計畫經費項目檔(Plan_CostItem)
    Sub SAVE_PLAN_COSTITEM_OTH(ByRef dtTemp As DataTable, ByRef dt As DataTable, ByRef da As SqlDataAdapter, ByRef Trans As SqlTransaction, ByRef SeqNO As String, i_PlanID As Integer)
        Dim sql As String = ""
        If Session(Hid_CostItem_GUID1.Value) Is Nothing Then Return 'Exit Sub
        'update 計畫經費項目檔(Plan_CostItem)
        If Not Session(Hid_CostItem_GUID1.Value) Is Nothing Then
            dtTemp = Session(Hid_CostItem_GUID1.Value)
            sql = "SELECT * FROM PLAN_COSTITEM WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, Trans)
            For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                If dr("PCID") <= 0 Then
                    dr("PCID") = DbAccess.GetNewId(Trans, "PLAN_COSTITEM_PCID_SEQ,PLAN_COSTITEM,PCID")
                Else
                    If gflag_ccopy Then
                        dr("PCID") = DbAccess.GetNewId(Trans, "PLAN_COSTITEM_PCID_SEQ,PLAN_COSTITEM,PCID")
                    Else
                        '異常新增，正常覆蓋
                        If Not Chk_ssCOSTITEM(SeqNO, Val(dr("PCID")), i_PlanID) Then dr("PCID") = DbAccess.GetNewId(Trans, "PLAN_COSTITEM_PCID_SEQ,PLAN_COSTITEM,PCID")
                    End If
                End If
                If rqPlanID = "" OrElse gflag_ccopy Then
                    '新增'或COPY
                    dr("PlanID") = i_PlanID ' sm.UserInfo.PlanID
                    dr("ComIDNO") = ComidValue.Value
                    dr("SeqNO") = SeqNO
                Else
                    '非:新增'或COPY '是:修改
                    dr("PlanID") = rqPlanID
                    dr("ComIDNO") = rqComIDNO
                    dr("SeqNO") = rqSeqNO
                End If
            Next
            dt = dtTemp.Copy
            DbAccess.UpdateDataTable(dt, da, Trans)
        End If
    End Sub

    '儲存計畫可能失敗  '更新訓練課程內容簡介(Plan_TrainDesc)
    Sub SAVE_PLAN_TRAINDESC_OTH(ByRef dtTemp As DataTable, ByRef dt As DataTable, ByRef da As SqlDataAdapter, ByRef Trans As SqlTransaction, ByRef SeqNO As String, i_PlanID As Integer)
        Dim sql As String = ""
        If Session(Hid_TrainDesc_GUID1.Value) Is Nothing Then Return 'Exit Sub

        '更新訓練課程內容簡介(Plan_TrainDesc)
        If Not Session(Hid_TrainDesc_GUID1.Value) Is Nothing Then
            dtTemp = Session(Hid_TrainDesc_GUID1.Value)
            sql = "SELECT * FROM PLAN_TRAINDESC WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, Trans)
            For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                If dr("PTDID") <= 0 Then
                    dr("PTDID") = DbAccess.GetNewId(Trans, "PLAN_TRAINDESC_PTDID_SEQ,PLAN_TRAINDESC,PTDID")
                Else
                    If gflag_ccopy Then
                        dr("PTDID") = DbAccess.GetNewId(Trans, "PLAN_TRAINDESC_PTDID_SEQ,PLAN_TRAINDESC,PTDID")
                    Else
                        '異常新增，正常覆蓋
                        If Not Chk_ssTRAINDESC(SeqNO, Val(dr("PTDID")), i_PlanID) Then
                            dr("PTDID") = DbAccess.GetNewId(Trans, "PLAN_TRAINDESC_PTDID_SEQ,PLAN_TRAINDESC,PTDID")
                        End If
                    End If
                End If
                If rqPlanID = "" OrElse gflag_ccopy Then
                    '新增'或COPY
                    dr("PlanID") = i_PlanID 'sm.UserInfo.PlanID
                    dr("ComIDNO") = ComidValue.Value
                    dr("SeqNO") = SeqNO
                Else
                    '非:新增'或COPY '是:修改
                    dr("PlanID") = rqPlanID
                    dr("ComIDNO") = rqComIDNO
                    dr("SeqNO") = rqSeqNO
                End If
            Next
            dt = dtTemp.Copy
            DbAccess.UpdateDataTable(dt, da, Trans)
        End If
    End Sub

    ''' <summary>
    ''' '儲存計畫可能失敗 '更新 CLASS_CLASSINFO.FENTERDATE2 /UPDATE EXAMDATE
    ''' </summary>
    ''' <param name="oConn"></param>
    ''' <param name="SeqNO"></param>
    ''' <param name="i_PlanID"></param>
    Sub SAVE_CLASSINFO_OTH(oConn As SqlConnection, SeqNO As String, i_PlanID As Integer)
        Dim prPlanID As String = ""
        Dim prComIDNO As String = ""
        Dim prSeqNO As String = ""

        If rqPlanID = "" OrElse gflag_ccopy Then
            '新增'或COPY
            prPlanID = i_PlanID 'sm.UserInfo.PlanID
            prComIDNO = ComidValue.Value
            prSeqNO = SeqNO
        Else
            '非:新增'或COPY '是:修改
            prPlanID = rqPlanID
            prComIDNO = rqComIDNO
            prSeqNO = rqSeqNO
        End If
        If prPlanID = "" OrElse prComIDNO = "" OrElse prSeqNO = "" Then Return 'Exit Sub

        Dim sFENTERDATE As String = FEnterDate.Text
        Dim sEXAMDATE As String = ExamDate.Text
        Dim SS1 As String = ""
        TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
        Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, oConn)
        If sFENTERDATE2 <> "" Then
            Dim pFEnterDate2 As String = TIMS.Cdate3(sFENTERDATE2)
            Dim pHR5 As String = CDate(sFENTERDATE2).Hour
            Dim pMM5 As String = CDate(sFENTERDATE2).Minute
        End If

        Dim ck_OCID As String = ""
        Dim drP As DataRow = TIMS.GetPCSDate(prPlanID, prComIDNO, prSeqNO, oConn)
        If drP Is Nothing Then Return 'Exit Sub
        If Convert.ToString(drP("OCID")) <> "" Then ck_OCID = Convert.ToString(drP("OCID"))
        If ck_OCID = "" Then Return ' Exit Sub'(查無資料，離開)

        Call TIMS.OpenDbConn(oConn)
        Dim sql_C1 As String = " SELECT FENTERDATE2,FENTERDATE,EXAMDATE,RID,PLANID FROM dbo.CLASS_CLASSINFO WHERE OCID=@OCID" & vbCrLf
        Using sCmd As New SqlCommand(sql_C1, oConn)
            Using dt As New DataTable
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = ck_OCID
                    dt.Load(.ExecuteReader())  'edit，by:20181016 'dt = DbAccess.GetDataTable(sCmd.CommandText, oConn, sCmd.Parameters)  'edit，by:20181016
                End With
                If TIMS.dtNODATA(dt) Then Return '(查無課程資料，異常離開)
            End Using
        End Using

        Dim v_rblADVANCE As String = TIMS.GetListValue(rblADVANCE) '訓練課程類型
        'Dim dr As DataRow = dt.Rows(0) '01-全天'02-上午'03-下午
        Dim v_ExamPeriod As String = TIMS.GetListValue(ExamPeriod)
        '甄試日期/時間
        Dim s_ExamDate As String = TIMS.GET_DateHM(ExamDate, HR6, MM6)
        'If TIMS.GFG_OJT_25050801_NoUse_ExamDateTime Then s_ExamDate = TIMS.Cdate3(ExamDate.Text) 
        Dim u_sql2 As String = "UPDATE CLASS_CLASSINFO SET EXAMDATE=@EXAMDATE,EXAMPERIOD=@EXAMPERIOD,ADVANCE=@ADVANCE WHERE OCID=@OCID" & vbCrLf
        Using uCmd As New SqlCommand(u_sql2, oConn)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("EXAMDATE", SqlDbType.VarChar).Value = If(s_ExamDate <> "", s_ExamDate, Convert.DBNull)
                .Parameters.Add("EXAMPERIOD", SqlDbType.VarChar).Value = If(v_ExamPeriod <> "", v_ExamPeriod, Convert.DBNull)
                .Parameters.Add("ADVANCE", SqlDbType.VarChar).Value = If(v_rblADVANCE <> "", v_rblADVANCE, Convert.DBNull) '訓練課程類型
                .Parameters.Add("OCID", SqlDbType.VarChar).Value = ck_OCID
                .ExecuteNonQuery()
            End With
        End Using

        Dim u_sql As String = " UPDATE CLASS_CLASSINFO SET FENTERDATE2=dbo.FN_GET_FENTERDATE2B(FENTERDATE,EXAMDATE,RID,PLANID) WHERE OCID=@OCID"
        Using uCmd As New SqlCommand(u_sql, oConn)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("OCID", SqlDbType.VarChar).Value = ck_OCID
                .ExecuteNonQuery()
            End With
        End Using
        'Call CloseDbConn(conn) 'If dt.Rows.Count > 0 Then Rst = Convert.ToString(dt.Rows(0)("?"))
    End Sub

    '計價新增 暫存
    Sub AddCost(ByVal num As Integer)
        '1.成本加工費法(限定自辦使用) '2.每人每時單價計價法 '3.每人輔助單價計價法 '4.個人單價計價法

        '1.成本加工費法(限定自辦使用) 'OPrice * Itemage * ItemCost
        '2.每人每時單價計價法 'OPrice * Itemage * ItemCost
        '3.每人輔助單價計價法 'OPrice * Itemage 
        '4.個人單價計價法 'OPrice * Itemage 
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Try
            If Session(Hid_CostItem_GUID1.Value) Is Nothing Then
                Dim sql As String = "SELECT * FROM PLAN_COSTITEM where 1<>1"
                dt = DbAccess.GetDataTable(sql, objconn)
                dt.Columns("PCID").AutoIncrement = True
                dt.Columns("PCID").AutoIncrementSeed = -1
                dt.Columns("PCID").AutoIncrementStep = -1
            Else
                dt = Session(Hid_CostItem_GUID1.Value)
            End If

            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("CostMode") = num
            Select Case num
                Case 1
                    dr("CostID") = CostID.SelectedValue
                    dr("ItemOther") = If(ItemOther.Text <> "", ItemOther.Text, Convert.DBNull)
                    dr("OPrice") = OPrice.Text
                    dr("Itemage") = Itemage.Text
                    dr("ItemCost") = ItemCost.Text
                    dr("AdmFlag") = "N"
                    CostID.SelectedIndex = -1
                    ItemOther.Text = ""
                    OPrice.Text = ""
                    Itemage.Text = ""
                    ItemCost.Text = ""
                Case 2
                    dr("OPrice") = OPrice2.Text
                    dr("Itemage") = Itemage2.Text
                    dr("ItemCost") = ItemCost2.Text
                    OPrice2.Text = ""
                    ItemCost.Text = ""
                    ItemCost2.Text = ""
                Case 3
                    dr("OPrice") = OPrice3.Text
                    dr("Itemage") = Itemage3.Text
                    OPrice3.Text = ""
                    Itemage3.Text = ""
                Case 4
                    dr("CostID") = CostID4.SelectedValue
                    dr("ItemOther") = If(ItemOther4.Text <> "", ItemOther4.Text, Convert.DBNull)
                    dr("OPrice") = OPrice4.Text
                    dr("Itemage") = Itemage4.Text
                    dr("AdmFlag") = "N"
                    CostID4.SelectedIndex = -1
                    ItemOther4.Text = ""
                    OPrice4.Text = ""
                    Itemage4.Text = ""
            End Select
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            Session(Hid_CostItem_GUID1.Value) = dt
            Call CreateCostItem()
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rqPlanID, rqComIDNO, rqSeqNO) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, ex.ToString)
        End Try
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '自辦計價 新增(隱藏)
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call AddCost(1)
    End Sub

    '每人每時計價 新增(隱藏)
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Call AddCost(2)
    End Sub

    '每人輔助計價 新增(隱藏)
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Call AddCost(3)
    End Sub

    '個人單價計價 新增(隱藏)
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Call AddCost(4)
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If Session(Hid_CostItem_GUID1.Value) Is Nothing Then Return 'Exit Sub
        Dim dt As DataTable = Session(Hid_CostItem_GUID1.Value)
        Dim dr As DataRow
        Select Case e.CommandName
            Case "edit"
                DataGrid1.EditItemIndex = e.Item.ItemIndex
            Case "del"
                If dt.Select("PCID='" & DataGrid1.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select("PCID='" & DataGrid1.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr.Delete()
                End If
            Case "update"
                Dim OPrice As TextBox = e.Item.FindControl("TextBox1")
                Dim Itemage As TextBox = e.Item.FindControl("TextBox2")
                Dim ItemCost As TextBox = e.Item.FindControl("TextBox3")
                If dt.Select("PCID='" & DataGrid1.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select("PCID='" & DataGrid1.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr("OPrice") = OPrice.Text
                    dr("Itemage") = Itemage.Text
                    dr("ItemCost") = ItemCost.Text
                End If
                DataGrid1.EditItemIndex = -1
            Case "cancel"
                DataGrid1.EditItemIndex = -1
        End Select
        'dt.AcceptChanges()
        Session(Hid_CostItem_GUID1.Value) = dt
        Call CreateCostItem()

        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Const cst_DG1_小計 As Integer = 4
        Const cst_DG1_CostID As Integer = 6

        e.Item.Cells(cst_DG1_CostID).Style("display") = "none"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As Label = e.Item.FindControl("Label4")
                Dim Itemage As Label = e.Item.FindControl("Label5")
                Dim ItemCost As Label = e.Item.FindControl("Label6")
                Dim btn1 As Button = e.Item.FindControl("Button4")
                Dim btn2 As Button = e.Item.FindControl("Button5")
                e.Item.Cells(0).Text = ""
                OPrice.Text = ""
                Itemage.Text = ""
                ItemCost.Text = ""
                e.Item.Cells(cst_DG1_小計).Text = ""
                Dim strCostName As String = Get_xCostName1(dtKEYCOSTITEM, drv)
                e.Item.Cells(0).Text = strCostName
                If Convert.ToString(drv("OPrice")) <> "" Then OPrice.Text = Convert.ToString(drv("OPrice"))
                If Convert.ToString(drv("Itemage")) <> "" Then Itemage.Text = Convert.ToString(drv("Itemage"))
                If Convert.ToString(drv("ItemCost")) <> "" Then ItemCost.Text = Convert.ToString(drv("ItemCost"))
                If IsNumeric(OPrice.Text) AndAlso IsNumeric(Itemage.Text) AndAlso IsNumeric(ItemCost.Text) Then e.Item.Cells(cst_DG1_小計).Text = CDbl(OPrice.Text) * CDbl(Itemage.Text) * CDbl(ItemCost.Text)
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn1.Enabled = Button2.Enabled
                btn2.Enabled = Button2.Enabled
            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As TextBox = e.Item.FindControl("TextBox1")
                Dim Itemage As TextBox = e.Item.FindControl("TextBox2")
                Dim ItemCost As TextBox = e.Item.FindControl("TextBox3")
                Dim btn1 As Button = e.Item.FindControl("Button6")
                Dim btn2 As Button = e.Item.FindControl("Button7")
                e.Item.Cells(0).Text = ""
                OPrice.Text = ""
                Itemage.Text = ""
                ItemCost.Text = ""
                e.Item.Cells(cst_DG1_小計).Text = ""
                Dim strCostName As String = Get_xCostName1(dtKEYCOSTITEM, drv)
                e.Item.Cells(0).Text = strCostName

                If Convert.ToString(drv("OPrice")) <> "" Then OPrice.Text = Convert.ToString(drv("OPrice"))
                If Convert.ToString(drv("Itemage")) <> "" Then Itemage.Text = Convert.ToString(drv("Itemage"))
                If Convert.ToString(drv("ItemCost")) <> "" Then ItemCost.Text = Convert.ToString(drv("ItemCost"))
                'e.Item.Cells(4).Text = ""
                If IsNumeric(OPrice.Text) AndAlso IsNumeric(Itemage.Text) AndAlso IsNumeric(ItemCost.Text) Then e.Item.Cells(cst_DG1_小計).Text = CDbl(OPrice.Text) * CDbl(Itemage.Text) * CDbl(ItemCost.Text)
                btn1.Attributes("onclick") = "var msg=check_Cost_Detail(" & OPrice.ClientID & "," & Itemage.ClientID & "," & ItemCost.ClientID & ");if (msg!=''){alert(msg);return false;}"
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        If Session(Hid_CostItem_GUID1.Value) Is Nothing Then Return 'Exit Sub
        Dim dt As DataTable = Session(Hid_CostItem_GUID1.Value)
        Dim dr As DataRow
        Select Case e.CommandName
            Case "edit"
                DataGrid2.EditItemIndex = e.Item.ItemIndex
            Case "del"
                If dt.Select("PCID='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select("PCID='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr.Delete()
                End If
            Case "update"
                Dim OPrice As TextBox = e.Item.FindControl("DataGrid2TextBox1")
                Dim Itemage As TextBox = e.Item.FindControl("DataGrid2TextBox2")
                Dim ItemCost As TextBox = e.Item.FindControl("DataGrid2TextBox3")
                If dt.Select("PCID='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select("PCID='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr("OPrice") = OPrice.Text
                    dr("Itemage") = Itemage.Text
                    dr("ItemCost") = ItemCost.Text
                End If
                DataGrid2.EditItemIndex = -1
            Case "cancel"
                DataGrid2.EditItemIndex = -1
        End Select
        'dt.AcceptChanges()
        Session(Hid_CostItem_GUID1.Value) = dt
        Call CreateCostItem()

        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        e.Item.Cells(5).Style("display") = "none"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As Label = e.Item.FindControl("DataGrid2Label1")
                Dim Itemage As Label = e.Item.FindControl("DataGrid2Label2")
                Dim ItemCost As Label = e.Item.FindControl("DataGrid2Label3")
                Dim btn1 As Button = e.Item.FindControl("Button12")
                Dim btn2 As Button = e.Item.FindControl("Button13")
                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                ItemCost.Text = Convert.ToString(drv("ItemCost"))
                'subtotal.Text = TIMS.Round(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                'subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                e.Item.Cells(3).Text = TIMS.ROUND(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn1.Enabled = Button2.Enabled
                btn2.Enabled = Button2.Enabled
            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As TextBox = e.Item.FindControl("DataGrid2TextBox1")
                Dim Itemage As TextBox = e.Item.FindControl("DataGrid2TextBox2")
                Dim ItemCost As TextBox = e.Item.FindControl("DataGrid2TextBox3")
                Dim btn1 As Button = e.Item.FindControl("Button14")
                Dim btn2 As Button = e.Item.FindControl("Button15")
                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                ItemCost.Text = Convert.ToString(drv("ItemCost"))
                e.Item.Cells(3).Text = TIMS.ROUND(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                btn1.Attributes("onclick") = "var msg=check_Cost_Detail(" & OPrice.ClientID & "," & Itemage.ClientID & "," & ItemCost.ClientID & ");if (msg!=''){alert(msg);return false;}"
        End Select
    End Sub

    Private Sub DataGrid3_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid3.ItemCommand
        If Session(Hid_CostItem_GUID1.Value) Is Nothing Then Return 'Exit Sub
        Dim dt As DataTable = Session(Hid_CostItem_GUID1.Value)
        Dim dr As DataRow
        Select Case e.CommandName
            Case "edit"
                DataGrid3.EditItemIndex = e.Item.ItemIndex
            Case "del"
                If dt.Select("PCID='" & DataGrid3.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select("PCID='" & DataGrid3.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr.Delete()
                End If
            Case "update"
                Dim OPrice As TextBox = e.Item.FindControl("DataGrid3TextBox1")
                Dim Itemage As TextBox = e.Item.FindControl("DataGrid3TextBox2")
                If dt.Select("PCID='" & DataGrid3.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select("PCID='" & DataGrid3.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr("OPrice") = OPrice.Text
                    dr("Itemage") = Itemage.Text
                End If
                DataGrid3.EditItemIndex = -1
            Case "cancel"
                DataGrid3.EditItemIndex = -1
        End Select
        'dt.AcceptChanges()
        Session(Hid_CostItem_GUID1.Value) = dt
        Call CreateCostItem()

        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        e.Item.Cells(4).Style("display") = "none"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As Label = e.Item.FindControl("DataGrid3Label1")
                Dim Itemage As Label = e.Item.FindControl("DataGrid3Label2")
                Dim btn1 As Button = e.Item.FindControl("Button16")
                Dim btn2 As Button = e.Item.FindControl("Button17")
                OPrice.Text = drv("OPrice")
                Itemage.Text = drv("Itemage")
                e.Item.Cells(2).Text = CDbl(OPrice.Text) * CDbl(Itemage.Text)
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn1.Enabled = Button2.Enabled
                btn2.Enabled = Button2.Enabled
            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As TextBox = e.Item.FindControl("DataGrid3TextBox1")
                Dim Itemage As TextBox = e.Item.FindControl("DataGrid3TextBox2")
                Dim btn1 As Button = e.Item.FindControl("Button18")
                Dim btn2 As Button = e.Item.FindControl("Button19")
                OPrice.Text = drv("OPrice")
                Itemage.Text = drv("Itemage")
                e.Item.Cells(2).Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
                btn1.Attributes("onclick") = "var msg=check_Cost_Detail(" & OPrice.ClientID & "," & Itemage.ClientID & ",null);if (msg!=''){alert(msg);return false;}"
        End Select
    End Sub

    Private Sub DataGrid4_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid4.ItemCommand
        If Session(Hid_CostItem_GUID1.Value) Is Nothing Then Return 'Exit Sub
        Dim dt As DataTable = Session(Hid_CostItem_GUID1.Value)
        Dim dr As DataRow
        Select Case e.CommandName
            Case "edit"
                DataGrid4.EditItemIndex = e.Item.ItemIndex
            Case "del"
                If dt.Select("PCID='" & DataGrid4.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select("PCID='" & DataGrid4.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr.Delete()
                End If
            Case "update"
                Dim OPrice As TextBox = e.Item.FindControl("DataGrid4TextBox1")
                Dim Itemage As TextBox = e.Item.FindControl("DataGrid4TextBox2")
                If dt.Select("PCID='" & DataGrid4.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select("PCID='" & DataGrid4.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr("OPrice") = OPrice.Text
                    dr("Itemage") = Itemage.Text
                End If
                DataGrid4.EditItemIndex = -1
            Case "cancel"
                DataGrid4.EditItemIndex = -1
        End Select
        'dt.AcceptChanges()
        Session(Hid_CostItem_GUID1.Value) = dt
        Call CreateCostItem()

        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid4.ItemDataBound
        e.Item.Cells(5).Style("display") = "none"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As Label = e.Item.FindControl("DataGrid4Label1")
                Dim Itemage As Label = e.Item.FindControl("DataGrid4Label2")
                Dim btn1 As Button = e.Item.FindControl("Button20")
                Dim btn2 As Button = e.Item.FindControl("Button21")
                Dim strCostName As String = Get_xCostName1(dtKEYCOSTITEM, drv)
                e.Item.Cells(0).Text = strCostName
                OPrice.Text = drv("OPrice")
                Itemage.Text = drv("Itemage")
                e.Item.Cells(3).Text = CDbl(OPrice.Text) * CDbl(Itemage.Text)
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn1.Enabled = Button11.Enabled
                btn2.Enabled = Button11.Enabled
            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As TextBox = e.Item.FindControl("DataGrid4TextBox1")
                Dim Itemage As TextBox = e.Item.FindControl("DataGrid4TextBox2")
                Dim btn1 As Button = e.Item.FindControl("Button22")
                Dim btn2 As Button = e.Item.FindControl("Button23")
                Dim strCostName As String = Get_xCostName1(dtKEYCOSTITEM, drv)
                e.Item.Cells(0).Text = strCostName
                OPrice.Text = drv("OPrice")
                Itemage.Text = drv("Itemage")
                e.Item.Cells(3).Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
                btn1.Attributes("onclick") = "var msg=check_Cost_Detail(" & OPrice.ClientID & "," & Itemage.ClientID & ",null);if (msg!=''){alert(msg);return false;}"
        End Select
    End Sub

    Public Shared Function Get_xCostName2(ByRef dtKEYCOSTITEM As DataTable, ByRef dr1 As DataRow) As String
        Dim strCostName As String = ""
        Dim vCostID As String = TIMS.ClearSQM(dr1("CostID"))
        Dim fff As String = "CostID='" & dr1("CostID") & "'"
        Select Case vCostID
            Case "99"
                strCostName = "其他-" & Convert.ToString(dr1("ItemOther"))
            Case Else
                If dtKEYCOSTITEM.Select(fff).Length <> 0 Then
                    strCostName = dtKEYCOSTITEM.Select(fff)(0)("CostName")
                End If
        End Select
        Return strCostName
    End Function

    Public Shared Function Get_xCostName1(ByRef dtKEYCOSTITEM As DataTable, ByRef drv As DataRowView) As String
        Dim strCostName As String = ""
        Dim vCostID As String = TIMS.ClearSQM(drv("CostID"))
        Dim fff As String = "CostID='" & vCostID & "'"
        Select Case vCostID
            Case "99"
                strCostName = "其他-" & Convert.ToString(drv("ItemOther"))
            Case Else
                If dtKEYCOSTITEM.Select(fff).Length <> 0 Then
                    strCostName = dtKEYCOSTITEM.Select(fff)(0)("CostName")
                End If
        End Select
        'e.Item.Cells(0).Text = strCostName
        Return strCostName
    End Function

    Private Sub DataGrid5_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid5.ItemCommand
        If Session(Hid_TrainDesc_GUID1.Value) Is Nothing Then Return 'Exit Sub
        Dim dt As DataTable = Session(Hid_TrainDesc_GUID1.Value)
        Select Case e.CommandName
            Case "edit"
                DataGrid5.EditItemIndex = e.Item.ItemIndex
            Case "del"
                Dim filter As String = ""
                filter = cst_PlanTrainDescPKName & "='" & e.CommandArgument & "'"
                If dt.Select(filter).Length <> 0 Then dt.Select(filter)(0).Delete()
                DataGrid5.EditItemIndex = -1
                'dt.AcceptChanges()
                Session(Hid_TrainDesc_GUID1.Value) = dt
            Case "save"
                Dim TPName As TextBox = e.Item.FindControl("TPName")
                Dim TPHour As TextBox = e.Item.FindControl("TPHour")
                Dim TPCont As TextBox = e.Item.FindControl("TPCont")
                Dim filter As String = ""
                filter = cst_PlanTrainDescPKName & "='" & e.CommandArgument & "'" 'PTDID
                If dt.Select(filter).Length <> 0 Then
                    Dim dr As DataRow = dt.Select(filter)(0)
                    dr("PName") = TPName.Text
                    dr("PHour") = Val(TPHour.Text)
                    dr("PCont") = TIMS.ClearSQM(TPCont.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                DataGrid5.EditItemIndex = -1
                'dt.AcceptChanges()
                Session(Hid_TrainDesc_GUID1.Value) = dt
            Case "cancel"
                DataGrid5.EditItemIndex = -1
        End Select
        '顯示 訓練內容簡介
        Call ShowTrainDesc()

        If (LayerState.Value = "") Then LayerState.Value = "10"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid5_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid5.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button30")
                Dim btn2 As Button = e.Item.FindControl("Button31")
                Dim LPName As Label = e.Item.FindControl("LPName")
                Dim LPHour As Label = e.Item.FindControl("LPHour")
                Dim LPCont As Label = e.Item.FindControl("LPCont")
                LPName.Text = drv("PName").ToString()
                LPHour.Text = drv("PHour").ToString()
                LPCont.Text = TIMS.ClearSQM(Convert.ToString(drv("PCont"))) '.ToString
                btn1.CommandArgument = drv(cst_PlanTrainDescPKName).ToString()
                btn2.CommandArgument = drv(cst_PlanTrainDescPKName).ToString()
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn1.Enabled = Button29.Enabled
                btn2.Enabled = Button29.Enabled
            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button32")
                Dim btn2 As Button = e.Item.FindControl("Button33")
                Dim TPName As TextBox = e.Item.FindControl("TPName")
                Dim TPHour As TextBox = e.Item.FindControl("TPHour")
                Dim TPCont As TextBox = e.Item.FindControl("TPCont")
                TPName.Text = drv("PName").ToString()
                TPHour.Text = drv("PHour").ToString()
                TPCont.Text = TIMS.ClearSQM(Convert.ToString(drv("PCont"))) '
                btn1.CommandArgument = drv(cst_PlanTrainDescPKName).ToString()
                btn2.CommandArgument = drv(cst_PlanTrainDescPKName).ToString()
                btn1.Attributes("onclick") = "return CheckDescData('" & TPName.ClientID & "','" & TPHour.ClientID & "','" & TPCont.ClientID & "');"
                btn1.Enabled = Button29.Enabled
                btn2.Enabled = Button29.Enabled
        End Select
    End Sub

    ''' <summary>
    ''' 草稿儲存-檢查輸入資料的正確性
    ''' </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function sUtl_CheckTemp1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True '沒有異常為True
        Errmsg = ""
        Call sUtl_SortOut1() '整理輸入值
        If RIDValue.Value = "" Then Errmsg &= "請選擇訓練機構!" & vbCrLf
        If GenSciHours.Text <> "" AndAlso Not TIMS.IsNumeric1(GenSciHours.Text) Then Errmsg &= "一般學科必須為數字" & vbCrLf
        If ProSciHours.Text <> "" AndAlso Not TIMS.IsNumeric1(ProSciHours.Text) Then Errmsg &= "專業學科必須為數字" & vbCrLf
        If ProTechHours.Text <> "" AndAlso Not TIMS.IsNumeric1(ProTechHours.Text) Then Errmsg &= "術科必須為數字" & vbCrLf
        If OtherHours.Text <> "" AndAlso Not TIMS.IsNumeric1(OtherHours.Text) Then Errmsg &= "其他時數必須為數字" & vbCrLf

        'Dim v_rblADVANCE As String = TIMS.GetListValue(rblADVANCE) '訓練課程類型
        'If v_rblADVANCE = "" Then Errmsg &= "訓練課程類型 單選，選項包括：基礎、進階，必填" & vbCrLf
        TNum.Text = TIMS.ClearSQM(TNum.Text)
        If TNum.Text = "" Then
            Errmsg &= "訓練人數必須輸入為數字!" & vbCrLf
        ElseIf TNum.Text <> "" AndAlso Not TIMS.IsNumberStr(TNum.Text) Then
            Errmsg &= "訓練人數必須為數字!" & vbCrLf
        ElseIf TNum.Text <> "" AndAlso Not TIMS.IsNumeric2(TNum.Text) Then
            Errmsg &= "訓練人數必須為數字!!" & vbCrLf
        ElseIf TNum.Text <> "" AndAlso TIMS.IsNumberStr(TNum.Text) AndAlso Hid_MaxTNum.Value <> "" AndAlso TIMS.IsNumberStr(Hid_MaxTNum.Value) _
            AndAlso TIMS.VAL1(TNum.Text) > TIMS.VAL1(Hid_MaxTNum.Value) Then
            Errmsg &= String.Format("訓練人數上限為{0}人, 輸入{1}人超過系統限制", Hid_MaxTNum.Value, TNum.Text) & vbCrLf
        End If

        THours.Text = TIMS.ClearSQM(THours.Text)
        STDate.Text = TIMS.ClearSQM(STDate.Text)
        FDDate.Text = TIMS.ClearSQM(FDDate.Text)
        If THours.Text = "" Then Errmsg &= "訓練時數必須輸入為數字" & vbCrLf
        If STDate.Text = "" Then Errmsg &= "訓練起日必須輸入日期格式" & vbCrLf
        If FDDate.Text = "" Then Errmsg &= "訓練迄日必須輸入日期格式" & vbCrLf
        If THours.Text <> "" AndAlso Not TIMS.IsNumeric1(THours.Text) Then Errmsg &= "訓練時數必須為數字" & vbCrLf
        If STDate.Text <> "" AndAlso Not TIMS.IsDate1(STDate.Text) Then Errmsg &= "訓練起日不是正確的日期格式" & vbCrLf
        If Errmsg = "" Then STDate.Text = TIMS.Cdate3(STDate.Text)
        If FDDate.Text <> "" AndAlso Not TIMS.IsDate1(FDDate.Text) Then Errmsg &= "訓練迄日不是正確的日期格式" & vbCrLf
        If Errmsg = "" Then FDDate.Text = TIMS.Cdate3(FDDate.Text)

        Dim flag_NoSameDay1 As Boolean = True '訓練起日不能和訓練迄日同一天
        '計畫：接受企業委託訓練 可同一天
        If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_NoSameDay1 = False
        If Errmsg = "" AndAlso STDate.Text <> "" AndAlso FDDate.Text <> "" Then
            If TIMS.IsDate1(STDate.Text) AndAlso TIMS.IsDate1(FDDate.Text) Then
                Dim intX As Long = DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(FDDate.Text))
                If intX = 0 AndAlso flag_NoSameDay1 Then Errmsg &= "訓練起日不能和訓練迄日同一天" & vbCrLf
                If intX < 0 Then Errmsg &= "訓練起日不能超過訓練迄日" & vbCrLf
            End If
        End If

        If CyclType.Text <> "" AndAlso Not TIMS.IsNumeric1(CyclType.Text) Then Errmsg &= "期別必須為數字" & vbCrLf
        If ClassCount.Text <> "" AndAlso Not TIMS.IsNumeric1(ClassCount.Text) Then Errmsg &= "班數必須為數字" & vbCrLf
        If DefGovCost.Text <> "" AndAlso Not TIMS.IsNumeric1(DefGovCost.Text) Then Errmsg &= "政府負擔費用必須為數字" & vbCrLf
        If DefUnitCost.Text <> "" AndAlso Not TIMS.IsNumeric1(DefUnitCost.Text) Then Errmsg &= "企業負擔費用必須為數字" & vbCrLf
        If DefStdCost.Text <> "" AndAlso Not TIMS.IsNumeric1(DefStdCost.Text) Then Errmsg &= "學員負擔費用必須為數字" & vbCrLf
        If TAddressZIPB3.Value <> "" AndAlso Not TIMS.IsNumeric1(TAddressZIPB3.Value) Then Errmsg &= "郵遞區號後2碼或後3碼必須為數字" & vbCrLf
        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    '草稿儲存
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '20080811 andy  學習券計畫同一訓練機構只允計申請一個班級暫不使用
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then return 'Exit Sub
        Dim Errmsg As String = ""
        Call sUtl_CheckTemp1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)

            If (LayerState.Value = "") Then LayerState.Value = "1"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("_onload", s_js11)
            Return 'Exit Sub
        End If

        '草稿儲存
        Call SAVE_PLAN_PLANINFO(2)
        If (LayerState.Value = "") Then LayerState.Value = "1"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("_onload", s_js1)
    End Sub

    '整理輸入值
    Sub sUtl_SortOut1()
        SEnterDate.Text = TIMS.Cdate3(TIMS.ClearSQM(SEnterDate.Text))
        FEnterDate.Text = TIMS.Cdate3(TIMS.ClearSQM(FEnterDate.Text))
        ExamDate.Text = TIMS.Cdate3(TIMS.ClearSQM(ExamDate.Text))
        CheckInDate.Text = TIMS.Cdate3(TIMS.ClearSQM(CheckInDate.Text))
        STDate.Text = TIMS.Cdate3(TIMS.ClearSQM(STDate.Text))
        FDDate.Text = TIMS.Cdate3(TIMS.ClearSQM(FDDate.Text))

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        GenSciHours.Text = TIMS.ClearSQM(GenSciHours.Text)
        ProSciHours.Text = TIMS.ClearSQM(ProSciHours.Text)
        ProTechHours.Text = TIMS.ClearSQM(ProTechHours.Text)
        OtherHours.Text = TIMS.ClearSQM(OtherHours.Text)
        TNum.Text = TIMS.ClearSQM(TNum.Text)
        THours.Text = TIMS.ClearSQM(THours.Text)

        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        ClassCount.Text = TIMS.ClearSQM(ClassCount.Text)
        DefGovCost.Text = TIMS.ClearSQM(DefGovCost.Text)
        DefUnitCost.Text = TIMS.ClearSQM(DefUnitCost.Text)
        DefStdCost.Text = TIMS.ClearSQM(DefStdCost.Text)
        TAddressZIPB3.Value = TIMS.ClearSQM(TAddressZIPB3.Value)
    End Sub

    'sUtl_CheckData1 儲存前先檢查輸入資料的正確性 (正式儲存檢核)
    ''' <summary>儲存前先檢查輸入資料的正確性</summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SUtl_CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True '沒有異常為True
        Errmsg = ""

        Call sUtl_SortOut1() '整理輸入值

        'Dim flag_TPlanID07_show As Boolean = False
        'If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_TPlanID07_show = True
        'If flag_TPlanID07_show Then
        '    TRNUNITNAME.Text = TIMS.ClearSQM(TRNUNITNAME.Text)
        '    TRNUNITTYPE.Text = TIMS.ClearSQM(TRNUNITTYPE.Text)
        '    TRNUNITEE.Text = TIMS.ClearSQM(TRNUNITEE.Text)
        '    '「委訓單位名稱」、「委訓單位類型」、「訓練對象」
        '    If TRNUNITNAME.Text = "" Then
        '        Errmsg &= "訓練對象及資格「委訓單位名稱」為必填欄位!" & vbCrLf
        '    End If
        '    If TRNUNITTYPE.Text = "" Then
        '        Errmsg &= "訓練對象及資格「委訓單位類型」為必填欄位!" & vbCrLf
        '    End If
        '    If TRNUNITEE.Text = "" Then
        '        Errmsg &= "訓練對象及資格「訓練對象」為必填欄位!" & vbCrLf
        '    End If
        'End If
        'If Errmsg <> "" Then Return False '有錯誤訊息'不可儲存

        'OJT-23041104：區域據點-開班資料查詢：【班級英文名稱】改為非必填 
        'OJT-23041103：區域據點-班級申請：【班級英文名稱】改為非必填
        If (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) = -1) Then
            Dim vClassEngName As String = TIMS.ClearSQM(ClassEngName.Text)
            If vClassEngName = "" Then Errmsg &= "請輸入班級英文名稱!" & vbCrLf
        End If

        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If trainValue.Value = "" Then Errmsg &= "請選擇訓練職類，訓練職類為必須選擇" & vbCrLf
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If cjobValue.Value = "" Then Errmsg &= "請選擇通俗職類，通俗職類為必須選擇" & vbCrLf
        THours.Text = TIMS.ClearSQM(THours.Text)
        If THours.Text = "" Then Errmsg &= "班別資料「訓練時數」必須填寫" & vbCrLf
        If Errmsg <> "" Then Return False 'Exit Sub

        Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    Errmsg &= "於處分日期起的期間，班級申請資料建檔不可正式儲存。" & vbCrLf
                    Return False
            End Select
        End If

        '照顧服務員自訓自用訓練計畫
        '因違反第20點規定：自處分日期起1年內，該單位不得申請照顧服務員自訓自用訓練計畫之訓練單位。處分日期:2017/05/23
        If TIMS.Utl_GetConfigSet("work2017_1") = TIMS.cst_YES Then
            If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, OBTERMS.cst_c07, objconn) Then
                Errmsg += OBTERMS.cst_c07_altMsg1
                'Return False
            End If
            If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, OBTERMS.cst_c20, objconn) Then
                Errmsg += OBTERMS.cst_c20_altMsg1
                'Return False
            End If
            If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, OBTERMS.cst_c21, objconn) Then
                Errmsg += OBTERMS.cst_c21_altMsg1
                'Return False
            End If
            If Errmsg <> "" Then Return False
        End If

        'If TIMS.sUtl_ChkTest() Then Errmsg &= "於處分日期起的期間，班級申請資料建檔不可正式儲存。"
        If Errmsg <> "" Then Return False '有錯誤訊息'不可儲存

        If SEnterDate.Text = "" Then
            Errmsg += "班別資料「報名開始日期」為必填欄位!" & vbCrLf
            Return False
        ElseIf FEnterDate.Text = "" Then
            Errmsg += "班別資料「報名結束日期」為必填欄位!" & vbCrLf
            Return False
        End If

        'Dim v_GetTrain3 As String = TIMS.GetCblValue(GetTrain3) '甄試方式
        If Not flag_TPlanID70_1 Then '區域據點 不卡控-甄試日期 
            If ExamDate.Text = "" Then '甄試日期 --自辦在職 為必填 --調整甄試日期欄位相關卡控
                Errmsg += "班別資料「甄試日期」為必填欄位!" & vbCrLf
                Return False
            End If
        End If

        If CheckInDate.Text = "" Then
            Errmsg += "班別資料「報到日期」為必填欄位!" & vbCrLf
            Return False
        ElseIf STDate.Text = "" Then
            Errmsg += "班別資料「訓練起日」為必填欄位!" & vbCrLf
            Return False
        ElseIf FDDate.Text = "" Then
            Errmsg += "班別資料「訓練迄日」為必填欄位!" & vbCrLf
            Return False
        End If

        If Errmsg <> "" Then Return False '有錯誤訊息'不可儲存
        'If (CDate(ExamDate.Text) >= CDate(CheckInDate.Text)) Then
        '    Errmsg += "班別資料[報到日期]必須大於[甄試日期]!" & vbCrLf
        'End If
        'If (CDate(CheckInDate.Text) > CDate(STDate.Text)) Then
        '    Errmsg += "班別資料[訓練起日]必須大於等於[報到日期]!" & vbCrLf
        'End If
        If (CDate(SEnterDate.Text) >= CDate(FEnterDate.Text)) Then Errmsg += "班別資料[報名結束日期]必須大於[報名開始日期]!" & vbCrLf
        If (CDate(STDate.Text) <= CDate(FEnterDate.Text)) Then Errmsg += "班別資料[訓練起日]必須大於[報名結束日期]!" & vbCrLf

        Dim v_ExamPeriod As String = TIMS.GetListValue(ExamPeriod)
        '甄試日期 --自辦在職 為必填 If (v_ExamPeriod <> "", v_ExamPeriod, Convert.DBNull)
        If ExamDate.Text <> "" AndAlso ExamPeriod.SelectedIndex = 0 AndAlso v_ExamPeriod = "" Then '20100329 add 甄試時段
            Errmsg += "班別資料「甄試日期」全天、上午、下午 時段請擇一選擇!" & vbCrLf
            'Common.MessageBox(Me, "「甄試日期」全天、上午、下午 時段請擇一選擇!")
            'Exit Function
        End If
        If ExamDate.Text <> "" AndAlso FEnterDate.Text <> "" AndAlso (CDate(ExamDate.Text) <= CDate(FEnterDate.Text)) Then
            Errmsg += "班別資料「甄試日期」必須大於「報名結束日期」!" & vbCrLf
            'Common.MessageBox(Me, "「甄試日期」必須大於「報名結束日期」!")
            'Exit Function
            'Else
            '計算天(所有天數)
            'Dim iDayFE As Integer = DateDiff(DateInterval.Day, CDate(FEnterDate.Text), CDate(ExamDate.Text))
            '假日天(扣除)
            'Dim iDayFE2 As Integer = TIMS.Get_SysHolidayDay(Me, CDate(FEnterDate.Text), CDate(ExamDate.Text), objconn)
            'If (iDayFE - iDayFE2) < 2 Then
            ' '「甄試日期」最快得安排於報名截止當日起2日後。
            ' Errmsg += "班別資料「甄試日期」最快得安排於「報名結束日期」當日起2日後!(工作天)" & vbCrLf
            'End If
            'If DateDiff(DateInterval.Day, CDate(FEnterDate.Text), CDate(ExamDate.Text))+iDayFE < 2 Then
            ' Errmsg += "班別資料「甄試日期」最快得安排於「報名結束日期」當日起2日後!" & vbCrLf
            'End If
        End If
        If ExamDate.Text <> "" AndAlso STDate.Text <> "" AndAlso (CDate(ExamDate.Text) > CDate(STDate.Text)) Then
            Errmsg += "班別資料「甄試日期」必須小於或等於[訓練起日]!" & vbCrLf
            'Common.MessageBox(Me, "[甄試日期]必須小於或等於[開訓日期]!") 'Exit Function
        End If
        'If ExamDate.Text <> "" ANDALSO TIMS.Chk_HOLDATE(Hid_RID1.Value, ExamDate.Text, objconn) Then ' Errmsg += "班別資料「甄試日期」不可為例假日!" & vbCrLf 'End If

        If RIDValue.Value <> "" Then Hid_RID1.Value = Convert.ToString(RIDValue.Value).Substring(0, 1)
        If Hid_RID1.Value = "" Then Hid_RID1.Value = Convert.ToString(sm.UserInfo.RID).Substring(0, 1)

        If Errmsg <> "" Then
            rst = False
            Return rst
        End If

        '2019年啟用 work2019x01:2019 政府政策性產業
        Dim sErrMsg1 As String = CHK_KID20_VAL()
        If trKID20.Visible AndAlso sErrMsg1 <> "" Then
            Errmsg &= sErrMsg1 ' Common.MessageBox(Me, sErrMsg1)
            Return False
        End If
        Dim sErrMsg2 As String = TIMS.CHK_KID60_VAL(CBLKID60)
        If fg_USE_CBLKID60_TP06 AndAlso sErrMsg2 <> "" Then
            Errmsg &= sErrMsg2 ' Common.MessageBox(Me, sErrMsg2)
            Return False 'Exit Sub
        End If
        Dim sErrMsg3 As String = CHK_KID25_VAL()
        If trKID25.Visible AndAlso sErrMsg3 <> "" Then
            Errmsg &= sErrMsg3 ' Common.MessageBox(Me, sErrMsg1)
            Return False
        End If

        Dim v_rblADVANCE As String = TIMS.GetListValue(rblADVANCE) '訓練課程類型
        If v_rblADVANCE = "" Then Errmsg &= "班別資料 訓練課程類型 單選必填!" & vbCrLf

        TAddressZip.Value = TIMS.ClearSQM(TAddressZip.Value)
        TAddressZIPB3.Value = TIMS.ClearSQM(TAddressZIPB3.Value)
        TAddress.Text = TIMS.ClearSQM(TAddress.Text)
        If TAddressZip.Value = "" Then Errmsg &= "班別資料 上課地址「郵遞區號前3碼」資料有誤!" & vbCrLf
        If TAddressZIPB3.Value = "" Then Errmsg &= "班別資料 上課地址「郵遞區號後2碼或後3碼」資料有誤!" & vbCrLf
        If TAddress.Text = "" Then Errmsg &= "班別資料 上課地址「地址」資料有誤!" & vbCrLf

        EAddressZip.Value = TIMS.ClearSQM(EAddressZip.Value)
        EAddressZIPB3.Value = TIMS.ClearSQM(EAddressZIPB3.Value)
        EAddress.Text = TIMS.ClearSQM(EAddress.Text)
        If EAddressZip.Value = "" Then Errmsg &= "班別資料 報名地點「郵遞區號前3碼」資料有誤!" & vbCrLf
        If EAddressZIPB3.Value = "" Then Errmsg &= "班別資料 報名地點「郵遞區號後2碼或後3碼」資料有誤!" & vbCrLf
        If EAddress.Text = "" Then Errmsg &= "班別資料 報名地點「地址」資料有誤!" & vbCrLf

        twiACTNO.Text = TIMS.ClearSQM(twiACTNO.Text)
        If twiACTNO.Text = "" Then
            Errmsg &= "班別資料「訓字保保險證號」為必填欄位!" & vbCrLf
        ElseIf twiACTNO.Text.Length < 2 Then
            Errmsg &= "班別資料「訓字保保險證號」資料長度有誤!" & vbCrLf
        ElseIf twiACTNO.Text.Substring(0, 2) <> "09" Then
            Errmsg &= "班別資料「訓字保保險證號」應為09開頭" & vbCrLf
        End If

        'lab_LayerC2 : (受訓資格)
        'Dim tmp_Errmsg1 As String = "" '受訓資格-lab_LayerC2.Text
        txtAge2.Text = TIMS.ChangeIDNO(txtAge2.Text)
        txtAge2.Text = TIMS.ClearSQM(txtAge2.Text)
        If rdoAge2.Checked Then
            If txtAge2.Text = "" Then Errmsg &= String.Format("{0}-「年齡有上限」為必填欄位，請輸入數字", lab_LayerC2.Text) & vbCrLf
            If Errmsg = "" AndAlso Not TIMS.IsNumeric2(txtAge2.Text) Then Errmsg &= String.Format("{0}-「年齡有上限」為必填欄位，請輸入數字", lab_LayerC2.Text) & vbCrLf
            If Errmsg = "" AndAlso Val(txtAge2.Text) <= 15 Then Errmsg &= String.Format("{0}-「年齡有上限」輸入範圍有誤，請輸入大於15的數字", lab_LayerC2.Text) & vbCrLf
        End If

        '計畫：接受企業委託計畫  (僅此計畫! 自辦不顯示喔) Hid_TRNUNIT.Value = sm.UserInfo.TPlanID
        If Hid_TRNUNIT.Value <> "" Then
            '委訓單位名稱
            TRNUNITNAME.Text = TIMS.ClearSQM(TRNUNITNAME.Text)
            TRNUNITTYPE.Text = TIMS.ClearSQM(TRNUNITTYPE.Text)
            If TRNUNITNAME.Text = "" Then
                Errmsg &= String.Format("{0}「委訓單位名稱」為必填欄位，請輸入文字", lab_LayerC2.Text) & vbCrLf
            End If
            'TRNUNITCHO-委訓單位類型
            '1:政府機關/ 2:公民營事業機構/ 3:學校/ 4:團體/ 9:其他(請說明)
            Dim v_TRNUNITCHO As String = TIMS.GetListValue(TRNUNITCHO)
            If v_TRNUNITCHO = "" Then
                Errmsg &= String.Format("{0}「委訓單位類型」為必須選擇，請選擇1項", lab_LayerC2.Text) & vbCrLf
            End If
            If v_TRNUNITCHO <> "" AndAlso TRNUNITTYPE.Text = "" Then
                Const cst_其他_請說明 As String = "9"
                If v_TRNUNITCHO = cst_其他_請說明 Then
                    Errmsg &= String.Format("{0}「委訓單位類型」選擇-其他-說明欄不可為空，請填寫說明欄", lab_LayerC2.Text) & vbCrLf
                End If
            End If
        End If

        'If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If Errmsg = "" AndAlso Hid_MaxTNum.Value <> "" Then
        '        If Val(TNum.Text) > Val(Hid_MaxTNum.Value) Then
        '            Dim sTPlanname As String = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        '            Errmsg &= "計畫人數上限為20人" & vbCrLf
        '            Return False
        '        End If
        '    End If
        'End If
        '檢核：
        '「是否為輔導考照班」選擇「是」時，第一筆為必填(至少需填一筆)，若沒填請跳出提示訊息。
        '是否為輔導考照班 COACHING RBCOACHING_Y
        '檢定職類代碼1/2/3 與考試級別1/2/3 EXAMIDS1/ EXAMLVID1 //EXAM1val.Value /EXLV1val.Value
        'txtGP1', 'txtXM1', 'txtLV1', 'EXAM1val', 'EXLV1val' 
        If (Not RBCOACHING_Y.Checked AndAlso (Not RBCOACHING_N.Checked)) Then
            Errmsg &= "其他「是否為輔導考照班」必須填寫，選擇「是」或「否」" & vbCrLf
        End If
        If (RBCOACHING_Y.Checked) AndAlso (EXAM1val.Value = "") Then
            Errmsg &= "其他「是否為輔導考照班」選擇「是」時，「檢定職類與考試級別」1.為必填(至少需填1筆)" & vbCrLf
        End If
        If (EXAM2val.Value <> "") AndAlso (EXAM1val.Value = "") Then
            Errmsg &= "其他「檢定職類與考試級別」2.有填寫時，「檢定職類與考試級別」1.為必填" & vbCrLf
        End If
        If (EXAM3val.Value <> "") AndAlso (EXAM2val.Value = "") Then
            Errmsg &= "其他「檢定職類與考試級別」3.有填寫時，「檢定職類與考試級別」2.為必填" & vbCrLf
        End If

        '專長能力標籤-ABILITY
        Const cst_errmsg36 As String = "[其他] 專長能力至少須填寫1個,請填 「專長能力標籤」1.名稱!"
        For i_SEQ As Integer = 1 To 4
            Dim s_SEQ As String = Convert.ToString(i_SEQ)
            Dim otxtA1 As TextBox = If(s_SEQ = "1", txtABILITY1, If(s_SEQ = "2", txtABILITY2, If(s_SEQ = "3", txtABILITY3, If(s_SEQ = "4", txtABILITY4, Nothing))))
            Dim otxtA2 As TextBox = If(s_SEQ = "1", txtABILITY_DESC1, If(s_SEQ = "2", txtABILITY_DESC2, If(s_SEQ = "3", txtABILITY_DESC3, If(s_SEQ = "4", txtABILITY_DESC4, Nothing))))
            otxtA1.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA1.Text), 30)
            otxtA2.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA2.Text), 200)
            If otxtA1.Text = "" Then
                Errmsg &= cst_errmsg36 & vbCrLf
                Exit For
            End If
            Exit For
        Next

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>'送出'正式儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then return 'Exit Sub
        '20080811 andy  學習券計畫同一訓練機構只允計申請一個班級暫不使用
        Dim Errmsg As String = ""
        Dim fgCHKTMP1 As Boolean = sUtl_CheckTemp1(Errmsg)
        If Not fgCHKTMP1 OrElse Errmsg <> "" Then
            If Errmsg = "" Then Errmsg = "檢核有誤，請確認儲存資料!"
            Common.MessageBox(Me, Errmsg)
            If (LayerState.Value = "") Then LayerState.Value = "1"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("_onload", s_js11)
            Return 'Exit Sub
        End If

        '儲存前先檢查輸入資料的正確性 (正式儲存檢核)
        Dim fgCHK1 As Boolean = SUtl_CheckData1(Errmsg)
        If Not fgCHK1 OrElse Errmsg <> "" Then
            If Errmsg = "" Then Errmsg = "檢核有誤，請確認儲存資料!!"
            Common.MessageBox(Me, Errmsg)
            If (LayerState.Value = "") Then LayerState.Value = "1"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("_onload", s_js11)
            Return 'Exit Sub
        End If

        '檢核@chkddlJGID 'If Page.IsValid Then '正式儲存 1
        Call SAVE_PLAN_PLANINFO(1)
        'Else 'args.Value ' Common.MessageBox(Page, "(前端驗證資料錯誤)儲存失敗!!")
        ' Page.RegisterStartupScript("_onload", "<script language=""javascript"">Layer_change(1);</script>") 'End If
    End Sub

    '回上一頁
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        'If ViewState(cst_searchTC02001) <> "" Then Session(cst_searchTC02001) =ViewState(cst_searchTC02001)
        Dim url1 As String = ""
        Dim rqMID As String = TIMS.Get_MRqID(Me)

        If ViewState("Redirect") Is Nothing Then
            If TIMS.ClearSQM(Request("todo")) = "1" Then
                url1 = "../04/TC_04_001.aspx?ID=" & rqMID
            ElseIf gflag_ccopy Then
                url1 = "../03/TC_03_002.aspx?ID=" & rqMID
            Else
                url1 = "../02/TC_02_001.aspx?ID=" & rqMID
            End If
        Else
            If ViewState("Redirect") <> "" Then url1 = ViewState("Redirect") & "?ID=" & rqMID
        End If
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '取得機構資訊帶入預設值
    Sub Get_OrgPlanInfo(Optional ByVal sType As Integer = 1)
        'sType :1 為第一次登入動作 sm.UserInfo.RID 'sType :2 為後續動作  RIDValue.Value
        Dim dr As DataRow = Nothing
        Dim sql As String = ""
        sql &= " Select a.RID ,b.orgname ,b.ComIDNO ,c.ContactName ,c.Phone ,c.ContactEmail ,c.ZipCode ,c.ZipCODE6W ,c.Address" & vbCrLf
        sql &= " FROM Auth_Relship a" & vbCrLf
        sql &= " JOIN Org_OrgInfo b On a.OrgID = b.OrgID And a.RID=@RID" & vbCrLf
        sql &= " JOIN Org_OrgPlanInfo c On a.RSID = c.RSID" & vbCrLf
        Using oCmd As New SqlCommand(sql, objconn)
            With oCmd
                .Parameters.Clear()
                Select Case sType
                    Case 1
                        .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                    Case 2
                        .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                End Select
            End With
            dr = TIMS.GetOneRow(oCmd, objconn)
        End Using
        If dr Is Nothing Then
            'Common.MessageBox(Me, "程式出錯，請聯絡東柏人員!")
            Common.MessageBox(Me, "程式出現例外狀況(查無業務機構資料)，請聯絡TIMS系統駐點人員!")
            Return 'Exit Sub
        End If

        RIDValue.Value = dr("RID").ToString()
        ComidValue.Value = dr("ComIDNO").ToString()
        center.Text = dr("orgname").ToString()
        EMail.Text = dr("ContactEmail").ToString()
        'RIDValue.Value = sm.UserInfo.RID
        'ComidValue.Value = dr("ComIDNO").ToString()
        'center.Text = dr("orgname")
        'If Table1.Visible = True Then EMail.Text = dr("ContactEmail").ToString()
        CCTName.Text = ""
        TAddressZip.Value = ""
        TAddressZIPB3.Value = ""
        If dr("ZipCode").ToString() <> "" Then
            TAddressZip.Value = Convert.ToString(dr("ZipCode")) 'TIMS.AddZero(Convert.ToString(dr("ZipCode")), 3)
            hidTAddressZIP6W.Value = Convert.ToString(dr("ZipCODE6W"))
            TAddressZIPB3.Value = TIMS.GetZIPCODEB3(hidTAddressZIP6W.Value)
            CCTName.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("ZipCode")), Convert.ToString(dr("ZipCODE6W")))
        End If
        TAddress.Text = dr("Address").ToString()
        If TAddressZip.Value <> "" Then TAddressZip.Value = Trim(TAddressZip.Value)
        If TAddressZIPB3.Value <> "" Then TAddressZIPB3.Value = Trim(TAddressZIPB3.Value)
        If TAddress.Text <> "" Then TAddress.Text = Trim(TAddress.Text)

        '報名地點/甄試地點
        ECTName.Text = ""
        EAddressZip.Value = ""
        EAddressZIPB3.Value = ""
        If dr("ZipCode").ToString() <> "" Then
            EAddressZip.Value = Convert.ToString(dr("ZipCode")) 'TIMS.AddZero(Convert.ToString(dr("ZipCode")), 3)
            hidEAddressZIP6W.Value = Convert.ToString(dr("ZipCODE6W"))
            EAddressZIPB3.Value = TIMS.GetZIPCODEB3(hidEAddressZIP6W.Value)
            ECTName.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("ZipCode")), Convert.ToString(dr("ZipCODE6W")))
        End If
        EAddress.Text = dr("Address").ToString()
        If EAddressZip.Value <> "" Then EAddressZip.Value = Trim(EAddressZip.Value)
        If EAddressZIPB3.Value <> "" Then EAddressZIPB3.Value = Trim(EAddressZIPB3.Value)
        If EAddress.Text <> "" Then EAddress.Text = Trim(EAddress.Text)

        '聯絡人姓名
        If ContactName.Text = "" Then ContactName.Text = Convert.ToString(dr("ContactName"))
        '聯絡人電話
        If ContactPhone.Text = "" Then ContactPhone.Text = Convert.ToString(dr("Phone"))
        '聯絡人電子郵件
        If ContactEmail.Text = "" Then ContactEmail.Text = Convert.ToString(dr("ContactEmail"))
    End Sub

    '機構資訊(隱藏)
    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Call Get_OrgPlanInfo(2) '取得機構資訊帶入選擇值 RIDValue.Value
        Page.RegisterStartupScript("Londing", "<script>Layer_change('');</script>")
    End Sub

    '新增內容簡介
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        'Dim sql As String
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Try
            'Dim da As SqlDataAdapter = nothing
            If Session(Hid_TrainDesc_GUID1.Value) Is Nothing Then
                Dim sql As String = "SELECT * FROM PLAN_TRAINDESC WHERE 1<>1"
                dt = DbAccess.GetDataTable(sql, objconn)
                dt.Columns(cst_PlanTrainDescPKName).AutoIncrement = True
                dt.Columns(cst_PlanTrainDescPKName).AutoIncrementSeed = -1
                dt.Columns(cst_PlanTrainDescPKName).AutoIncrementStep = -1
            Else
                dt = Session(Hid_TrainDesc_GUID1.Value)
            End If
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("PName") = PName.Text
            dr("PHour") = If(PHour.Text = "", 0, Val(PHour.Text))
            dr("PCont") = If(PCont.Text = "", " ", TIMS.ClearSQM(PCont.Text))
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            Session(Hid_TrainDesc_GUID1.Value) = dt
            Call ShowTrainDesc() '顯示 訓練內容簡介
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rqPlanID, rqComIDNO, rqSeqNO) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, ex.ToString)
        End Try

        PName.Text = ""
        PHour.Text = ""
        PCont.Text = ""

        If (LayerState.Value = "") Then LayerState.Value = "10"
        Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js11)
    End Sub

    '匯入簡介
    Private Sub Btn_TrainDescImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_TrainDescImport.Click
        Const cst_flag_tims1 As String = "tims1" '一般 TIMS匯入
        Const cst_flag_tims4768 As String = "tims4768" '採 新匯入方式

        '47 補助辦理照顧服務員職業訓練'68 照顧服務員自訓自用訓練計畫
        Dim sTPlanID_TYPE As String = cst_flag_tims1
        If TIMS.Cst_TPlanID47AppPlan7.IndexOf(sm.UserInfo.TPlanID) > -1 Then sTPlanID_TYPE = cst_flag_tims4768

        Dim Errmsg As String = ""
        Select Case sTPlanID_TYPE
            Case cst_flag_tims1
                Dim flagOK As Boolean = Sub_CsvImp1(Errmsg)
                If flagOK Then
                    Call ShowTrainDesc() '顯示 訓練內容簡介
                Else
                    Common.MessageBox(Me, Errmsg)
                End If
            Case cst_flag_tims4768
                Dim flagOK As Boolean = sub_XLSImp1(Errmsg)
                If flagOK Then
                    Call ShowTrainDesc() '顯示 訓練內容簡介
                Else
                    Common.MessageBox(Me, Errmsg)
                End If
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
        End Select

        If (LayerState.Value = "") Then LayerState.Value = "10"
        Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js11)
    End Sub

    '匯入csv檔案
    Function Sub_CsvImp1(ByRef Errmsg As String) As Boolean
        '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim rst As Boolean = False '無誤 true 有誤 false

        Dim Reason As String = ""        '儲存錯誤的原因
        'Dim dtWrong As New DataTable    '儲存錯誤資料的DataTable
        'Dim drWrong As DataRow
        ''建立錯誤資料格式Table- -Start
        'dtWrong.Columns.Add(New DataColumn("Index"))
        'dtWrong.Columns.Add(New DataColumn("PName"))
        'dtWrong.Columns.Add(New DataColumn("IDNO"))
        'dtWrong.Columns.Add(New DataColumn("Reason"))
        ''建立錯誤資料格式Table- -End        

        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)

        Const cst_flag As String = ","
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        'Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        'Dim da As SqlDataAdapter = Nothing

        If File1.Value = "" Then
            Errmsg &= "請選擇匯入的檔案" & vbCrLf
            Return rst
        End If

        '檢查檔案格式與大小- -Start
        If File1.PostedFile.ContentLength = 0 Then
            'Common.MessageBox(Me, "檔案位置錯誤!")
            Errmsg &= "檔案位置錯誤!" & vbCrLf
            Return rst
            'Exit Function
        End If

        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Errmsg &= "檔案類型錯誤!" & vbCrLf
            Return rst
        End If

        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If Not LCase(MyFileType) = "csv" Then
            Errmsg &= "檔案類型錯誤，必須為CSV檔!" & vbCrLf
            Return rst
            'Exit Function
        End If
        '檢查檔案格式與大小- -End

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        '上傳檔案
        File1.PostedFile.SaveAs(Server.MapPath(Upload_Path & MyFileName))

        '將檔案讀出放入記憶體
        Dim sr As System.IO.Stream = IO.File.OpenRead(Server.MapPath(Upload_Path & MyFileName))
        Dim srr As System.IO.StreamReader = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

        Dim RowIndex As Integer = 0 '讀取行累計數
        Dim OneRow As String        'srr.ReadLine 一行一行的資料
        'Dim col As String           '欄位
        Dim colArray As Array

        Do While srr.Peek >= 0
            If Reason <> "" Then Exit Do
            OneRow = srr.ReadLine
            If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈
            If RowIndex <> 0 Then
                Reason = ""
                colArray = Split(OneRow, cst_flag)
                If Reason = "" Then Reason += CheckImportData(colArray) '檢查資料正確性
                If Reason = "" Then
                    Dim PName As String = colArray(0).ToString() '單元名稱
                    Dim PHour As String = colArray(1).ToString() '時數
                    Dim PCont As String = colArray(2).ToString() '課程大綱
                    'PCont = TIMS.ClearSQM(PCont)
                    If Session(Hid_TrainDesc_GUID1.Value) Is Nothing Then
                        Dim sql As String = ""
                        sql = "SELECT * FROM PLAN_TRAINDESC WHERE 1<>1"
                        dt = DbAccess.GetDataTable(sql, objconn)
                        dt.Columns(cst_PlanTrainDescPKName).AutoIncrement = True
                        dt.Columns(cst_PlanTrainDescPKName).AutoIncrementSeed = -1
                        dt.Columns(cst_PlanTrainDescPKName).AutoIncrementStep = -1
                    Else
                        dt = Session(Hid_TrainDesc_GUID1.Value)
                    End If
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("PName") = PName
                    dr("PHour") = PHour
                    dr("PCont") = TIMS.ClearSQM(PCont)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    'TIMS.Get_CourseCourseInfo(
                    Session(Hid_TrainDesc_GUID1.Value) = dt
                End If
            End If
            RowIndex += 1 '讀取行累計數
        Loop
        sr.Close()
        srr.Close()
        IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
        If Reason <> "" Then Errmsg = Reason '將error message input [Errmsg] 
        If Reason = "" Then rst = True '沒有任何的錯誤
        Return rst
    End Function

    '匯入XLS檔案
    Function sub_XLSImp1(ByRef Errmsg As String) As Boolean
        '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim rst As Boolean = False '無誤 true 有誤 false

        'Sub sub_XLSImp1()
        Const Cst_FileSavePath As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        Const cst_firstColumn As String = "單元名稱"
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""

        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing

        If File1.Value = "" Then
            'Common.MessageBox(Me, "未輸入匯入檔案位置")
            'Exit Function
            Errmsg &= "未輸入匯入檔案位置!" & vbCrLf
            Return rst
        End If
        If File1.PostedFile.ContentLength = 0 Then
            'Common.MessageBox(Me, "檔案位置錯誤!")
            'Exit Function
            Errmsg &= "檔案位置錯誤!" & vbCrLf
            Return rst
        End If

        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)

        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            'Common.MessageBox(Me, "檔案類型錯誤!")
            'Exit Function
            Errmsg &= "檔案類型錯誤!" & vbCrLf
            Return rst
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If MyFileType <> "xls" Then
            'Common.MessageBox(Me, "檔案類型錯誤，必須為XLS檔!")
            'Exit Function
            Errmsg &= "檔案類型錯誤，必須為XLS檔!" & vbCrLf
            Return rst
        End If

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        '上傳檔案
        File1.PostedFile.SaveAs(Server.MapPath(Cst_FileSavePath & MyFileName))

        '取得內容
        Dim Reason As String = "" '儲存錯誤的原因
        Dim fullFilNm1 As String = Server.MapPath(Cst_FileSavePath & MyFileName).ToString()
        Dim dt_xls As DataTable = TIMS.GetDataTable_XlsFile(fullFilNm1, "", Reason, cst_firstColumn)
        IO.File.Delete(Server.MapPath(Cst_FileSavePath & MyFileName)) '刪除檔案

        If Reason <> "" Then
            'Common.MessageBox(Me, Reason) 'Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝") 'Exit Function
            Errmsg &= "資料有誤，故無法匯入，請修正Excel檔案!" & vbCrLf
            Errmsg &= Reason & vbCrLf
            Return rst
        End If

        'xls 方式 讀取寫入資料庫
        If dt_xls.Rows.Count = 0 Then '有資料
            'Common.MessageBox(Me, "查無匯入資料!!") 'Exit Function
            Errmsg &= "查無匯入資料!!" & vbCrLf
            Return rst
        End If

        '有資料
        Reason = ""
        Dim iRowIndex As Integer = 1
        For i As Integer = 0 To dt_xls.Rows.Count - 1
            If iRowIndex <> 0 Then
                Dim colArray As Array = dt_xls.Rows(i).ItemArray
                Reason = CheckImportData(colArray)
                If Reason <> "" Then Exit For
                If Reason = "" Then
                    '無錯誤存檔 '匯入資料
                    Dim PName As String = colArray(0).ToString() '單元名稱
                    Dim PHour As String = colArray(1).ToString() '時數
                    Dim PCont As String = colArray(2).ToString() '課程大綱
                    'PCont = TIMS.ClearSQM(PCont)
                    If Session(Hid_TrainDesc_GUID1.Value) Is Nothing Then
                        Dim sql As String = ""
                        sql = "SELECT * FROM PLAN_TRAINDESC WHERE 1<>1"
                        dt = DbAccess.GetDataTable(sql, objconn)
                        dt.Columns(cst_PlanTrainDescPKName).AutoIncrement = True
                        dt.Columns(cst_PlanTrainDescPKName).AutoIncrementSeed = -1
                        dt.Columns(cst_PlanTrainDescPKName).AutoIncrementStep = -1
                    Else
                        dt = Session(Hid_TrainDesc_GUID1.Value)
                    End If
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("PName") = PName
                    dr("PHour") = PHour
                    dr("PCont") = TIMS.ClearSQM(PCont)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    'TIMS.Get_CourseCourseInfo(
                    Session(Hid_TrainDesc_GUID1.Value) = dt
                End If
            End If
            iRowIndex += 1
        Next

        If Reason <> "" Then Errmsg = Reason '將error message input [Errmsg] 
        If Reason = "" Then rst = True '沒有任何的錯誤
        Return rst
    End Function

    '產生新的GUID 避免記憶體相同 而異常
    Sub CREATE_NEW_GUID21()
        'If IsPostBack Then Exit Sub
        Hid_TrainDesc_GUID1.Value = TIMS.GetGUID()
        Session(Hid_TrainDesc_GUID1.Value) = Nothing
        'Const cst_TC03001_TRAINDESC_GUID1 As String = "TC03001_TRAINDESC_GUID1"
        'If Not Session(cst_TC03001_TRAINDESC_GUID1) Is Nothing Then
        '    '清理上一個SESSION GUID
        '    Dim TC03001_TRAINDESC_GUID1 As String = Session(cst_TC03001_TRAINDESC_GUID1)
        '    Session(TC03001_TRAINDESC_GUID1) = Nothing
        '    Session.Contents.Remove(TC03001_TRAINDESC_GUID1)
        'End If
        'Session(cst_TC03001_TRAINDESC_GUID1) = Hid_TrainDesc_GUID1.Value '記錄這一個SESSION GUID

        Hid_CostItem_GUID1.Value = TIMS.GetGUID()
        Session(Hid_CostItem_GUID1.Value) = Nothing
        'Const cst_TC03001_COSTITEM_GUID1 As String = "TC03001_COSTITEM_GUID1"
        'If Not Session(cst_TC03001_COSTITEM_GUID1) Is Nothing Then
        '    '清理上一個SESSION GUID
        '    Dim TC03001_COSTITEM_GUID1 As String = Session(cst_TC03001_COSTITEM_GUID1)
        '    Session(TC03001_COSTITEM_GUID1) = Nothing
        '    Session.Contents.Remove(TC03001_COSTITEM_GUID1)
        'End If
        ''記錄這一個SESSION GUID
        'Session(cst_TC03001_COSTITEM_GUID1) = Hid_CostItem_GUID1.Value
    End Sub

    ''' <summary>
    ''' 檢核-2019年啟用 work2019x01:2019 政府政策性產業-檢核
    ''' </summary>
    ''' <returns></returns>
    Function CHK_KID20_VAL() As String
        Dim Errmsg As String = ""
        '「5+2」產業創新計畫 5+2產業'【台灣AI行動計畫】 KID='08''【數位國家創新經濟發展方案】KID='09'
        '【國家資通安全發展方案】KID='10''【前瞻基礎建設計畫】'【新南向政策】KID='19'
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID20_1)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "「5+2」產業創新計畫，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_2)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【台灣AI行動計畫】，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_3)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【數位國家創新經濟發展方案】，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_4)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【國家資通安全發展方案】，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_5)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【前瞻基礎建設計畫】，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_6)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【新南向政策】，不可複選(僅可單一勾選)" & vbCrLf
        'tmp01 = TIMS.GetCblValue(CBLKID22)
        'If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【進階政策性產業類別】，不可複選(僅可單一勾選)" & vbCrLf
        Return Errmsg
    End Function

    ''' <summary>取值-2019年啟用 work2019x01:2019 政府政策性產業-取值</summary>
    ''' <returns></returns>
    Function GET_KID20_VAL() As String
        Dim rst As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID20_1)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_2)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_3)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_4)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_5)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_6)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        Return rst
    End Function

    ''' <summary> 檢核-2025年啟用-2025 政府政策性產業</summary>
    ''' <returns></returns>
    Function CHK_KID25_VAL() As String
        Dim Errmsg As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID25_1)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "亞洲矽谷，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_2)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "重點產業，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_3)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "台灣AI行動計畫，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_4)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "智慧國家方案，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_5)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "國家人才競爭力躍升方案，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_6)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "新南向政策，不可複選(僅可單一勾選)" & vbCrLf
        'tmp01 = TIMS.GetCblValue(CBLKID25_7)
        'If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "AI加值應用，不可複選(僅可單一勾選)" & vbCrLf
        'tmp01 = TIMS.GetCblValue(CBLKID25_8)
        'If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "職場續航，不可複選(僅可單一勾選)" & vbCrLf
        'tmp01 = TIMS.GetCblValue(CBLKID22B)
        'If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【進階政策性產業類別】，不可複選(僅可單一勾選)" & vbCrLf
        Return Errmsg
    End Function

    ''' <summary>取值-啟用-2025 政府政策性產業</summary>
    ''' <returns></returns>
    Function GET_KID25_VAL() As String
        Dim rst As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID25_1)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_2)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_3)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_4)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_5)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_6)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        'tmp01 = TIMS.GetCblValue(CBLKID25_7)
        'If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        'tmp01 = TIMS.GetCblValue(CBLKID25_8)
        'If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        Return rst
    End Function


    ''' <summary>
    ''' 計畫：接受企業委託計畫  (僅此計畫! 自辦不顯示喔)
    ''' </summary>
    ''' <param name="flag_TPlanID07_show"></param>
    ''' <param name="v_TPlanID"></param>
    ''' <returns></returns>
    Function GET_TRNUNIT1_LAB_LC2(ByRef flag_TPlanID07_show As Boolean, ByVal v_TPlanID As String) As String
        Dim rst As String = ""
        '計畫：接受企業委託計畫  (僅此計畫! 自辦不顯示喔)
        Const cst_受訓資格 As String = "受訓資格"
        Const cst_訓練對象及資格 As String = "訓練對象及資格"
        flag_TPlanID07_show = False
        rst = cst_受訓資格
        If TIMS.Cst_TPlanID07.IndexOf(v_TPlanID) > -1 Then
            flag_TPlanID07_show = True
            rst = cst_訓練對象及資格
        End If
        Return rst
    End Function

    ''' <summary>計畫：接受企業委託計畫  (僅此計畫! 自辦不顯示喔) Hid_TRNUNIT.Value = sm.UserInfo.TPlanID</summary>
    Sub SHOW_TRNUNIT1()
        Dim flag_TPlanID07_show As Boolean = False
        lab_LayerC2.Text = GET_TRNUNIT1_LAB_LC2(flag_TPlanID07_show, sm.UserInfo.TPlanID)
        Hid_TRNUNIT.Value = If(flag_TPlanID07_show, sm.UserInfo.TPlanID, "")

        trTRNUNITNAME.Visible = flag_TPlanID07_show 'False
        trTRNUNITTYPE.Visible = flag_TPlanID07_show 'False
        trTRNUNITEE.Visible = flag_TPlanID07_show 'False
    End Sub

    ''' <summary>專長能力標籤-ABILITY</summary>
    Private Sub SHOW_PLAN_ABILITYS()
        'If g_flagNG Then
        '    sm.LastErrorMessage = cst_errmsg3
        '    Exit Sub
        'End If
        'Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        'Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        'Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return 'rst 'Exit Sub

        Dim oParms As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}}
        Dim sSql As String = ""
        sSql &= " SELECT PABID,PLANID,COMIDNO,SEQNO,SEQ_ID,ABILITY,ABILITY_DESC"
        sSql &= " FROM PLAN_ABILITY WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        'sSql &= " ORDER BY SEQ_ID DESC" & vbCrLf
        Dim tb_dt As DataTable = DbAccess.GetDataTable(sSql, objconn, oParms)

        For Each tb_dr As DataRow In tb_dt.Rows
            Dim s_SEQ As String = Convert.ToString(tb_dr("SEQ_ID"))
            Dim otxtA1 As TextBox = If(s_SEQ = "1", txtABILITY1, If(s_SEQ = "2", txtABILITY2, If(s_SEQ = "3", txtABILITY3, If(s_SEQ = "4", txtABILITY4, Nothing))))
            Dim otxtA2 As TextBox = If(s_SEQ = "1", txtABILITY_DESC1, If(s_SEQ = "2", txtABILITY_DESC2, If(s_SEQ = "3", txtABILITY_DESC3, If(s_SEQ = "4", txtABILITY_DESC4, Nothing))))
            If otxtA1 IsNot Nothing Then otxtA1.Text = Convert.ToString(tb_dr("ABILITY"))
            If otxtA2 IsNot Nothing Then otxtA2.Text = Convert.ToString(tb_dr("ABILITY_DESC"))
        Next

    End Sub

    ''' <summary>專長能力標籤-ABILITY</summary>
    Sub SAVE_PLAN_ABILITY()
        'If upt_PlanX.Value = "" Then Return
        'tmpPCS = upt_PlanX.Value  '有儲存資料過了
        'PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
        'ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
        'SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
        'If upt_PlanX.Value = "" Then Exit Sub '無有效值離開
        Dim rqPlanID As String = PlanID_value 'TIMS.GetMyValue2(htSS, "rqPlanID")
        Dim rqComIDNO As String = ComIDNO_value 'TIMS.GetMyValue2(htSS, "rqComIDNO")
        Dim rqSeqNO As String = SeqNO_value 'TIMS.GetMyValue2(htSS, "rqSeqNO")
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return '(有異常離開)

        Dim iRst As Integer = 0
        For i_SEQ As Integer = 1 To 4
            Dim s_SEQ As String = Convert.ToString(i_SEQ)
            Dim otxtA1 As TextBox = If(s_SEQ = "1", txtABILITY1, If(s_SEQ = "2", txtABILITY2, If(s_SEQ = "3", txtABILITY3, If(s_SEQ = "4", txtABILITY4, Nothing))))
            Dim otxtA2 As TextBox = If(s_SEQ = "1", txtABILITY_DESC1, If(s_SEQ = "2", txtABILITY_DESC2, If(s_SEQ = "3", txtABILITY_DESC3, If(s_SEQ = "4", txtABILITY_DESC4, Nothing))))
            otxtA1.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA1.Text), 30)
            otxtA2.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA2.Text), 200)

            Dim pms_s1 As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"SEQ_ID", i_SEQ}}
            Dim sql_s1 As String = "SELECT PABID FROM PLAN_ABILITY WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SEQ_ID=@SEQ_ID"
            Dim dt As DataTable = DbAccess.GetDataTable(sql_s1, objconn, pms_s1)
            If otxtA1.Text <> "" Then
                If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                    Dim iPABID As Integer = DbAccess.GetNewId(objconn, "PLAN_ABILITY_PABID_SEQ,PLAN_ABILITY,PABID")
                    Dim iParms As New Hashtable From {{"PABID", iPABID},
                        {"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"SEQ_ID", i_SEQ},
                        {"ABILITY", otxtA1.Text}, {"ABILITY_DESC", otxtA2.Text}, {"MODIFYACCT", sm.UserInfo.UserID}}
                    Dim isSql As String = ""
                    isSql &= " INSERT INTO PLAN_ABILITY(PABID, PLANID, COMIDNO, SEQNO, SEQ_ID, ABILITY, ABILITY_DESC, MODIFYACCT, MODIFYDATE)" & vbCrLf
                    isSql &= " VALUES(@PABID,@PLANID,@COMIDNO,@SEQNO,@SEQ_ID,@ABILITY,@ABILITY_DESC,@MODIFYACCT,GETDATE())" & vbCrLf
                    iRst = DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
                Else
                    Dim iPABID As Integer = CInt(dt.Rows(0)("PABID"))
                    Dim uParms As New Hashtable From {{"PABID", iPABID},
                        {"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"SEQ_ID", i_SEQ},
                        {"ABILITY", otxtA1.Text}, {"ABILITY_DESC", otxtA2.Text}, {"MODIFYACCT", sm.UserInfo.UserID}
                    }
                    Dim usSql As String = ""
                    usSql &= " UPDATE PLAN_ABILITY" & vbCrLf
                    usSql &= " SET ABILITY=@ABILITY,ABILITY_DESC=@ABILITY_DESC,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                    usSql &= " WHERE PABID=@PABID AND PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SEQ_ID=@SEQ_ID" & vbCrLf
                    iRst = DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
                End If
            Else
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim pms_d1 As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"SEQ_ID", i_SEQ}}
                    Dim sql_d1 As String = "DELETE PLAN_ABILITY WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SEQ_ID=@SEQ_ID"
                    iRst = DbAccess.ExecuteNonQuery(sql_d1, objconn, pms_d1)
                End If
            End If

        Next

    End Sub

End Class
