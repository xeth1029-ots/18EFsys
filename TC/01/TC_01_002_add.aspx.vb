Partial Class TC_01_002_add
    Inherits AuthBasePage

    'Org_Apply / Org_orginfo / Auth_Relship 
    'Org_Comments

    Sub SUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("'ORG_ORGINFO','ORG_ORGPLANINFO','ID_PLAN','ID_DISTRICT','KEY_PLAN'", objconn)
        Call TIMS.sUtl_SetMaxLen(dt, "ORGNAME", TBtitle) '訓練機構
        Call TIMS.sUtl_SetMaxLen(dt, "COMIDNO", TBID) '統一編號
        Call TIMS.sUtl_SetMaxLen(dt, "COMCIDNO", TBseqno) '立案證號
        'Case "YEARS".ToUpper(), "NAME".ToUpper(), "PLANNAME".ToUpper(), "SEQ".ToUpper() '隸屬機構
        'Call TIMS.sUtl_SetMaxLen(dt, "YEARS", TBtitle) '隸屬機構
        Call TIMS.sUtl_SetMaxLen(dt, "ORGPNAME", TB_OrgPName) '分支單位名稱
        Call TIMS.sUtl_SetMaxLen(dt, "ADDRESS", TBaddress) '地址
        Call TIMS.sUtl_SetMaxLen(dt, "PLANMASTER", PlanMaster) '計畫主持人
        Call TIMS.sUtl_SetMaxLen(dt, "PLANMasterPhone", PlanMasterPhone) '主持人電話
        Call TIMS.sUtl_SetMaxLen(dt, "MASTERNAME", TBm_name) '負責人姓名
        Call TIMS.sUtl_SetMaxLen(dt, "MasterPhone", TBm_Phone) '負責人電話

        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTNAME", TBContactName) '聯絡人姓名
        Call TIMS.sUtl_SetMaxLen(dt, "PHONE", TBtel) '聯絡人電話
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTCELLPHONE", TBcontact_cellphone) '聯絡人行動電話

        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTEMAIL", TBmail) '聯絡人E-MAIL
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTFAX", ContactFax) '聯絡人傳真 
        Call TIMS.sUtl_SetMaxLen(dt, "STAFFNAME", staffName) '個人資料檔案保管人員 
        Call TIMS.sUtl_SetMaxLen(dt, "STAFFPHONE", staffPhone) '個人資料檔案保管人員電話 
        Call TIMS.sUtl_SetMaxLen(dt, "STAFFEMAIL", staffEmail) '個人資料檔案保管人員電子郵件 
        Call TIMS.sUtl_SetMaxLen(dt, "ACTNO", TB_ActNo) '保險證號
        Call TIMS.sUtl_SetMaxLen(dt, "TRAINCAP", TB_TrainCap) '訓練容量
        Call TIMS.sUtl_SetMaxLen(dt, "FIRECONTROLSTATE", TB_FireControlState) '消防安檢狀況
        Call TIMS.sUtl_SetMaxLen(dt, "PROTRAINKIND", TB_ProTrainKind) '專長訓練職類
        Call TIMS.sUtl_SetMaxLen(dt, "BANKNAME", BankName) '銀行(庫局)
        Call TIMS.sUtl_SetMaxLen(dt, "EXBANKNAME", ExBankName) '分行(支庫局)
        Call TIMS.sUtl_SetMaxLen(dt, "ACCNO", AccNo) '金融機構帳號
        Call TIMS.sUtl_SetMaxLen(dt, "ORGADDRESS", TBaddress_Org) '訓練機構屬性設定-立案地址/會址
    End Sub

    'Dim gFlagEnv As Boolean=True 'true:正式環境。(false:測試用) / TestStr
    'Dim gflag_test As Boolean=True 'true:測試環境。(false:正式環境) / TestStr
    'ProcessType : 修改(年度對應功能 ) Update/ 共用 Share /審核確認 InsertChk /新增 Insert
    Dim Rq_ProcessType As String

    Dim Re_orgid As String
    Dim Re_planid As String
    Dim Re_rid As String

    ''' <summary> 新增一筆 Rq_ProcessType </summary>
    Const cst_Insert As String = "Insert"
    ''' <summary> 修改(年度對應功能 ) Rq_ProcessType </summary>
    Const cst_Update As String = "Update"
    ''' <summary> 共用 Rq_ProcessType </summary>
    Const cst_Share As String = "Share"
    ''' <summary> 審核確認 Rq_ProcessType </summary>
    Const cst_InsertChk As String = "InsertChk"

    'Dim ProcessType2 As String
    'Dim sqlstr_id As String
    'Dim FunDr As DataRow
    Dim planid2 As String 'orglevel
    Dim maxrid As String 'max(rid) 
    Dim Parent_list As String

    Const cst_vs_sqlstr_list As String = "sqlstr_list"
    Const vs_RSID As String = "RSID"
    Const vs_Redirect As String = "Redirect"
    Const vs_comidno As String = "comidno"
    Const vs_sort As String = "sort"
    'Const flag_TPlanID17_Auth_AccRWPlan As String="flag_TPlanID17_Auth_AccRWPlan"
    'Const flag_TPlanID17_plan_planinfo As String="flag_TPlanID17_plan_planinfo"

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
        If Session("_Search") IsNot Nothing Then Session("_Search") = Session("_Search")
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        Call SUtl_PageInit1()
        '檢查Session是否存在 End

        '200808 andy 機構名稱限制只能經由 BT_CHGORG(變更)按鈕來更改
        'TBtitle.Enabled=False
        Rq_ProcessType = Convert.ToString(Request("ProcessType"))
        TBtitle.Enabled = (Rq_ProcessType = cst_Insert)

        city_code.Attributes.Add("onblur", "getZipName('TBCity',this,this.value);")
        city_code_org.Attributes.Add("onblur", "getZipName('TBCity_Org',this,this.value);")

        If Not Page.IsPostBack Then
            Call Check_TPlanID28(sm.UserInfo.TPlanID)
            '判斷登入角色若為
            '1.產業人才投資計畫角色登入者則為唯讀
            '2.提升勞工自主學習計畫角色登入者則為可以使用
        End If

        ViewState(vs_RSID) = "" '接收值
        Re_orgid = TIMS.ClearSQM(Request("orgid"))
        If Request("RSID") IsNot Nothing Then ViewState(vs_RSID) = TIMS.ClearSQM(Request("RSID"))

        If Re_orgid <> "" Then
            OrgIDValue.Value = Re_orgid
            Call Set_GWOrgKind2(Re_orgid)
        Else
            If sm.UserInfo.LID = "2" Then Call Set_GWOrgKind2(sm.UserInfo.OrgID)
        End If
        '20080828 andy 變更機構名稱功能限制只在修改時才可使用
        If Rq_ProcessType <> cst_Update Then BT_CHGORG.Enabled = False

        Re_planid = TIMS.ClearSQM(Request("planid"))
        Re_rid = TIMS.ClearSQM(Request("rid"))

        If sm.UserInfo.FunDt Is Nothing Then
            Common.RespWrite(Me, "<script>alert('Session過期');</script>")
            Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        End If

        MenuTable.Style.Item("display") = "none"

        If Not Page.IsPostBack Then
            CCreate1()
        End If

        'Button1.Attributes("onclick")="history.go(-1);"
        LastYearExeRate.Attributes("onclick") = "SetLastYearExeRate();"

        '20090424 by jimmy add 3+2郵遞區號查詢 link --- start
        '20090424 by jimmy add 3+2郵遞區號查詢 link --- end

        Dim rqACTION As String = TIMS.ClearSQM(Request("ACTION"))
        If rqACTION <> "" Then
            Select Case rqACTION'Request("ACTION")
                Case "VIEW"
                    MenuTable.Style.Item("display") = "inline"
                    HistoryTable.Style.Item("display") = "none"
                    BT_CHGORG.Visible = False
                    Me.bt_save.Visible = False
                    If Session("Redirect") Is Nothing Then
                        Me.Button1.Visible = False
                    Else
                        Me.ViewState(vs_Redirect) = Session("Redirect")
                        Session("Redirect") = Nothing
                        Me.Button1.Visible = True
                    End If
                    Bt1_city_zip.Visible = False
                    Table1.Disabled = True
                    level_list.Enabled = False
                    OrgKindList.Enabled = False
                    ExeRate.Visible = False
                    choice_button.Visible = False
                    btn_clear.Visible = False
                    'Me.lblProecessType.Text="檢視"
                    HistoryTable.Disabled = True
                    '顯示開班歷史。
                    SearchHistory(Me.TBID.Text)
            End Select
        End If
    End Sub

    Sub CCreate1()
        Dim rqOrgName As String = TIMS.ClearSQM(Request("OrgName"))
        Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim Re_distid As String = TIMS.ClearSQM(Request("distid"))

        '(訓練機構屬性設定)-郵遞區號查詢
        LitZipCODE.Text = TIMS.Get_WorkZIPB3Link2()
        LitZipCODEOrg.Text = TIMS.Get_WorkZIPB3Link2()
        Dim bt1_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code, ZipCODEB3, hidZipCODE6W, TBCity, hidZipCODEB3_N, TBaddress)
        Bt1_city_zip.Attributes.Add("onclick", bt1_Attr_VAL)
        Dim bt2_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code_org, ZipCODEB3_Org, hidZipCODE6W_Org, TBCity_Org, hidZipCODEB3_Org_N, TBaddress_Org)
        Bt2_city_zip_org.Attributes.Add("onclick", bt2_Attr_VAL)

        'Dim parms As Hashtable=New Hashtable()
        'If Not Page.IsPostBack Then
        Page.RegisterStartupScript("Load", "<script>SetLastYearExeRate();</script>")
        If Session("_Search") IsNot Nothing Then Session("_Search") = Session("_Search")

        'Call level_list_SelectedIndexChanged(sender, e)
        Call SUtl_levellistSel()

        Select Case Rq_ProcessType
            Case cst_Insert
                'Me.lblProecessType.Text="新增"
                Table3.Visible = False ''機構年度評鑑資料
                '20100208 按新增時代查詢之 機構名稱 & 統一編號
                If rqOrgName <> "" Then TBtitle.Text = rqOrgName
                If rqComIDNO <> "" Then
                    TBID.Text = rqComIDNO
                    HidComidno.Value = TIMS.Chg_Subst8(TBID.Text)
                End If
            Case cst_Update
                'Me.lblProecessType.Text="修改"
                Table3.Visible = True '機構年度評鑑資料
            Case cst_Share
                'Me.lblProecessType.Text="共用"
                Table3.Visible = True '機構年度評鑑資料
            Case cst_InsertChk
                'Me.lblProecessType.Text="審核"
                Table3.Visible = False ''機構年度評鑑資料
            Case Else
                'Me.lblProecessType.Text="(開發檢視中) 情況不明-請將操作步驟告知系統管理者"
                Table3.Visible = False ''機構年度評鑑資料
        End Select

        'Dim roleid As String=sm.UserInfo.RoleID '角色權限。
        DistrictList = TIMS.Get_DistID(DistrictList)
        DistrictList.Items.Remove(DistrictList.Items.FindByValue(""))

        If Not (Rq_ProcessType = cst_Share) Then
            Select Case sm.UserInfo.LID
                Case 0
                    level_list.Items.Insert(0, New ListItem("分署", "1"))
                    '署 '非共用'非新增
                    If Not (Rq_ProcessType = cst_Insert) Then level_list.Items.Insert(0, New ListItem("署", "0"))  '轄區是署
                Case 1 '分署 
                    '分署 '非共用 '非新增
                    If Not (Rq_ProcessType = cst_Insert) Then level_list.Items.Insert(0, New ListItem("分署", "1"))
            End Select
        End If

        'sm.UserInfo.RoleID '角色權限。
        Common.SetListItem(DistrictList, sm.UserInfo.DistID)
        Common.SetListItem(level_list, sm.UserInfo.LID)

        If Re_distid <> "" Then
            If Re_distid.Equals("000") AndAlso sm.UserInfo.LID = 0 Then
                Common.SetListItem(DistrictList, Re_distid)
                DistrictList.Enabled = False
                TIMS.Tooltip(DistrictList, "轄區是署,鎖定")

                level_list.Enabled = False
                TIMS.Tooltip(level_list, "轄區是署,鎖定")
            End If
        End If

        Select Case Convert.ToString(sm.UserInfo.RoleID)
            Case "-1" '目前無此資料
            Case "0" '表示為 超級管理者
                'Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                If sm.UserInfo.DistID <> "000" Then
                    DistrictList.Enabled = False
                    TIMS.Tooltip(DistrictList, "管理者權限")
                End If

            Case "1" '角色為系統管理者。階層為分署(中心)
                If Rq_ProcessType = cst_Update Then
                    level_list.Enabled = False
                    TIMS.Tooltip(level_list, "轄區是分署")
                End If
                'Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False
                TIMS.Tooltip(DistrictList, "分署的權限")

            Case Else '角色為非系統管理者。'階層為分署(中心)以下
                level_list.Enabled = False
                TIMS.Tooltip(level_list, "委訓單位的權限")

                'Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False
                TIMS.Tooltip(DistrictList, "委訓單位的權限")

        End Select

        'If roleid="0" Then '表示為超級管理者
        'ElseIf roleid="1" Then '角色為系統管理者。階層為分署(中心)
        'ElseIf roleid > "1" Then '角色為非系統管理者。'階層為分署(中心)以下
        'End If

        OrgKindList = TIMS.Get_OrgType(OrgKindList, objconn)

        'Dim City As String=""
        'Dim planid As String=""
        If Rq_ProcessType = cst_Update Then
            TBplan.Enabled = False
            TIMS.Tooltip(TBplan, "不能修改~")
            'Dim sqlstr_classid, sql_orglevel As String

            Dim cParms As New Hashtable From {{"ORGID", Re_orgid}, {"PlanID", Re_planid}, {"RID", Re_rid}}
            Dim sqlstr_classid As String = ""
            sqlstr_classid &= " SELECT 'X'" & vbCrLf
            sqlstr_classid &= " FROM ORG_ORGINFO a" & vbCrLf
            sqlstr_classid &= " JOIN AUTH_RELSHIP b ON a.orgid=b.orgid" & vbCrLf
            sqlstr_classid &= " JOIN PLAN_PLANINFO c ON b.PlanID=c.PlanID AND b.RID=c.RID" & vbCrLf
            sqlstr_classid &= " WHERE a.ORGID=@ORGID AND b.PlanID=@PlanID AND c.RID=@RID" & vbCrLf
            Dim dtOAP As DataTable = DbAccess.GetDataTable(sqlstr_classid, objconn, cParms)
            If dtOAP.Rows.Count > 0 Then
                'If Not Convert.IsDBNull(classid_list("OCID")) Then '有開課,不能修改
                Me.level_list.Enabled = False
                Me.TBplan.Enabled = False
                Me.choice_button.Disabled = True '變更隸屬機構
                Me.DistrictList.Enabled = False
                TIMS.Tooltip(level_list, "有開課,不能修改~")
                TIMS.Tooltip(TBplan, "有開課,不能修改~")
                TIMS.Tooltip(choice_button, "有開課,不能修改~") '變更隸屬機構
                TIMS.Tooltip(DistrictList, "有開課,不能修改~")
                TBID.Enabled = False
                TIMS.Tooltip(TBID, "有開課,不能修改~")
            End If

            choice_button.Disabled = True '變更隸屬機構
            TIMS.Tooltip(choice_button, "不能修改~")

            '補助地方政府訓練  開放更改隸屬機構
            'Select Case sm.UserInfo.TPlanID
            '    Case "17"
            '        'Rq_ProcessType=cst_Update 
            '        'sm.UserInfo.TPlanID="17"
            '        Me.ViewState(flag_TPlanID17_Auth_AccRWPlan)=False
            '        Me.ViewState(flag_TPlanID17_plan_planinfo)=False
            '        '有開共用權限 Auth_AccRWPlan
            '        Dim sql As String=""
            '        sql="" & vbCrLf
            '        sql=" SELECT COUNT(1) cnt FROM Auth_AccRWPlan WHERE planid='" & Re_planid & "' AND RID='" & Re_rid & "' "
            '        If CInt(DbAccess.ExecuteScalar(sql, objconn)) > 0 Then Me.ViewState(flag_TPlanID17_Auth_AccRWPlan)=True
            '        '此機構已有開班計畫
            '        If Not Me.choice_button.Disabled Then '可變更隸屬機構，再次確認此機構已有開班計畫
            '            sql="" & vbCrLf
            '            sql &= " SELECT count(1) cnt" & vbCrLf
            '            sql &= " FROM Org_orginfo a" & vbCrLf
            '            sql &= " JOIN Auth_Relship b ON a.orgid=b.orgid" & vbCrLf
            '            sql &= " JOIN plan_planinfo c ON b.PlanID=c.PlanID AND a.comidno=c.comidno" & vbCrLf
            '            sql &= " WHERE a.orgid='" & Re_orgid & "'" & vbCrLf
            '            sql &= " AND b.PlanID='" & Re_planid & "'" & vbCrLf
            '            If CInt(DbAccess.ExecuteScalar(sql, objconn)) > 0 Then
            '                Me.ViewState(flag_TPlanID17_plan_planinfo)=True
            '                TIMS.Tooltip(level_list, "此機構已有開班計畫")
            '                TIMS.Tooltip(TBplan, "此機構已有開班計畫")
            '                TIMS.Tooltip(choice_button, "此機構已有開班計畫") '變更隸屬機構
            '                TIMS.Tooltip(DistrictList, "此機構已有開班計畫")
            '                TIMS.Tooltip(TBID, "此機構已有開班計畫")
            '            End If
            '        End If
            '    Case Else
            '        choice_button.Disabled=True '變更隸屬機構
            '        TIMS.Tooltip(choice_button, "不能修改~")
            'End Select

            'Dim sql_orglevel As String=""
            'sql_orglevel="" & vbCrLf
            'sql_orglevel &= " SELECT b.orglevel" & vbCrLf
            'sql_orglevel &= " FROM Org_orginfo a" & vbCrLf
            'sql_orglevel &= " JOIN Auth_Relship b ON a.orgid=b.orgid" & vbCrLf
            'sql_orglevel &= " WHERE a.orgid='" & Re_orgid & "'" & vbCrLf
            'sql_orglevel &= " AND b.PlanID='" & Re_planid & "'" & vbCrLf
            'sql_orglevel &= " AND b.RID='" & Re_rid & "'" & vbCrLf
            'Dim classid_list As String=Convert.ToString(DbAccess.ExecuteScalar(sql_orglevel, objconn))

            'os_parms.Clear()
            Dim os_parms As New Hashtable From {{"orgid", Re_orgid}, {"PlanID", Re_planid}, {"RID", Re_rid}}
            '機構階層【0:署(職訓局) 1:分署(中心) 2:委訓(補助單位) 3:(委訓)】
            Dim i_orglevel As Integer = Get_orglevel(os_parms)
            If i_orglevel > 1 Then
                Common.SetListItem(level_list, "2")
                level_list.Enabled = False
                TIMS.Tooltip(level_list, "業務權限檢核機構階層 為委訓單位")
            ElseIf i_orglevel = 1 Then
                Common.SetListItem(level_list, "1")
                level_list.Enabled = False
                TIMS.Tooltip(level_list, "業務權限檢核機構階層 為分署")
                DistrictList.Enabled = False
                TIMS.Tooltip(DistrictList, "業務權限檢核機構階層 為分署")
            End If

            '計算共用數量
            Dim check_auth As String = "SELEC * FROM AUTH_RELSHIP WHERE ORGID=" & Re_orgid
            If DbAccess.GetCount(check_auth, objconn) > 1 Then '有共用過,訓練機構共同資料不能修改
                DistrictList.Enabled = False
                TBID.Enabled = False
                TIMS.Tooltip(DistrictList, "有共用過,訓練機構共同資料不能修改")
                TIMS.Tooltip(TBID, "有共用過,訓練機構共同資料不能修改")
            End If
            '點選共用按鈕,只能更改計畫階段and 訓練機構承辦人資料 =============
        ElseIf Rq_ProcessType = cst_Share Then
            Me.TBtitle.Enabled = False
            Me.TBID.Enabled = False
            Me.TBseqno.Enabled = False
            Me.OrgKindList.Enabled = False
            TIMS.Tooltip(TBtitle, "共用機構資訊,不能修改~")
            TIMS.Tooltip(TBID, "共用機構資訊,不能修改~")
            TIMS.Tooltip(TBseqno, "共用機構資訊,不能修改~")
            TIMS.Tooltip(OrgKindList, "共用機構資訊,不能修改~")
        End If
        btn_clear.Disabled = choice_button.Disabled
        '=================================

        Select Case Rq_ProcessType
            Case cst_InsertChk, cst_Update, cst_Share
                Dim list As DataRow = Nothing

                If Rq_ProcessType = cst_InsertChk Then

                    Dim parms_list As New Hashtable From {{"comIDNO", rqComIDNO}}
                    Dim sql_list As String = ""
                    sql_list &= " SELECT a.ComIDNO ,a.OrgKind ,a.OrgName ,a.ComCIDNO ,a.IsConUnit ,a.OrgPName ,a.ZipCode ,a.Address ,a.Phone" & vbCrLf
                    sql_list &= " ,a.MasterName,a.MasterPhone" & vbCrLf
                    sql_list &= " ,a.ContactName ,a.ContactEmail ,a.ContactCellPhone ,a.TrainCap ,a.ProTrainKind ,a.FireControlState ,a.ComSumm" & vbCrLf
                    sql_list &= " ,a.ActNo ,a.PlanMaster" & vbCrLf
                    sql_list &= " ,a.PlanMasterPhone ,a.ContactFax ,a.staffName ,a.staffPhone ,a.staffEmail ,a.PlanID ,a.Result ,a.Note" & vbCrLf
                    sql_list &= " FROM dbo.ORG_APPLY a" & vbCrLf
                    sql_list &= " WHERE a.comIDNO=@comIDNO" & vbCrLf
                    list = DbAccess.GetOneRow(sql_list, objconn, parms_list)
                    TBID.Enabled = False
                    If list IsNot Nothing Then
                        Dim sTPlanID As String = TIMS.GetTPlanID(list("PlanID"), objconn)
                        Call Check_TPlanID28(sTPlanID)
                    End If
                End If
                If Rq_ProcessType = cst_Update OrElse Rq_ProcessType = cst_Share Then
                    Dim parms_list As New Hashtable From {{"ORGID", Re_orgid}, {"PlanID", Re_planid}, {"RID", Re_rid}}
                    Dim sql_list As String = ""
                    sql_list &= " SELECT a.OrgID ,a.OrgKind ,a.OrgName ,a.ComIDNO ,a.ComCIDNO ,a.IsConUnit ,a.TradeID ,a.EmpNum ,a.OrgUrl ,a.OrgKind2" & vbCrLf
                    sql_list &= " ,a.LastYearExeRate ,a.IsConTTQS ,a.BankName ,a.ExBankName ,a.AccNo ,a.AccName ,a.OrgKind1" & vbCrLf
                    sql_list &= " ,a.OrgZipCode,a.OrgZipCODE6W,a.OrgAddress" & vbCrLf 'ORG_ORGINFO a
                    sql_list &= " ,b.RSID ,b.PlanID ,b.RID ,b.Relship ,b.OrgLevel ,b.DistID" & vbCrLf 'AUTH_RELSHIP b
                    sql_list &= " ,c.OrgPName,c.ZipCode,c.ZipCODE6W,c.Address" & vbCrLf 'ORG_ORGPLANINFO c 
                    sql_list &= " ,c.Phone ,c.MasterName,c.MasterPhone ,c.ContactName ,c.ContactEmail ,c.ContactCellPhone ,c.TrainCap ,c.ProTrainKind" & vbCrLf
                    sql_list &= " ,c.FireControlState ,c.ComSumm ,c.ActNo ,c.ModifyAcct ,c.ModifyDate ,c.PlanMaster ,c.PlanMasterPhone ,c.ContactFax" & vbCrLf
                    sql_list &= " ,c.staffName ,c.staffPhone ,c.staffEmail ,c.ContactSex ,c.ContactTitle ,c.PayTax ,c.AssistUnit ,c.AssistUnit01" & vbCrLf
                    sql_list &= " ,c.AssistUnit02 ,c.AssistUnit03 ,c.AssistUnitOther  ,c.eComment ,c.MemberNum" & vbCrLf
                    sql_list &= " ,c.ActMemberNum ,c.ActStaffNum ,c.Accessible ,c.Textbook ,c.TeachAids ,c.HumanHelp ,e.zipname ,f.ctname" & vbCrLf
                    sql_list &= " FROM ORG_ORGINFO a" & vbCrLf
                    sql_list &= " JOIN AUTH_RELSHIP b ON a.orgid=b.orgid" & vbCrLf
                    sql_list &= " JOIN ORG_ORGPLANINFO c ON c.RSID=b.RSID" & vbCrLf
                    sql_list &= " LEFT JOIN ID_ZIP e ON e.zipcode=a.orgzipcode" & vbCrLf
                    sql_list &= " LEFT JOIN ID_CITY f ON f.ctid=e.ctid" & vbCrLf
                    sql_list &= " WHERE a.ORGID =@ORGID AND b.PlanID=@PlanID AND b.RID=@RID" & vbCrLf
                    ViewState(cst_vs_sqlstr_list) = sql_list
                    list = DbAccess.GetOneRow(sql_list, objconn, parms_list)
                    If list Is Nothing Then
                        bt_save.Enabled = False
                        bt_save.Visible = False
                        Dim Errmsg As String = ""
                        Errmsg += vbCrLf & "計畫權限取得有誤，請重新輸入查詢值!!"
                        Common.MessageBox(Me, Errmsg)
                        Return 'Exit Sub
                    End If

                    Dim sTPlanID As String = TIMS.GetTPlanID(list("PlanID"), objconn)
                    Call Check_TPlanID28(sTPlanID)
                    '2018 add:merge TC_01_017 資料修改介面-載入訓練機構屬性設定結果資訊(產投 & 充飛計畫用)
                    If sTPlanID <> "" AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                        Call SetOrgKind1(Convert.ToString(list("orgkind1"))) '計畫別 & 機構別
                        '立案地址/會址
                        city_code_org.Value = Convert.ToString(list("orgzipcode"))
                        hidZipCODE6W_Org.Value = Convert.ToString(list("orgzipcode6W"))
                        ZipCODEB3_Org.Value = TIMS.GetZIPCODEB3(hidZipCODE6W_Org.Value)
                        TBCity_Org.Text = String.Concat("(", Convert.ToString(list("orgzipcode")), ")", TIMS.Get_ZipNM(list("ctname"), list("zipname")))
                        TBaddress_Org.Text = Convert.ToString(list("orgaddress"))
                        'AndAlso sm.UserInfo.LID > 0
                        If Rq_ProcessType = cst_Update Then
                            '地址有資料
                            Dim fgAddressHaveData As Boolean = (city_code_org.Value <> "" AndAlso ZipCODEB3_Org.Value <> "" AndAlso TBaddress_Org.Text <> "")
                            If fgAddressHaveData AndAlso sm.UserInfo.LID > 0 Then
                                dl_typeid1.Enabled = False
                                dl_typeid2.Enabled = False
                                TIMS.Display_None(Bt2_city_zip_org)
                                city_code_org.Disabled = True
                                hidZipCODE6W_Org.Disabled = True
                                ZipCODEB3_Org.Disabled = True
                                TBCity_Org.Enabled = False
                                TBaddress_Org.Enabled = False
                                Dim tit1 As String = "不提供分署修改"
                                TIMS.Tooltip(dl_typeid1, tit1)
                                TIMS.Tooltip(dl_typeid2, tit1)
                                TIMS.Tooltip(city_code_org, tit1)
                                TIMS.Tooltip(hidZipCODE6W_Org, tit1)
                                TIMS.Tooltip(ZipCODEB3_Org, tit1)
                                TIMS.Tooltip(TBCity_Org, tit1)
                                TIMS.Tooltip(TBaddress_Org, tit1)
                            End If
                        End If
                    End If

                    Dim myarray As String() = list("Relship").Split("/")
                    Dim range As Integer = myarray.Length - 3
                    Parent_list = If(range > 0, myarray(range), myarray(0))
                    '共享時不配入預設的RID值-----------------------Modify-Chris
                    If Rq_ProcessType <> cst_Share Then RIDValue.Value = Parent_list
                    '共享時不配入預設的RID值-----------------------Modify-Chris

                    '是否為管控單位 1:是 0:否 '補助計劃
                    Dim v_IsConUnit As String = TIMS.ClearSQM(list("IsConUnit"))
                    If (v_IsConUnit = "") Then v_IsConUnit = "0"
                    Common.SetListItem(IsConUnit, v_IsConUnit)
                End If

                If list("OrgLevel").ToString() <> "" Then
                    Common.SetListItem(level_list, list("OrgLevel"))
                    level_list.Enabled = False
                    TIMS.Tooltip(level_list, "相同的機構層級")
                End If

                OrgIDValue.Value = If(Convert.ToString(list("OrgID")) <> "", Convert.ToString(list("OrgID")), OrgIDValue.Value)
                'planid=list("PlanID")
                PlanIDValue.Value = list("PlanID")
                TBtitle.Text = list("OrgName")
                TBID.Text = list("ComIDNO")
                HidComidno.Value = TIMS.Chg_Subst8(TBID.Text)
                ViewState(vs_comidno) = TBID.Text
                TBseqno.Text = list("ComCIDNO")
                TB_ActNo.Text = Convert.ToString(list("ActNo"))

                city_code.Value = Convert.ToString(list("ZipCode"))
                hidZipCODE6W.Value = Convert.ToString(list("ZipCode6W"))
                ZipCODEB3.Value = TIMS.GetZIPCODEB3(hidZipCODE6W.Value)
                'TBCity.Text=TIMS.GET_FullCCTName(objconn, city_code.Value, ZipCODEB3.Value)
                TBaddress.Text = Convert.ToString(list("Address"))

                TBseqno.Text = Convert.ToString(list("ComCIDNO"))
                Me.TBtel.Text = Convert.ToString(list("Phone"))
                Me.TBm_name.Text = Convert.ToString(list("MasterName"))
                Me.TBm_Phone.Text = Convert.ToString(list("MasterPhone"))

                Me.TBContactName.Text = Convert.ToString(list("ContactName"))
                Me.TBmail.Text = Convert.ToString(list("ContactEmail"))
                Me.TBcontact_cellphone.Text = Convert.ToString(list("ContactCellPhone"))
                Me.TB_TrainCap.Text = Convert.ToString(list("TrainCap"))
                Me.TB_FireControlState.Text = Convert.ToString(list("FireControlState"))
                Me.TB_ProTrainKind.Text = Convert.ToString(list("ProTrainKind"))
                Me.ComSumm.Text = Convert.ToString(list("ComSumm"))

                Select Case Rq_ProcessType
                    Case cst_Update
                        If Convert.ToString(list("DistID")) <> "" Then Common.SetListItem(DistrictList, Convert.ToString(list("DistID")))
                    Case cst_Share
                        If Convert.ToString(sm.UserInfo.DistID) <> "" Then Common.SetListItem(DistrictList, Convert.ToString(sm.UserInfo.DistID))
                End Select

                If Convert.ToString(list("OrgKind")) <> "" Then Common.SetListItem(OrgKindList, Convert.ToString(list("OrgKind")))
                TB_OrgPName.Text = Convert.ToString(list("OrgPName"))
                If Rq_ProcessType <> cst_Share Then
                    PlanMaster.Text = list("PlanMaster").ToString
                    PlanMasterPhone.Text = list("PlanMasterPhone").ToString
                End If
                ContactFax.Text = list("ContactFax").ToString
                staffName.Text = list("staffName").ToString
                staffPhone.Text = list("staffPhone").ToString
                staffEmail.Text = list("staffEmail").ToString
                MemberNum.Text = list("MemberNum").ToString
                ActMemberNum.Text = list("ActMemberNum").ToString
                ActStaffNum.Text = list("ActStaffNum").ToString

                Common.SetListItem(Accessible, "NG")
                Common.SetListItem(Textbook, "NG")
                Common.SetListItem(TeachAids, "NG")
                Common.SetListItem(HumanHelp, "NG")
                If Convert.ToString(list("Accessible")) <> "" Then Common.SetListItem(Accessible, list("Accessible"))
                If Convert.ToString(list("Textbook")) <> "" Then Common.SetListItem(Textbook, list("Textbook"))
                If Convert.ToString(list("TeachAids")) <> "" Then Common.SetListItem(TeachAids, list("TeachAids"))
                If Convert.ToString(list("HumanHelp")) <> "" Then Common.SetListItem(HumanHelp, list("HumanHelp"))

                If list("LastYearExeRate").ToString < 0 Then '上年度是否辦理本計劃
                    Common.SetListItem(LastYearExeRate, "-1") '否
                Else
                    Common.SetListItem(LastYearExeRate, "1") '是
                    txtLastYearExeRate.Text = list("LastYearExeRate").ToString
                End If
                'If Convert.ToString(list("LastYearExeRate")) <> "" Then
                '    If list("LastYearExeRate").ToString < 0 Then '上年度是否辦理本計劃
                '        Common.SetListItem(LastYearExeRate, "-1") '否
                '    Else
                '        Common.SetListItem(LastYearExeRate, "1") '是
                '        txtLastYearExeRate.Text=list("LastYearExeRate").ToString
                '    End If
                'Else
                '    Common.SetListItem(LastYearExeRate, "1") '是
                '    txtLastYearExeRate.Text=list("LastYearExeRate").ToString
                'End If

                Common.SetListItem(IsConTTQS, list("IsConTTQS").ToString) '通過TTQS
                BankName.Text = list("BankName").ToString
                ExBankName.Text = list("ExBankName").ToString
                AccNo.Text = list("AccNo").ToString
                AccName.Text = list("AccName").ToString
                Call Set_GWOrgKind2(list("OrgID").ToString)
        End Select

        Select Case Rq_ProcessType
            Case cst_Update, cst_Share
                Dim sqlstr_orgname As String = " SELECT b.OrgName PlanName FROM AUTH_RELSHIP a JOIN org_orginfo b ON a.orgid=b.orgid WHERE a.RID='" & Parent_list & "' "
                '顯示管控單位名稱
                Dim orgName As String = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_orgname, objconn))
                'Dim sqlstr_PlanName As String=" SELECT c.Years + d.Name + e.PlanName + c.seq + '_' AS PlanName FROM AUTH_RELSHIP a JOIN org_orginfo b ON a.orgid=b.orgid JOIN ID_Plan c ON c.PlanID=a.PlanID "
                Dim sqlstr_PlanName As String = ""
                sqlstr_PlanName &= " SELECT concat(c.Years,d.Name,e.PlanName,c.seq,'_') PlanName" & vbCrLf
                sqlstr_PlanName &= " FROM AUTH_RELSHIP a" & vbCrLf
                sqlstr_PlanName &= " JOIN ORG_ORGINFO b ON a.orgid=b.orgid" & vbCrLf
                sqlstr_PlanName &= " JOIN ID_Plan c ON c.PlanID=a.PlanID" & vbCrLf
                sqlstr_PlanName &= " JOIN ID_District d ON d.DistID=c.DistID" & vbCrLf
                sqlstr_PlanName &= " JOIN Key_Plan e ON c.TPlanID=e.TPlanID" & vbCrLf
                sqlstr_PlanName &= " WHERE a.planid='" & PlanIDValue.Value & "' AND a.RID='" & Re_rid & "'"
                '顯示管控單位計畫
                Dim plan_name As String = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_PlanName, objconn))
                Me.TBplan.Text = String.Concat(plan_name, orgName)

                Dim sql_s1 As String = ""
                Dim parms_s1 As New Hashtable From {{"ORGID", Re_orgid}, {"YEARS", CStr(sm.UserInfo.Years)}} ' .Clear()
                sql_s1 = " SELECT * FROM ORG_COMMENTS WHERE ORGID =@ORGID AND YEARS =@YEARS"
                Dim dr As DataRow = DbAccess.GetOneRow(sql_s1, objconn, parms_s1)
                If dr Is Nothing Then
                    Me.Table3.Visible = False
                Else
                    Me.LYears.Text = sm.UserInfo.Years.ToString
                    Me.Point01A.Text = dr("Point01A").ToString
                    Me.Point01B.Text = dr("Point01B").ToString
                    Me.Point02A.Text = dr("Point02A").ToString
                    Me.Point02B.Text = dr("Point02B").ToString
                    Me.Point03A.Text = dr("Point03A").ToString
                    Me.Point03B.Text = dr("Point03B").ToString
                End If
        End Select

        If (hidZipCODE6W.Value <> String.Concat(city_code.Value, ZipCODEB3.Value) AndAlso city_code.Value <> "" AndAlso ZipCODEB3.Value <> "") Then hidZipCODE6W.Value = TIMS.GetZIPCODE6W(city_code.Value, ZipCODEB3.Value)
        If city_code.Value <> "" Then TBCity.Text = TIMS.GET_FullCCTName(objconn, city_code.Value, hidZipCODE6W.Value)
        'End If
    End Sub

    ''' <summary>
    '''  取得機構層級 orglevel  
    '''  機構階層【0:署(職訓局) 1:分署(中心) 2:委訓(補助單位) 3:(委訓)】
    ''' </summary>
    ''' <returns></returns>
    Function Get_orglevel(ByRef s_parms As Hashtable) As Integer
        Dim rst As Integer = 3

        Dim sql_orglevel As String = ""
        sql_orglevel &= " SELECT b.orglevel" & vbCrLf
        sql_orglevel &= " FROM Org_orginfo a" & vbCrLf
        sql_orglevel &= " JOIN Auth_Relship b ON a.orgid=b.orgid" & vbCrLf
        sql_orglevel &= " WHERE a.orgid=@orgid" & vbCrLf
        sql_orglevel &= " AND b.PlanID=@PlanID" & vbCrLf
        sql_orglevel &= " AND b.RID=@RID" & vbCrLf

        Dim dt1 As New DataTable

        Dim sCmd As New SqlCommand(sql_orglevel, objconn)

        DbAccess.HashParmsChange(sCmd, s_parms)

        DbAccess.Open(objconn)

        dt1.Load(sCmd.ExecuteReader())

        If dt1.Rows.Count = 0 Then Return rst

        Dim dr1 As DataRow = dt1.Rows(0)

        rst = TIMS.CINT1(dr1("orglevel"))

        Return rst
    End Function

    ''' <summary> 判斷登入計畫 開啟產投輸入視窗。(產投必輸入) </summary>
    ''' <param name="xTPlanID"></param>
    Sub Check_TPlanID28(ByVal xTPlanID As String)
        '**by Milor 20080502--User依據新的使用手冊要求把動態顯示改為直接顯示上年度字樣----start
        LabLastYear.Text = "上年度是否辦理本計劃"
        LabLastYear2.Text = "上年度核定人數執行率"
        '**by Milor 20080502----end

        '若為產業人才投資方案計劃則顯示轉用輸入表格
        '**by Milor 20080522--所有的計畫都要顯示計畫主持人，不限制只有產學訓----start
        Dim flag_TrTPlanID2854_Show As Boolean = False
        Dim flag_TrTPlanID28_Show As Boolean = False
        'If xTPlanID="" Then xTPlanID=sm.UserInfo.TPlanID 'xTPlanID 永遠有值 不管接收或是sm
        If xTPlanID <> "" Then
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(xTPlanID) > -1 Then flag_TrTPlanID2854_Show = True '產投/充飛才顯示
            If TIMS.Cst_TPlanID28.IndexOf(xTPlanID) > -1 Then flag_TrTPlanID28_Show = True '產投才顯示
        End If

        'TPlanID28A.Visible=TPlanID28.Visible
        LastYear.Value = (CInt(sm.UserInfo.Years) - 1).ToString
        TPlanID28B.Visible = flag_TrTPlanID2854_Show
        TPlanID28C.Visible = flag_TrTPlanID2854_Show
        TPlanID28D.Visible = flag_TrTPlanID2854_Show
        TrTPlanID28F1.Visible = flag_TrTPlanID2854_Show
        TrTPlanID28F2.Visible = flag_TrTPlanID2854_Show

        '2018 add:產投&充飛-機構屬性設定資料區顯示控制
        '共用-不顯示 (新增／修改)顯示
        TrTPlanID28OrgType1.Visible = False '其它動作不顯示
        Select Case Rq_ProcessType
            Case cst_Insert, cst_Update
                '新增-依登入時以登入時所選計畫做為判定依據(產投顯示)
                '修改-依該筆單位所屬計畫判定(產投顯示)
                TrTPlanID28OrgType1.Visible = flag_TrTPlanID28_Show
                'Case cst_Share 'TrTPlanID28OrgType1.Visible=(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1)
        End Select
    End Sub

    ''' <summary>
    ''' 產業人才投資計畫角色登入者則為唯讀
    ''' </summary>
    ''' <param name="OrgID"></param>
    Sub Set_GWOrgKind2(ByVal OrgID As String)
        Dim Kind_Flag As Boolean = True
        GWOrgKind.Text = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'GWOrgKind.Text=""
            '',case when ip.Years >='2015' then dbo.DECODE(b.OrgKind,'10','提升勞工自主學習計畫','產業人才投資計畫')
            ''  else dbo.DECODE(b.OrgKind,'10','勞工團體辦理勞工在職進修計畫','勞工在職進修計畫') end OrgPlanName
            'Dim OrgPlanName As String=TIMS.GET_OrgPlanName(OrgID)
            'If OrgPlanName <> "" Then GWOrgKind.Text="-" & OrgPlanName
            Dim OrgKind2 As String = TIMS.Get_OrgKind2(OrgID, TIMS.c_ORGID, objconn) 'G / W

            Dim parms_s1 As New Hashtable From {{"YEARS", $"{sm.UserInfo.Years}"}, {"OrgKind2", OrgKind2}}
            Dim sql_s1 As String = " SELECT DISTINCT OrgPlanName ,YEARS ,OrgKind2 FROM VIEW_RIDNAME WHERE YEARS=@YEARS AND OrgKind2=@OrgKind2"
            Dim dr As DataRow = DbAccess.GetOneRow(sql_s1, objconn, parms_s1)
            Dim OrgPlanName As String = ""
            If dr IsNot Nothing Then OrgPlanName = dr("OrgPlanName").ToString
            If OrgPlanName <> "" Then GWOrgKind.Text = "-" & OrgPlanName
            If OrgKind2 = "G" Then Kind_Flag = False '產業人才投資計畫角色登入者則為唯讀
            '**by Milor 20080522--計畫主持人不限定計畫都要顯示----start
            'TPlanID28.Visible=Kind_Flag  
            'TPlanID28A.Visible=TPlanID28.Visible
            TPlanID28B.Visible = Kind_Flag
            TPlanID28C.Visible = Kind_Flag
            TPlanID28D.Visible = Kind_Flag
            '**by Milor 20080522----end        
        End If
    End Sub

    ''' <summary>
    ''' 自動執行 計算 上年度核定人數執行率  
    ''' </summary>
    Sub AutoRunFindExeRate()
        Dim intLastYearExeRate As Integer = 0
        If TBID.Text <> "" AndAlso LastYear.Value <> "" Then
            txtLastYearExeRate.Text = FindExeRate(TBID.Text, LastYear.Value)
            If txtLastYearExeRate.Text = "" Then
                txtLastYearExeRate.Text = "0"
            Else
                intLastYearExeRate = Val(txtLastYearExeRate.Text)
                If intLastYearExeRate = 0 Then txtLastYearExeRate.Text = "0"
            End If
            If txtLastYearExeRate.Text <> "0" Then Common.SetListItem(LastYearExeRate, "1") '是
        End If
    End Sub

    Sub Save_OrgPlanInfodr2(ByRef drOrgPlan As DataRow, ByVal iRSID As Integer)
        'RSID=DbAccess.GetId(objTrans, "AUTH_RELSHIP_RSID_SEQ")
        If iRSID > -1 Then drOrgPlan("RSID") = iRSID

        drOrgPlan("OrgPName") = TB_OrgPName.Text
        drOrgPlan("ActNo") = TB_ActNo.Text

        city_code.Value = TIMS.ClearSQM(city_code.Value)
        city_code.Value = If(TIMS.IsZipCode(city_code.Value, objconn), city_code.Value, "")
        ZipCODEB3.Value = TIMS.ClearSQM(ZipCODEB3.Value)
        hidZipCODE6W.Value = TIMS.GetZIPCODE6W(city_code.Value, ZipCODEB3.Value)
        TBaddress.Text = TIMS.ClearSQM(TBaddress.Text)

        TBtel.Text = TIMS.ClearSQM(TBtel.Text)
        If (TBtel.Text <> "") Then TBtel.Text = Replace(TBtel.Text.ToString, ";", "；")
        TBm_name.Text = TIMS.ClearSQM(TBm_name.Text)
        TBm_Phone.Text = TIMS.ClearSQM(TBm_Phone.Text)

        drOrgPlan("ZipCode") = If(city_code.Value <> "", Val(city_code.Value), Convert.DBNull)
        drOrgPlan("ZipCODE6W") = If(hidZipCODE6W.Value <> "", hidZipCODE6W.Value, Convert.DBNull)
        drOrgPlan("Address") = TBaddress.Text

        drOrgPlan("Phone") = TBtel.Text
        drOrgPlan("MasterName") = TBm_name.Text
        drOrgPlan("MasterPhone") = If(TBm_Phone.Text <> "", TBm_Phone.Text, Convert.DBNull)
        drOrgPlan("ContactName") = TBContactName.Text
        drOrgPlan("ContactEmail") = TBmail.Text
        drOrgPlan("ContactCellPhone") = TBcontact_cellphone.Text
        drOrgPlan("TrainCap") = TB_TrainCap.Text
        drOrgPlan("FireControlState") = TB_FireControlState.Text

        drOrgPlan("ProTrainKind") = If(TB_ProTrainKind.Text = "", Convert.DBNull, TB_ProTrainKind.Text) '專長訓練職類
        drOrgPlan("ComSumm") = If(ComSumm.Text = "", Convert.DBNull, ComSumm.Text) '機構簡介
        drOrgPlan("PlanMaster") = If(PlanMaster.Text = "", Convert.DBNull, PlanMaster.Text)
        drOrgPlan("PlanMasterPhone") = If(PlanMasterPhone.Text = "", Convert.DBNull, PlanMasterPhone.Text)
        drOrgPlan("ContactFax") = If(ContactFax.Text = "", Convert.DBNull, ContactFax.Text)
        drOrgPlan("staffName") = If(staffName.Text = "", Convert.DBNull, staffName.Text)
        drOrgPlan("staffPhone") = If(staffPhone.Text = "", Convert.DBNull, staffPhone.Text)
        drOrgPlan("staffEmail") = If(staffEmail.Text = "", Convert.DBNull, staffEmail.Text)
        drOrgPlan("MemberNum") = If(MemberNum.Text = "", Convert.DBNull, MemberNum.Text)
        drOrgPlan("ActMemberNum") = If(ActMemberNum.Text = "", Convert.DBNull, ActMemberNum.Text)
        drOrgPlan("ActStaffNum") = If(ActStaffNum.Text = "", Convert.DBNull, ActStaffNum.Text)
        drOrgPlan("Accessible") = If(Accessible.SelectedValue = "NG", Convert.DBNull, Accessible.SelectedValue)
        drOrgPlan("Textbook") = If(Textbook.SelectedValue = "NG", Convert.DBNull, Textbook.SelectedValue)
        drOrgPlan("TeachAids") = If(TeachAids.SelectedValue = "NG", Convert.DBNull, TeachAids.SelectedValue)
        drOrgPlan("HumanHelp") = If(HumanHelp.SelectedValue = "NG", Convert.DBNull, HumanHelp.SelectedValue)
        drOrgPlan("ModifyAcct") = sm.UserInfo.UserID
        drOrgPlan("ModifyDate") = Now()
    End Sub

    ''' <summary>儲存(動作)</summary>
    ''' <param name="save_flag_ok">儲存狀況: true:正常/false:異常</param>
    Sub Insert_Auth_Relship(ByRef save_flag_ok As Boolean)
        Call AutoRunFindExeRate()

        'Dim DoubleComIDNO_flag As Boolean=False
        'Dim Errmsg As String=""
        'Dim chkTable As DataTable=Nothing
        'Dim chkdr As DataRow=Nothing
        'Dim Checkda As SqlDataAdapter=Nothing
        'Dim sqlAdapter, sqlAdapter2, sqlAdapter3 As SqlDataAdapter 'sqlAdapter--orginfo,sqlAdapter2--Auth_Relship,sqlAdapter3--orgplaninfo
        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim sqlAdapter2 As SqlDataAdapter = Nothing
        Dim sqlAdapter3 As SqlDataAdapter = Nothing
        Dim sqlAdapter4 As SqlDataAdapter = Nothing
        Dim sqlAdapter5 As SqlDataAdapter = Nothing

        Dim strSqlchk As String = ""
        Dim sqlTable As New DataTable 'Org_OrgInfo
        Dim sqldr As DataRow = Nothing
        'Dim sqlstrUpdate As String
        Dim iRSID As Integer = 0
        Dim iOrgID As Integer = 0
        Dim sql As String = ""
        Dim RelShipID As String = "" 'RelShip
        Dim s_RID As String = ""
        If Session("_Search") IsNot Nothing Then Session("_Search") = Session("_Search")

        Re_orgid = TIMS.ClearSQM(Request("orgid"))
        If Rq_ProcessType = cst_Update Then
            If Re_orgid = "" OrElse Not TIMS.IsNumberStr(Re_orgid) OrElse Val(Re_orgid) <= 0 Then
                'DbAccess.RollbackTrans(oTrans)
                'Call TIMS.CloseDbConn(oConn)
                Dim strScript1 As String = String.Concat("<script language=""javascript"">", "alert('機構設定異常，請重新操作!');", "</script>")
                Page.RegisterStartupScript("", strScript1)
                Exit Sub
            End If
            iOrgID = Val(Re_orgid)
        End If


        Dim oConn As SqlConnection = DbAccess.GetConnection()
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
        Try

            Select Case Rq_ProcessType
                Case cst_Update
                    'Re_orgid=TIMS.ClearSQM(Request("orgid"))
                    sql = String.Concat(" SELECT * FROM ORG_ORGINFO WHERE ORGID=", iOrgID)
                    sqlTable = DbAccess.GetDataTable(sql, sqlAdapter, oTrans)
                    sqldr = sqlTable.Rows(0)
                Case cst_Insert, cst_InsertChk
                    sql = " SELECT * FROM ORG_ORGINFO WHERE 1<>1 "
                    sqlTable = DbAccess.GetDataTable(sql, sqlAdapter, oTrans)
                    sqldr = sqlTable.NewRow
                    sqlTable.Rows.Add(sqldr)
                    'ORG_ORGINFO_ORGID_SEQ
                    iOrgID = DbAccess.GetNewId(oTrans, "ORG_ORGINFO_ORGID_SEQ,ORG_ORGINFO,ORGID")
                    sqldr("ORGID") = iOrgID
            End Select

            'If DoubleComIDNO_flag Then
            '    Me.txtLastYearExeRate.Text="0"
            '    Errmsg += vbCrLf & "請重新計算" & LabLastYear2.Text
            '    Common.MessageBox(Me, Errmsg)
            '    Exit Sub
            'End If

            Dim drDetail As DataRow = Nothing
            Dim drOrgPlan As DataRow = Nothing
            Dim dtRelShip As DataTable = Nothing
            Dim dtOrgPlan As DataTable = Nothing
            'Dim dtPlan_Planinfo As DataTable=Nothing
            'Dim dtAuth_AccRWPlan As DataTable=Nothing
            Dim sqlstr_Update As String = ""
            Dim update_OrgPlan As String = ""

            Select Case Rq_ProcessType
                Case cst_Insert, cst_InsertChk
                    '採新增功能 Auth_Relship 
                    sqlstr_Update = " SELECT * FROM AUTH_RELSHIP WHERE 1<>1 "
                    dtRelShip = DbAccess.GetDataTable(sqlstr_Update, sqlAdapter2, oTrans)
                    drDetail = dtRelShip.NewRow

                    update_OrgPlan = " SELECT * FROM ORG_ORGPLANINFO WHERE 1<>1 "
                    dtOrgPlan = DbAccess.GetDataTable(update_OrgPlan, sqlAdapter3, oTrans)
                    drOrgPlan = dtOrgPlan.NewRow

                Case cst_Update
                    '修改
                    'Dim sqlPlan_Planinfo As String=""
                    'Dim sqlAuth_AccRWPlan As String=""
                    'If sm.UserInfo.TPlanID="17" Then
                    '    sqlPlan_Planinfo=" SELECT * FROM PLAN_PLANINFO WHERE planid='" & Re_planid & "' AND RID='" & Re_rid & "' "
                    '    dtPlan_Planinfo=DbAccess.GetDataTable(sqlPlan_Planinfo, sqlAdapter4, oTrans)
                    '    sqlAuth_AccRWPlan=" SELECT * FROM AUTH_ACCRWPLAN WHERE planid='" & Re_planid & "' AND RID='" & Re_rid & "' "
                    '    dtAuth_AccRWPlan=DbAccess.GetDataTable(sqlAuth_AccRWPlan, sqlAdapter5, oTrans)
                    'End If
                    'Auth_Relship 
                    If Convert.ToString(ViewState(vs_RSID)) <> "" Then
                        sqlstr_Update = " SELECT * FROM AUTH_RELSHIP WHERE RSID='" & ViewState(vs_RSID) & "' AND orgid=" & Re_orgid
                    Else
                        sqlstr_Update = " SELECT * FROM AUTH_RELSHIP WHERE orgid=" & Re_orgid & " AND RID='" & Re_rid & "'"
                    End If
                    dtRelShip = DbAccess.GetDataTable(sqlstr_Update, sqlAdapter2, oTrans)
                    drDetail = dtRelShip.Rows(0)
                    '判斷RSID為何
                    Dim check_RSID As String = " SELECT RSID FROM AUTH_RELSHIP WHERE orgid=" & Re_orgid & " AND RID='" & Re_rid & "'"
                    Dim RSID_str As String = Convert.ToString(DbAccess.ExecuteScalar(check_RSID, oTrans))
                    update_OrgPlan = " SELECT * FROM Org_OrgPlanInfo WHERE RSID=" & RSID_str
                    dtOrgPlan = DbAccess.GetDataTable(update_OrgPlan, sqlAdapter3, oTrans)
                    drOrgPlan = dtOrgPlan.Rows(0)

            End Select
            s_RID = Me.RIDValue.Value

            'SELECT ISCONUNIT,COUNT(1) CNT FROM  ORG_ORGINFO GROUP BY ISCONUNIT
            'LastYearExeRate SELECT LastYearExeRate,COUNT(1) CNT FROM  ORG_ORGINFO GROUP BY LastYearExeRate ORDER BY 1
            '以下為新增和修改同樣更正的欄位,Org_OrgInfo
            Dim v_OrgKindList As String = TIMS.GetListValue(OrgKindList)
            '是否為管控單位 1:是0:否
            Dim v_IsConUnit As String = TIMS.GetListValue(IsConUnit)
            Dim v_LastYearExeRate As String = TIMS.GetListValue(LastYearExeRate)
            txtLastYearExeRate.Text = TIMS.ClearSQM(txtLastYearExeRate.Text)
            '上年度是否辦理本計劃 -1:否 1:是
            Dim s_LastYearExeRate As String = If(v_LastYearExeRate = "-1", v_LastYearExeRate, txtLastYearExeRate.Text)
            Dim v_IsConTTQS As String = TIMS.GetListValue(IsConTTQS) '通過TTQS

            Select Case Rq_ProcessType
                Case cst_Update, cst_Insert, cst_InsertChk
                    city_code_org.Value = TIMS.ClearSQM(city_code_org.Value)
                    city_code_org.Value = If(TIMS.IsZipCode(city_code_org.Value, objconn), city_code_org.Value, "")
                    ZipCODEB3_Org.Value = TIMS.ClearSQM(ZipCODEB3_Org.Value)
                    hidZipCODE6W_Org.Value = TIMS.GetZIPCODE6W(city_code_org.Value, ZipCODEB3_Org.Value)
                    TBaddress_Org.Text = TIMS.ClearSQM(TBaddress_Org.Text)

                    '以下為新增和修改同樣更正的欄位,Org_OrgInfo
                    sqldr("ModifyAcct") = sm.UserInfo.UserID
                    sqldr("ModifyDate") = Now()
                    sqldr("OrgKind") = v_OrgKindList
                    sqldr("OrgKind2") = If(v_OrgKindList = "10", "W", "G")
                    sqldr("OrgName") = TBtitle.Text
                    sqldr("ComIDNO") = TBID.Text
                    sqldr("ComCIDNO") = TBseqno.Text
                    '是否為管控單位 1:是0:否
                    sqldr("IsConUnit") = If(v_IsConUnit <> "", v_IsConUnit, "0")
                    sqldr("LastYearExeRate") = s_LastYearExeRate
                    sqldr("IsConTTQS") = If(v_IsConTTQS <> "", v_IsConTTQS, Convert.DBNull) '通過TTQS
                    sqldr("BankName") = BankName.Text
                    sqldr("ExBankName") = ExBankName.Text
                    sqldr("AccNo") = AccNo.Text
                    sqldr("AccName") = AccName.Text

                    Dim orgtype1Dr As DataRow = GetOrgType1ByTypeID(dl_typeid1.SelectedValue, dl_typeid2.SelectedValue)
                    hidZipCODE6W_Org.Value = TIMS.GetZIPCODE6W(city_code_org.Value, ZipCODEB3_Org.Value)
                    '2018 add:merge TC_01_017 寫入訓練機構屬性設定(只提供產投 & 充飛計畫在 新增/修改 模式時填寫)
                    Dim sTPlanID As String = TIMS.GetTPlanID(PlanIDValue.Value, objconn)
                    If sTPlanID <> "" AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                        If Rq_ProcessType = cst_Update Then
                            If sm.UserInfo.LID = 0 Then
                                sqldr("orgkind1") = If(orgtype1Dr IsNot Nothing, orgtype1Dr("ORGTYPEID1"), Convert.DBNull)
                                sqldr("orgzipcode") = If(city_code_org.Value = "", Convert.DBNull, Val(city_code_org.Value))
                                sqldr("orgzipCODE6W") = If(hidZipCODE6W_Org.Value <> "", hidZipCODE6W_Org.Value, Convert.DBNull)
                                sqldr("orgaddress") = If(TBaddress_Org.Text = "", Convert.DBNull, TBaddress_Org.Text)
                            End If
                        ElseIf Rq_ProcessType = cst_Insert Then
                            sqldr("orgkind1") = If(orgtype1Dr IsNot Nothing, orgtype1Dr("ORGTYPEID1"), Convert.DBNull)
                            sqldr("orgzipcode") = If(city_code_org.Value = "", Convert.DBNull, Val(city_code_org.Value))
                            sqldr("orgzipCODE6W") = If(hidZipCODE6W_Org.Value <> "", hidZipCODE6W_Org.Value, Convert.DBNull)
                            sqldr("orgaddress") = If(TBaddress_Org.Text = "", Convert.DBNull, TBaddress_Org.Text)
                        End If
                    End If

                    'If Rq_ProcessType=cst_Insert OrElse Rq_ProcessType=cst_Update Then
                    '    If sTPlanID <> "" AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                    '        Dim orgtype1Dr As DataRow=getOrgType1ByTypeID(dl_typeid1.SelectedValue, dl_typeid2.SelectedValue)
                    '        hidZipCODE6W_Org.Value=TIMS.GetZIPCODE6W(city_code_org.Value, ZipCODEB3_Org.Value)
                    '        sqldr("orgkind1")=If(orgtype1Dr IsNot Nothing, orgtype1Dr("orgtypeid1"), Convert.DBNull)
                    '        sqldr("orgzipcode")=If(city_code_org.Value="", Convert.DBNull, Val(city_code_org.Value))
                    '        sqldr("orgzipCODE6W")=If(hidZipCODE6W_Org.Value <> "", hidZipCODE6W_Org.Value, Convert.DBNull)
                    '        sqldr("orgaddress")=If(TBaddress_Org.Text="", Convert.DBNull, TBaddress_Org.Text)
                    '    End If
                    'End If
                    ' 2018 add 記錄交易log 
                    Call SaveOrgOrgInfoLog(oTrans, sqldr, Rq_ProcessType)
                    DbAccess.UpdateDataTable(sqlTable, sqlAdapter, oTrans)
                    'If ProcessType=cst_Insert Or ProcessType=cst_InsertChk Then OrgID=DbAccess.GetId(objTrans, "ORG_ORGINFO_ORGID_SEQ")

                    'OrgPlanInfo (Org_OrgPlanInfo )
                    Save_OrgPlanInfodr2(drOrgPlan, -1)
            End Select

            Select Case Rq_ProcessType
                Case cst_Insert, cst_InsertChk '(Case cst_Insert, cst_InsertChk)()
                    '新增 Auth_Relship
                    'Dim strUnkey As String
                    'strUnkey=sqldr("orgid")
                    'Auth_Relship
                    sqldr = drDetail
                    sqldr("PlanID") = PlanIDValue.Value '若沒選擇隸屬機構
                    sqldr("OrgID") = iOrgID '新增的key值,從Org_OrgInfo抓取
                    If level_list.SelectedValue = 1 Then
                        '取目前選取計畫的relship
                        Dim sqlstr_A As String = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID='" & s_RID & "'"
                        RelShipID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, oTrans))
                        '選取新增項目的最大項
                        Dim sqlstr_next As String = " SELECT MAX(rid) maxrid FROM AUTH_RELSHIP WHERE orglevel=1 AND relship LIKE '" & RelShipID & "%'"
                        maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, oTrans))
                        If maxrid = "" Then
                            sqldr("RID") = Chr(Asc(s_RID) + 1)
                            sqldr("Relship") = String.Concat(RelShipID, Chr(Asc(s_RID) + 1), "/")
                            sqldr("OrgLevel") = 1
                        Else
                            sqldr("RID") = Chr(Asc(maxrid) + 1)
                            sqldr("Relship") = String.Concat(RelShipID, Chr(Asc(maxrid) + 1), "/")
                            sqldr("OrgLevel") = 1
                        End If
                    ElseIf level_list.SelectedValue = 2 Then
                        '取目前選取計畫的relship
                        Dim sqlstr_A As String = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID='" & s_RID & "'"
                        RelShipID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, oTrans))
                        '取目前選取計畫的orglevel
                        Dim sqlstr_orglevel As String = " SELECT a.ORGLEVEL FROM AUTH_RELSHIP a WHERE RID='" & Me.RIDValue.Value & "'"
                        planid2 = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_orglevel, oTrans))
                        Dim sqlstr_next As String = ""
                        If Left(s_RID, 1) = "A" Then
                            Select Case planid2
                                Case "0"
                                    sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,Len(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & s_RID & "%' AND orglevel='2' "
                                Case Else
                                    sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,Len(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & s_RID & "%' AND orglevel='3' "
                            End Select
                        Else
                            Select Case planid2
                                Case "1", "2"
                                    '選取新增項目的最大項
                                    '新增ex@id=c01底下的機構
                                    If s_RID.Length > 2 Then
                                        sqlstr_next = ""
                                        sqlstr_next &= " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP "
                                        sqlstr_next &= " WHERE relship LIKE '" & RelShipID & "%' AND orglevel=" & planid2 & "+1 "
                                        sqlstr_next &= " AND LEN(rid)=LEN('" & s_RID & "')+3 "
                                    Else '新增ex@id=c底下的機構
                                        sqlstr_next = ""
                                        sqlstr_next &= " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP "
                                        sqlstr_next &= " WHERE relship LIKE '" & RelShipID & "%' AND orglevel=" & planid2 & "+1 "
                                    End If
                                Case "3"
                                    Dim new_plan_list As String = String.Concat(RelShipID, s_RID)
                                    'sqlstr_next=" SELECT convert(varchar, max(convert(int,substring(rid,2,len(rid)-1)))) as newmix  from  Auth_Relship where   relship like '" & new_plan_list & "%' and orglevel=" & planid2 & "" '★
                                    sqlstr_next = ""
                                    sqlstr_next &= " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP "
                                    sqlstr_next &= " WHERE relship LIKE '" & new_plan_list & "%' AND orglevel=" & planid2
                            End Select
                        End If

                        maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, oTrans))

                        '欲新增的計畫是否存在
                        If maxrid = "" OrElse maxrid = "0" Then
                            If Left(s_RID, 1) = "A" Then
                                If planid2 = "0" Then
                                    sqldr("RID") = String.Concat(s_RID, "01")
                                    sqldr("Relship") = String.Concat(RelShipID, s_RID, "01", "/")
                                    sqldr("OrgLevel") = 2
                                Else
                                    sqldr("RID") = String.Concat(s_RID, "001")
                                    sqldr("Relship") = String.Concat(RelShipID, s_RID, "001", "/")
                                    sqldr("OrgLevel") = If(planid2 <> "3", CInt(planid2) + 1, 3)
                                End If
                            Else
                                If planid2 = "1" Then '沒有子單位,署(局)
                                    sqldr("RID") = String.Concat(s_RID, "01")
                                    sqldr("Relship") = String.Concat(RelShipID, s_RID, "01", "/")
                                    sqldr("OrgLevel") = CInt(planid2) + 1
                                ElseIf planid2 > "1" Then '沒有子單位,orglevel=2,分署(中心),orglevel=3,委訓(2005/3/22)
                                    sqldr("RID") = String.Concat(s_RID, "001")
                                    sqldr("Relship") = String.Concat(RelShipID, s_RID, "001", "/")
                                    sqldr("OrgLevel") = If(planid2 <> "3", CInt(planid2) + 1, 3)
                                End If
                            End If
                        Else
                            '有子單位,加1
                            If Left(s_RID, 1) = "A" Then
                                If planid2 = "0" Then
                                    sqldr("RID") = s_RID.Substring(0, 1) & (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 2))
                                    sqldr("Relship") = RelShipID + s_RID.Substring(0, 1) & (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 2)) + "/"
                                    sqldr("OrgLevel") = 2
                                ElseIf planid2 > "0" Then
                                    sqldr("RID") = s_RID & (CInt(Right(maxrid, 3)) + 1).ToString("000")
                                    sqldr("Relship") = RelShipID + s_RID & (CInt(Right(maxrid, 3)) + 1).ToString("000") + "/"
                                    sqldr("OrgLevel") = If(planid2 <> "3", CInt(planid2) + 1, 3)
                                End If
                            Else
                                If planid2 = "1" Then '署(局)
                                    '有子單位,加1
                                    sqldr("RID") = s_RID.Substring(0, 1) & (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 1))
                                    sqldr("Relship") = RelShipID + s_RID.Substring(0, 1) & (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 1)) + "/"
                                    sqldr("OrgLevel") = CInt(planid2) + 1
                                ElseIf planid2 > "1" Then '分署(中心),委訓(2005/3/22)
                                    If s_RID.Length > 2 Then '新增ex@id=c01底下的機構
                                        sqldr("RID") = s_RID.Substring(0, 1) & maxrid.Substring(0, maxrid.Length - 3) & (CInt(Right(maxrid, 3)) + 1).ToString("000")
                                        sqldr("Relship") = RelShipID + s_RID.Substring(0, 1) & maxrid.Substring(0, maxrid.Length - 3) & (CInt(Right(maxrid, 3)) + 1).ToString("000") + "/"
                                    Else
                                        sqldr("RID") = s_RID & (CInt(Right(maxrid, 3)) + 1).ToString("000")
                                        sqldr("Relship") = RelShipID + s_RID & (CInt(Right(maxrid, 3)) + 1).ToString("000") + "/"
                                    End If
                                    sqldr("OrgLevel") = If(planid2 <> "3", CInt(planid2) + 1, 3)
                                End If
                            End If
                        End If
                    End If
                    Dim v_DistrictList As String = TIMS.GetListValue(DistrictList)
                    'AUTH_RELSHIP-新增
                    sqldr("DistID") = v_DistrictList 'DistrictList.SelectedValue
                    sqldr("ModifyAcct") = sm.UserInfo.UserID
                    sqldr("ModifyDate") = Now()
                    dtRelShip.Rows.Add(sqldr)

                    iRSID = DbAccess.GetNewId(oTrans, "AUTH_RELSHIP_RSID_SEQ,AUTH_RELSHIP,RSID")
                    ViewState(vs_RSID) = Convert.ToString(iRSID)
                    sqldr("RSID") = iRSID
                    'Auth_Relship 檢查RID
                    Dim chkRID As String = Convert.ToString(sqldr("RID"))
                    '=======================
                    ' 2018 add 記錄交易log (auth_relship ==> sys_trans_log)
                    Call SaveAuthRelshipLog(oTrans, sqldr, Rq_ProcessType)
                    '=======================
                    DbAccess.UpdateDataTable(dtRelShip, sqlAdapter2, oTrans)
                    If ChkRID_ERROR(chkRID, 1, oConn, oTrans) Then
                        DbAccess.RollbackTrans(oTrans)
                        Call TIMS.CloseDbConn(oConn)
                        Dim strScript1 As String = String.Concat("<script language=""javascript"">", "alert('機構設定異常，請重新操作!!');", "</script>")
                        Page.RegisterStartupScript("", strScript1)
                        Exit Sub
                    End If

                    'Org_OrgPlaninfo  新增
                    'RSID=DbAccess.GetId(objTrans, "AUTH_RELSHIP_RSID_SEQ")
                    drOrgPlan("RSID") = iRSID
                    dtOrgPlan.Rows.Add(drOrgPlan)
                    '=======================
                    ' 2018 add 記錄交易log (org_orgplaninfo ==> sys_trans_log)
                    Call SaveOrgOrgPlanInfoLog(oTrans, drOrgPlan, Rq_ProcessType)
                    '=======================
                    DbAccess.UpdateDataTable(dtOrgPlan, sqlAdapter3, oTrans)

                    If Rq_ProcessType = cst_Insert Then
                        Dim strScript As String = String.Concat("<script language=""javascript"">", "alert('新增成功!!');", "location.href='TC_01_002.aspx?ID='+document.getElementById('Re_ID').value;", "</script>")
                        Page.RegisterStartupScript("", strScript)
                    End If

                    If Rq_ProcessType = cst_InsertChk Then
                        Dim chkTable As DataTable = Nothing
                        Dim chkdr As DataRow = Nothing
                        Dim Checkda As SqlDataAdapter = Nothing
                        Dim chksql As String = ""
                        Dim Account As String = ""
                        TBID.Text = TIMS.ClearSQM(TBID.Text)
                        chksql = " SELECT * FROM AUTH_APPLY WHERE OrgID=-1 AND ComIDNO='" & TBID.Text & "'"
                        chkTable = DbAccess.GetDataTable(chksql, Checkda, oTrans)
                        If chkTable.Rows.Count > 0 Then
                            chkdr = chkTable.Rows(0)
                            chkdr("OrgID") = iOrgID
                            Account = chkdr("Account")
                            '=======================
                            ' 2018 add 記錄交易log (auth_apply ==> sys_trans_log)
                            Call SaveAuthApplyLog(oTrans, chkdr)
                            '=======================
                            DbAccess.UpdateDataTable(chkTable, Checkda, oTrans)
                        End If
                        'Org_Apply
                        chksql = ""
                        chksql &= " SELECT * FROM ORG_APPLY WHERE Result IS NULL AND ComIDNO='" & TBID.Text & "'"
                        chkTable = DbAccess.GetDataTable(chksql, Checkda, oTrans)
                        If chkTable.Rows.Count > 0 Then
                            chkdr = chkTable.Rows(0)
                            chkdr("Result") = "Y"
                            '=======================
                            ' 2018 add 記錄交易log (org_apply ==> sys_trans_log)
                            Call SaveOrgApplyLog(oTrans, chkdr)
                            '=======================
                            DbAccess.UpdateDataTable(chkTable, Checkda, oTrans)
                        End If
                        Dim strScript As String = String.Concat("<script language=""javascript"">", "alert('機構審核成功!!請接著審核登入帳號[", Account, "]');", "location.href='TC_01_002.aspx?ID='+document.getElementById('Re_ID').value;", "</script>")
                        Page.RegisterStartupScript("", strScript)
                    End If

                Case cst_Update '(Case cst_Update) 修改
                    sqldr = drDetail

                    '補助地方政府訓練  更改隸屬機構
                    'Dim flag_TPlanID17 As Boolean=False
                    'If sm.UserInfo.TPlanID="17" Then
                    '    If Not Me.ViewState(cst_vs_sqlstr_list) Is Nothing Then
                    '        Dim list As DataRow=Nothing
                    '        list=DbAccess.GetOneRow(Me.ViewState(cst_vs_sqlstr_list), oTrans)
                    '        Dim myarray As Array
                    '        myarray=list("Relship").Split("/")
                    '        Dim range As Integer=myarray.Length - 3
                    '        Parent_list=myarray(range)
                    '    End If
                    '    'Parent_list 原RID之父層( 'Re_rid 原RID )
                    '    'RIDValue.Value 原RID之父層(或選擇後之RID父層)
                    '    If Parent_list <> RIDValue.Value And Re_rid <> RIDValue.Value Then  '有更改隸屬機構
                    '        flag_TPlanID17=True '確認要重設設定 RID
                    '    End If
                    'End If
                    'OrElse flag_TPlanID17 
                    If sqldr("PlanID", DataRowVersion.Current) <> sqldr("PlanID", DataRowVersion.Original) Then
                        Dim v_DistrictList As String = TIMS.GetListValue(DistrictList)
                        If Parent_list <> RIDValue.Value Then '計畫階層改變,更新關係檔
                            sqldr("PlanID") = If(PlanIDValue.Value <> "", PlanIDValue.Value, Convert.DBNull)
                            sqldr("DistID") = v_DistrictList 'DistrictList.SelectedValue
                            sqldr("ModifyAcct") = sm.UserInfo.UserID
                            sqldr("ModifyDate") = Now()
                            If level_list.SelectedValue = 2 Then
                                '取目前選取計畫的relshup
                                Dim sqlstr_A As String = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID='" & s_RID & "'"
                                RelShipID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, oTrans))
                                '取目前選取計畫的orglevel
                                Dim sqlstr_orglevel As String = " SELECT a.orglevel FROM AUTH_RELSHIP a WHERE RID='" & Me.RIDValue.Value & "'"
                                planid2 = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_orglevel, oTrans))
                                Dim sqlstr_next As String = ""
                                If Left(s_RID, 1) = "A" Then
                                    If planid2 = "0" Then
                                        sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & s_RID & "%' AND orglevel='2' "
                                    Else
                                        sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & s_RID & "%' AND orglevel='3' "
                                    End If
                                Else
                                    If planid2 = "1" OrElse planid2 = "2" Then
                                        '選取新增項目的最大項
                                        '新增ex@id=c01底下的機構
                                        If s_RID.Length > 2 Then
                                            sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,Len(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & "%' AND orglevel=" & planid2 & "+1 AND LEN(rid)=LEN('" & s_RID & "')+3 "
                                        Else '新增ex@id=c底下的機構
                                            sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,Len(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & "%' AND orglevel=" & planid2 & "+1 "
                                        End If
                                    ElseIf planid2 = "3" Then
                                        Dim new_plan_list As String = String.Concat(RelShipID, s_RID)
                                        sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,Len(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & new_plan_list & "%' AND orglevel=" & planid2 & " "
                                    End If
                                End If

                                maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, oTrans))

                                '欲新增的計畫是否存在
                                If maxrid = "" OrElse maxrid = "0" Then
                                    If Left(s_RID, 1) = "A" Then
                                        If planid2 = "0" Then '沒有子單位
                                            sqldr("RID") = s_RID & "01"
                                            sqldr("Relship") = String.Format("{0}{1}{2}/", RelShipID, s_RID, "01")
                                            sqldr("OrgLevel") = 2
                                        ElseIf planid2 > "0" Then '沒有子單位,orglevel=2,分署(中心),orglevel=3,委訓(2005/3/22)
                                            sqldr("RID") = s_RID & "001"
                                            sqldr("Relship") = String.Format("{0}{1}{2}/", RelShipID, s_RID, "001")
                                            sqldr("OrgLevel") = If(planid2 <> "3", (CInt(planid2) + 1), 3)
                                        End If
                                    Else
                                        If planid2 = "1" Then '沒有子單位
                                            sqldr("RID") = s_RID & "01"
                                            sqldr("Relship") = String.Format("{0}{1}{2}/", RelShipID, s_RID, "01")
                                            sqldr("OrgLevel") = CInt(planid2) + 1
                                        ElseIf planid2 > "1" Then '沒有子單位,orglevel=2,分署(中心),orglevel=3,委訓(2005/3/22)
                                            sqldr("RID") = s_RID & "001"
                                            sqldr("Relship") = String.Format("{0}{1}{2}/", RelShipID, s_RID, "001")
                                            sqldr("OrgLevel") = If(planid2 <> "3", (CInt(planid2) + 1), 3)
                                        End If
                                    End If
                                Else
                                    If Left(s_RID, 1) = "A" Then
                                        If planid2 = "0" Then '署(局)
                                            '有子單位,加1
                                            sqldr("RID") = String.Concat(s_RID.Substring(0, 1), (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 2)))
                                            sqldr("Relship") = String.Concat(RelShipID, s_RID.Substring(0, 1), (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 2)), "/")
                                            sqldr("OrgLevel") = 2
                                        ElseIf planid2 > "0" Then '分署(中心),委訓(2005/3/22)
                                            '有子單位,加1
                                            sqldr("RID") = String.Concat(s_RID, (CInt(Right(maxrid, 3)) + 1).ToString("000"))
                                            sqldr("Relship") = String.Concat(RelShipID, s_RID, (CInt(Right(maxrid, 3)) + 1).ToString("000"), "/")
                                            sqldr("OrgLevel") = If(planid2 <> "3", (CInt(planid2) + 1), 3)
                                        End If
                                    Else
                                        If planid2 = "1" Then '署(局)
                                            '有子單位,加1
                                            sqldr("RID") = String.Concat(s_RID.Substring(0, 1), (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 1)))
                                            sqldr("Relship") = String.Concat(RelShipID, s_RID.Substring(0, 1), (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 1)), "/")
                                            sqldr("OrgLevel") = CInt(planid2) + 1
                                        ElseIf planid2 > "1" Then '分署(中心),委訓(2005/3/22)
                                            '有子單位,加1
                                            If s_RID.Length > 2 Then '新增ex@id=c01底下的機構
                                                sqldr("RID") = String.Concat(s_RID.Substring(0, 1), maxrid.Substring(0, maxrid.Length - 3) & (CInt(Right(maxrid, 3)) + 1).ToString("000"))
                                                sqldr("Relship") = String.Concat(RelShipID, s_RID.Substring(0, 1), maxrid.Substring(0, maxrid.Length - 3), (CInt(Right(maxrid, 3)) + 1).ToString("000"), "/")
                                            Else
                                                sqldr("RID") = String.Concat(s_RID, (CInt(Right(maxrid, 3)) + 1).ToString("000"))
                                                sqldr("Relship") = String.Concat(RelShipID, s_RID, (CInt(Right(maxrid, 3)) + 1).ToString("000"), "/")
                                            End If
                                            sqldr("OrgLevel") = If(planid2 <> "3", (CInt(planid2) + 1), 3)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                    Dim sqldr_RID As String = Convert.ToString(sqldr("RID"))
                    'If flag_TPlanID17 Then
                    '    If dtAuth_AccRWPlan IsNot Nothing AndAlso dtAuth_AccRWPlan.Rows.Count > 0 Then
                    '        For Each dr As DataRow In dtAuth_AccRWPlan.Rows
                    '            dr("RID")=sqldr_RID
                    '        Next
                    '    End If
                    '    If dtPlan_Planinfo IsNot Nothing AndAlso dtPlan_Planinfo.Rows.Count > 0 Then
                    '        For Each dr As DataRow In dtPlan_Planinfo.Rows
                    '            dr("RID")=sqldr_RID
                    '        Next
                    '    End If
                    '    If Me.ViewState(flag_TPlanID17_plan_planinfo) Then DbAccess.UpdateDataTable(dtPlan_Planinfo, sqlAdapter4, oTrans)
                    '    If Me.ViewState(flag_TPlanID17_Auth_AccRWPlan) Then DbAccess.UpdateDataTable(dtAuth_AccRWPlan, sqlAdapter5, oTrans)
                    'End If

                    '=======================
                    ' 2018 add 記錄交易log (auth_relship ==> sys_trans_log)
                    '=======================
                    Call SaveAuthRelshipLog(oTrans, sqldr, Rq_ProcessType)
                    '=======================
                    ' 2018 add 記錄交易log (org_orgplaninfo ==> sys_trans_log)                   
                    '=======================
                    Call SaveOrgOrgPlanInfoLog(oTrans, drOrgPlan, Rq_ProcessType)

                    DbAccess.UpdateDataTable(dtRelShip, sqlAdapter2, oTrans)
                    DbAccess.UpdateDataTable(dtOrgPlan, sqlAdapter3, oTrans)

                    Dim strScript As String = String.Concat("<script language=""javascript"">", "alert('修改成功!!');", "location.href='TC_01_002.aspx?ID='+document.getElementById('Re_ID').value;", "</script>")
                    Page.RegisterStartupScript("", strScript)

                Case cst_Share 'Case cst_Share (新增)
                    If RIDValue.Value = "" Then
                        DbAccess.RollbackTrans(oTrans)
                        Call TIMS.CloseDbConn(oConn)
                        Dim strScript1 As String = String.Concat("<script language=""javascript"">", "alert('請重新選擇隸屬機構!!!!');", "</script>")
                        Page.RegisterStartupScript("", strScript1)
                        Exit Sub
                    End If

                    Dim drShare As DataRow
                    Dim dtShare As DataTable 'Auth_Relship
                    Dim sqlstr_share As String
                    'Dim count_org, i As Integer
                    Dim v_DistrictList As String = TIMS.GetListValue(DistrictList)
                    PlanIDValue.Value = TIMS.ClearSQM(PlanIDValue.Value)
                    RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                    Dim sqlstr_relship As String = "SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID='" & RIDValue.Value & "'"
                    Dim drREL As DataRow = DbAccess.GetOneRow(sqlstr_relship, oTrans)
                    If drREL Is Nothing Then
                        DbAccess.RollbackTrans(oTrans)
                        Call TIMS.CloseDbConn(oConn)
                        Dim strScript1 As String = String.Concat("<script language=""javascript"">", "alert('共用機構設定參數有誤!');", "</script>")
                        Page.RegisterStartupScript("", strScript1)
                        Exit Sub
                    End If
                    Dim relship_str As String = Convert.ToString(drREL("RELSHIP"))
                    '2005/5/24修正共用是否重複之判斷(同轄區,同計畫,同orgid)
                    Dim check_list As String = "select b.Relship FROM ORG_ORGINFO a JOIN AUTH_RELSHIP b on a.orgid=b.orgid where b.DistID='" & v_DistrictList & "' and b.PlanID='" & PlanIDValue.Value & "' and b.orgid='" & Re_orgid & "' and b.relship  like '" & relship_str & "%'"
                    '檢查機構是否重複
                    '檢查---------------------------------Start
                    Dim da As SqlDataAdapter = Nothing
                    Dim dt As DataTable = DbAccess.GetDataTable(check_list, da, oTrans)
                    For Each dr As DataRow In dt.Rows
                        Dim OrgLevel As Array = Split(dr("Relship"), "/")
                        If OrgLevel.Length > 1 Then
                            Dim ParentsOrg As String
                            ParentsOrg = OrgLevel(OrgLevel.Length - 3)
                            If ParentsOrg = RIDValue.Value Then
                                DbAccess.RollbackTrans(oTrans)
                                Call TIMS.CloseDbConn(oConn)
                                Dim strScript1 As String = String.Concat("<script language=""javascript"">", "alert('共用機構設定重複!');", "</script>")
                                Page.RegisterStartupScript("", strScript1)
                                Exit Sub
                            End If
                        End If
                    Next
                    '檢查---------------------------------End

                    'count_org=Convert.ToInt16(DbAccess.GetCount(check_list, objTrans))
                    'Do While i <= count_org

                    '    i=i + 1
                    'Loop
                    'If DbAccess.GetCount(check_list, objTrans) > 0 Then
                    '    '判斷是否有新增重複的資料
                    '    Dim strScript1 As String
                    '    strScript1="<script language=""javascript"">" + vbCrLf
                    '    strScript1 += "alert('共用機構設定重複!!!!');" + vbCrLf
                    '    strScript1 += "</script>"
                    '    Page.RegisterStartupScript("", strScript1)
                    '    Exit Sub
                    'End If

                    'sqldr=drDetail
                    'sqlTable=dtRelship
                    'Auth_Relship
                    sqlstr_share = "SELECT * FROM AUTH_RELSHIP WHERE 1<>1"
                    dtShare = DbAccess.GetDataTable(sqlstr_share, sqlAdapter2, oTrans)
                    drShare = dtShare.NewRow
                    'Org_OrgPlanInfo
                    update_OrgPlan = "SELECT * FROM ORG_ORGPLANINFO WHERE 1<>1"
                    dtOrgPlan = DbAccess.GetDataTable(update_OrgPlan, sqlAdapter3, oTrans)
                    drOrgPlan = dtOrgPlan.NewRow
                    Re_orgid = TIMS.ClearSQM(Request("orgid")) '新增的key值,抓取前一頁回傳值
                    drShare("orgid") = Re_orgid

                    If level_list.SelectedValue = 1 Then
                        '取目前選取計畫的relship
                        Dim sqlstr_A As String = "SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID='" & s_RID & "'"
                        RelShipID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, oTrans))
                        '選取新增項目的最大項
                        Dim sqlstr_next As String = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(RID,2,LEN(RID)-1))) NEWMIX FROM AUTH_RELSHIP WHERE ORGLEVEL=1 AND RELSHIP LIKE '" & RelShipID & "%'"
                        maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, oTrans))
                        If maxrid = "" Then
                            drShare("RID") = Chr(Asc(s_RID) + 1)
                            drShare("Relship") = String.Concat(RelShipID, Chr(Asc(s_RID) + 1), "/")
                            drShare("OrgLevel") = 1
                        Else
                            drShare("RID") = Chr(Asc(maxrid) + 1)
                            drShare("Relship") = String.Concat(RelShipID, Chr(Asc(maxrid) + 1), "/")
                            drShare("OrgLevel") = 1
                        End If
                    ElseIf level_list.SelectedValue = 2 Then
                        '取目前選取計畫的relship
                        Dim sqlstr_A As String = "SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID='" & s_RID & "'"
                        RelShipID = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_A, oTrans))
                        '取目前選取計畫的orglevel
                        Dim sqlstr_orglevel As String = " SELECT a.ORGLEVEL FROM AUTH_RELSHIP a WHERE RID='" & RIDValue.Value & "'"
                        planid2 = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_orglevel, oTrans))
                        Dim sqlstr_next As String = ""
                        If Left(s_RID, 1) = "A" Then
                            Select Case planid2
                                Case "0"
                                    sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & s_RID & "%' AND orglevel=2"
                                Case Else
                                    sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & s_RID & "%' AND orglevel=3"
                            End Select
                        Else
                            Select Case planid2
                                Case "1", "2"
                                    '選取新增項目的最大項
                                    '新增ex@id=c01底下的機構
                                    If s_RID.Length > 2 Then
                                        sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & "%' AND orglevel=" & planid2 & "+1 AND LEN(rid)=LEN('" & s_RID & "')+3 "
                                    Else '新增ex@id=c底下的機構
                                        sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & RelShipID & "%' AND orglevel=" & planid2 & "+1 "
                                    End If
                                Case "3"
                                    Dim new_plan_list As String = String.Concat(RelShipID, s_RID)
                                    sqlstr_next = " SELECT MAX(CONVERT(NUMERIC, dbo.SUBSTR3(rid,2,LEN(rid)-1))) newmix FROM AUTH_RELSHIP WHERE relship LIKE '" & new_plan_list & "%' AND orglevel=" & planid2 & " "
                            End Select
                        End If

                        maxrid = Convert.ToString(DbAccess.ExecuteScalar(sqlstr_next, oTrans))

                        If maxrid = "" OrElse maxrid = "0" Then
                            If Left(s_RID, 1) = "A" Then
                                If planid2 = "0" Then '沒有子單位
                                    drShare("RID") = String.Concat(s_RID, "01")
                                    drShare("Relship") = String.Concat(RelShipID, s_RID, "01", "/")
                                    drShare("OrgLevel") = 2
                                ElseIf planid2 > "0" Then '沒有子單位,orglevel=2,分署(中心),orglevel=3,委訓(2005/3/22)
                                    drShare("RID") = String.Concat(s_RID, "001")
                                    drShare("Relship") = String.Concat(RelShipID, s_RID, "001", "/")
                                    drShare("OrgLevel") = If(planid2 <> "3", CInt(planid2) + 1, 3)
                                End If
                            Else
                                If planid2 = "1" Then '沒有子單位
                                    drShare("RID") = String.Concat(s_RID, "01")
                                    drShare("Relship") = String.Concat(RelShipID, s_RID, "01", "/")
                                    drShare("OrgLevel") = CInt(planid2) + 1
                                ElseIf planid2 > "1" Then '沒有子單位,orglevel=2,分署(中心),orglevel=3,委訓(2005/3/22)
                                    drShare("RID") = String.Concat(s_RID, "001")
                                    drShare("Relship") = String.Concat(RelShipID, s_RID, "001", "/")
                                    drShare("OrgLevel") = If(planid2 <> "3", CInt(planid2) + 1, 3)
                                End If
                            End If
                        Else
                            If Left(s_RID, 1) = "A" Then
                                If planid2 = "0" Then
                                    drShare("RID") = String.Concat(s_RID.Substring(0, 1), (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 2)))
                                    drShare("Relship") = String.Concat(RelShipID, s_RID.Substring(0, 1), (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 2)), "/")
                                    drShare("OrgLevel") = 2
                                ElseIf planid2 > "0" Then
                                    drShare("RID") = String.Concat(s_RID, (CInt(Right(maxrid, 3)) + 1).ToString("000"))
                                    drShare("Relship") = String.Concat(RelShipID, s_RID, (CInt(Right(maxrid, 3)) + 1).ToString("000"), "/")
                                    drShare("OrgLevel") = If(planid2 <> "3", CInt(planid2) + 1, 3)
                                End If
                            Else
                                If planid2 = "1" Then
                                    drShare("RID") = String.Concat(s_RID.Substring(0, 1), (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 1)))
                                    drShare("Relship") = String.Concat(RelShipID, s_RID.Substring(0, 1), (CInt(maxrid) + 1).ToString(New String("0", CInt(planid2) + 1)), "/")
                                    drShare("OrgLevel") = CInt(planid2) + 1
                                ElseIf planid2 > "1" Then
                                    If s_RID.Length > 2 Then '新增ex@id=c01底下的機構
                                        drShare("RID") = String.Concat(s_RID.Substring(0, 1), maxrid.Substring(0, maxrid.Length - 3), (CInt(Right(maxrid, 3)) + 1).ToString("000"))
                                        drShare("Relship") = String.Concat(RelShipID, s_RID.Substring(0, 1), maxrid.Substring(0, maxrid.Length - 3), (CInt(Right(maxrid, 3)) + 1).ToString("000"), "/")
                                    Else
                                        drShare("RID") = String.Concat(s_RID, (CInt(Right(maxrid, 3)) + 1).ToString("000"))
                                        drShare("Relship") = String.Concat(RelShipID, s_RID, (CInt(Right(maxrid, 3)) + 1).ToString("000"), "/")
                                    End If
                                    drShare("OrgLevel") = If(planid2 <> "3", CInt(planid2) + 1, 3)
                                End If
                            End If
                        End If
                    End If

                    'Auth_Relship之共用新增
                    drShare("PlanID") = TIMS.ClearSQM(PlanIDValue.Value)
                    drShare("DistID") = DistrictList.SelectedValue
                    drShare("ModifyAcct") = sm.UserInfo.UserID
                    drShare("ModifyDate") = Now()
                    dtShare.Rows.Add(drShare)

                    iRSID = DbAccess.GetNewId(oTrans, "AUTH_RELSHIP_RSID_SEQ,AUTH_RELSHIP,RSID")
                    ViewState(vs_RSID) = Convert.ToString(iRSID)
                    drShare("RSID") = iRSID
                    'Auth_Relship 檢查RID
                    Dim chkRID As String = Convert.ToString(drShare("RID"))

                    '=======================
                    ' 2018 add 記錄交易log (auth_relship ==> sys_trans_log)
                    '=======================
                    Call SaveAuthRelshipLog(oTrans, drShare, Rq_ProcessType)

                    DbAccess.UpdateDataTable(dtShare, sqlAdapter2, oTrans)
                    If ChkRID_ERROR(chkRID, 1, oConn, oTrans) Then
                        DbAccess.RollbackTrans(oTrans)
                        Call TIMS.CloseDbConn(oConn)
                        Dim strScript1 As String = String.Concat("<script language=""javascript"">", "alert('機構設定異常，請重新操作!!!');", "</script>")
                        Page.RegisterStartupScript("", strScript1)
                        Exit Sub
                    End If

                    'OrgPlanInfo之共用新增
                    Save_OrgPlanInfodr2(drOrgPlan, iRSID)
                    dtOrgPlan.Rows.Add(drOrgPlan)
                    '=======================
                    ' 2018 add 記錄交易log (org_orgplaninfo ==> sys_trans_log)
                    '=======================
                    Call SaveOrgOrgPlanInfoLog(oTrans, drOrgPlan, Rq_ProcessType)

                    DbAccess.UpdateDataTable(dtOrgPlan, sqlAdapter3, oTrans)
                    Dim strScript As String = String.Concat("<script language=""javascript"">", "alert('共用修改成功!!');", "location.href='TC_01_002.aspx?ID='+document.getElementById('Re_ID').value;", "</script>")
                    Page.RegisterStartupScript("", strScript)
            End Select
            DbAccess.CommitTrans(oTrans)

        Catch ex As Exception
            If oTrans IsNot Nothing Then DbAccess.RollbackTrans(oTrans)
            Dim ExErrMsg1 As String = String.Concat("訓練機構設定失敗!!", ex.Message)
            TIMS.WriteTraceLog(Me, ex, ExErrMsg1)
            Common.MessageBox(Page, ExErrMsg1)
            'Dim sErrMsg1 As String=""
            'sErrMsg1 &= TIMS.GetErrorMsgSys()
            'sErrMsg1 &= "ex.Message:" & ex.Message
            'sErrMsg1 &= ex.ToString
            'TIMS.WriteTraceLog(Nothing, ex, sErrMsg1)
            'TIMS.WriteTraceLog(Me.Page, ex, ex.Message)
            'Common.MessageBox(Page, ex.ToString)
            'Common.MessageBox(Page, "訓練機構設定失敗!!")
            'If Not oTrans Is Nothing Then DbAccess.RollbackTrans(oTrans)
            Call TIMS.CloseDbConn(oConn)
            save_flag_ok = False
            Return

        End Try
        Call TIMS.CloseDbConn(oConn)
        save_flag_ok = True
    End Sub

    ''' <summary> 檢查輸入10碼的統編 'TBID.Text HidComidno.Value  </summary>
    ''' <param name="Errmsg"></param>
    Sub Chk_First8C(ByRef Errmsg As String)
        '其他計畫要輸入正確的統編 或 正確輸入8碼的統編
        If HidComidno.Value = "" Then Errmsg += "請輸入統一編號" & vbCrLf
        If Len(HidComidno.Value) <> 8 Then Errmsg += "統一編號請填寫8位數字" & vbCrLf
        Dim flagIsValid As Boolean = True
        flagIsValid = True
        If HidComidno.Value = "00000000" Then flagIsValid = False
        If TIMS.isValidTWBID(HidComidno.Value) > 8 Then flagIsValid = False
        If Not flagIsValid Then
            HidComidno.Value = ""
            Errmsg += "前8碼，請輸入正確的統一編號" & vbCrLf
        End If
        If Errmsg <> "" Then Exit Sub '有一般錯誤，不再執行下列的檢查。
        '查看有無此機構代碼的建立 (8碼確認)
        If TIMS.Get_OrgIDforComIDNO(objconn, HidComidno.Value) = "" Then
            Errmsg += "前8碼，尚未建立該統編資料" & vbCrLf
        End If
    End Sub

    ''' <summary>
    ''' 都沒異常錯誤，執行資料庫面的檢查。
    ''' </summary>
    ''' <param name="Errmsg"></param>
    Sub Chk_DoubleX(ByRef Errmsg As String)
        If Errmsg <> "" Then Exit Sub '有一般錯誤，不再執行下列的檢查。

        Dim DoubleComIDNO_flag As Boolean = False
        Call TIMS.OpenDbConn(objconn)
        Select Case Rq_ProcessType
            Case cst_Update
                'strSqlchk=" SELECT 'x' FROM Org_OrgInfo a JOIN Auth_Relship b ON a.orgid=b.orgid WHERE a.OrgID <> '" & Re_orgid & "' AND ComIDNO='" & Me.TBID.Text & "' "
                Dim strSqlchk As String = " SELECT 'x' FROM Org_OrgInfo a JOIN Auth_Relship b ON a.orgid=b.orgid WHERE a.OrgID <> @OrgID AND ComIDNO=@ComIDNO "
                Dim sCmd As New SqlCommand(strSqlchk, objconn)
                Dim dt As New DataTable
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("OrgID", SqlDbType.VarChar).Value = Re_orgid
                    .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = Me.TBID.Text
                    dt.Load(.ExecuteReader())
                End With
                If dt.Rows.Count > 0 Then
                    DoubleComIDNO_flag = True
                    Errmsg += "機構設定重複!!!!" & vbCrLf
                End If
                If Not DoubleComIDNO_flag Then
                    '未重複做其他檢查。
                    '(修改統編)
                    If ViewState(vs_comidno) <> TBID.Text Then '統編有修改過
                        Dim strsql_IDNO As String = "" & vbCrLf
                        'Org_OrgInfo / plan_planinfo
                        strsql_IDNO = "" & vbCrLf
                        strsql_IDNO &= " SELECT 'X'" & vbCrLf
                        strsql_IDNO &= " FROM Org_OrgInfo a" & vbCrLf
                        strsql_IDNO &= " JOIN plan_planinfo b ON a.ComIDNO=b.ComIDNO" & vbCrLf
                        strsql_IDNO &= " WHERE b.ComIDNO=@ComIDNO" & vbCrLf
                        Dim sCmd2 As New SqlCommand(strsql_IDNO, objconn)
                        Dim dt2 As New DataTable
                        With sCmd2
                            .Parameters.Clear()
                            .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = ViewState(vs_comidno)
                            dt2.Load(.ExecuteReader())
                        End With
                        If dt2.Rows.Count > 0 Then
                            DoubleComIDNO_flag = True
                            Errmsg += "此機構已有申請計畫,因此不能修改統一編號!!!!" & vbCrLf
                        End If
                    End If
                End If
            Case cst_Insert, cst_InsertChk
                Dim strSqlchk As String = " SELECT 'x' FROM Org_OrgInfo a JOIN Auth_Relship b ON a.orgid=b.orgid WHERE ComIDNO=@ComIDNO "
                Dim sCmd As New SqlCommand(strSqlchk, objconn)
                Dim dt As New DataTable
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = Me.TBID.Text
                    dt.Load(.ExecuteReader())
                End With
                If dt.Rows.Count > 0 Then
                    '判斷是否有新增重複的資料
                    DoubleComIDNO_flag = True
                    Errmsg += "新增機構設定-統一編號重複!!!!" & vbCrLf
                End If
        End Select

        If DoubleComIDNO_flag Then
            Me.txtLastYearExeRate.Text = "0"
            Errmsg += "請重新計算" & LabLastYear2.Text & vbCrLf
            'Common.MessageBox(Me, Errmsg)
            'Exit Sub
        End If
    End Sub

    ''' <summary> 檢核問題。</summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim blnRst As Boolean = True
        Errmsg = ""

        'PlanIDValue
        Dim sTPlanID As String = TIMS.GetTPlanID(PlanIDValue.Value, objconn)
        If PlanIDValue.Value <> "" AndAlso PlanIDValue.Value <> "0" Then
            TIMS.LOG.Debug(String.Format("#PlanIDValue.{0},計畫選擇有誤!(請選擇隸屬機構)", PlanIDValue.Value))
            If sTPlanID = "" Then Errmsg += "計畫選擇有誤!(請選擇隸屬機構)" & vbCrLf
        End If

        ZipCODEB3.Value = TIMS.ClearSQM(ZipCODEB3.Value)

        TB_ActNo.Text = TIMS.ClearSQM(TB_ActNo.Text)
        TBtitle.Text = TIMS.ClearSQM(TBtitle.Text)
        TBID.Text = TIMS.ClearSQM(TBID.Text)

        HidComidno.Value = TIMS.ClearSQM(HidComidno.Value)
        TBseqno.Text = TIMS.ClearSQM(TBseqno.Text)
        TBaddress.Text = TIMS.ClearSQM(TBaddress.Text)
        TBm_name.Text = TIMS.ClearSQM(TBm_name.Text)
        TBm_Phone.Text = TIMS.ClearSQM(TBm_Phone.Text)
        TBContactName.Text = TIMS.ClearSQM(TBContactName.Text)
        TBtel.Text = TIMS.ClearSQM(TBtel.Text)
        TBmail.Text = TIMS.ChangeEmail(TIMS.ClearSQM(TBmail.Text))
        staffName.Text = TIMS.ClearSQM(staffName.Text)
        staffPhone.Text = TIMS.ClearSQM(staffPhone.Text)
        staffEmail.Text = TIMS.ClearSQM(staffEmail.Text)
        'TBseqno

        'chk_ActNo
        If TB_ActNo.Text <> "" AndAlso Len(TB_ActNo.Text) > 20 Then Errmsg += "保險證號請輸入20字以內" & vbCrLf
        If TBtitle.Text = "" Then Errmsg += "請輸入機構名稱全銜" & vbCrLf
        If OrgKindList.SelectedValue = "" Then Errmsg += "請選擇機構別" & vbCrLf
        TBID.Text = TIMS.ChangeIDNO(TBID.Text) '整理
        HidComidno.Value = "" '存入8碼 統編 檢查使用。
        If Len(TBID.Text) = 8 Then HidComidno.Value = TBID.Text 'HidComidno.Value 
        If Len(TBID.Text) = 10 Then HidComidno.Value = TBID.Text.Substring(0, 8) 'HidComidno.Value 永遠只能入前8碼喔   '產投提供10碼的輸入 Then
        If Not TIMS.Check123(TBID.Text) Then Errmsg += "統一編號請填寫數字" & vbCrLf
        If sTPlanID <> "" AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 AndAlso Len(TBID.Text) = 10 Then
            '是產投計畫而且輸入了10碼資訊。 產投可輸入10碼的統編
            '檢查輸入10碼的統編 'TBID.Text HidComidno.Value 
            Call Chk_First8C(Errmsg)
        Else
            '其他計畫要輸入正確的統編 或 正確輸入8碼的統編
            If HidComidno.Value = "" Then Errmsg += "請輸入統一編號" & vbCrLf
            If Len(HidComidno.Value) <> 8 Then Errmsg += "統一編號請填寫8位數字" & vbCrLf
            Dim flagIsValid As Boolean = True
            flagIsValid = True
            If HidComidno.Value = "00000000" Then flagIsValid = False '(假的)

            If Len(RIDValue.Value) <> 1 Then
                If flagIsValid AndAlso Not TIMS.isValidTWBID(HidComidno.Value) Then flagIsValid = False
                If flagIsValid AndAlso Len(HidComidno.Value) <> 8 Then flagIsValid = False
                If Not flagIsValid Then
                    HidComidno.Value = ""
                    Errmsg += "請輸入正確的統一編號" & vbCrLf
                End If
            End If
        End If
        If TBseqno.Text = "" Then Errmsg += "請輸入 立案登記編號" & vbCrLf
        If RIDValue.Value = "" OrElse TBplan.Text = "" Then Errmsg += "請選擇管控單位" & vbCrLf
        If TBCity.Text = "" Then Errmsg += "請選擇縣市" & vbCrLf
        If TBaddress.Text = "" Then Errmsg += "請輸入地址" & vbCrLf
        Call TIMS.CheckZipCODEB3(ZipCODEB3.Value, "郵遞區號 後2碼或後3碼", True, Errmsg)

        'If Len(RIDValue.Value) <> 1 Then If TBm_name.Text="" Then Errmsg += "請輸入負責人姓名" & vbCrLf
        If TBm_name.Text = "" Then Errmsg += "請輸入負責人姓名" & vbCrLf
        If TBContactName.Text = "" Then Errmsg += "請輸入聯絡人姓名" & vbCrLf
        If TBtel.Text = "" Then Errmsg += "請輸入聯絡人電話" & vbCrLf
        'If TBmail.Text="" Then Errmsg += "請輸入 聯絡人E-MAIL" & vbCrLf
        If TBmail.Text <> "" AndAlso Not TIMS.CheckEmail(TBmail.Text) Then
            If Not TIMS.CheckEmail(TBmail.Text) Then Errmsg += "(EMAIL格式有誤)請重新輸入 聯絡人E-MAIL" & vbCrLf
        End If
        If TrTPlanID28F1.Visible = True Then
            If staffName.Text = "" Then Errmsg += "個人資料檔案保管人員 為必填欄位" & vbCrLf
            If staffPhone.Text = "" Then Errmsg += "個人資料檔案保管人員電話 為必填欄位" & vbCrLf
        End If
        If TrTPlanID28F2.Visible = True Then
            If staffEmail.Text = "" Then Errmsg += "個人資料檔案保管人員電子郵件 為必填欄位" & vbCrLf
            If staffEmail.Text <> "" Then
                If Not TIMS.CheckEmail(staffEmail.Text) Then Errmsg += "請重新輸入 個人資料檔案保管人員電子郵件" & vbCrLf
            End If
        End If
        If TB_ActNo.Text = "" Then Errmsg += "請輸入保險證號" & vbCrLf
        'If TB_ProTrainKind.Text="" Then Errmsg += "請輸入專長訓練職類" & vbCrLf
        'If ComSumm.Text="" Then Errmsg += "請輸入機構簡介" & vbCrLf
        If MemberNum.Text <> "" AndAlso Not TIMS.Check123(MemberNum.Text) Then Errmsg += "「會員人數」請輸入數字" & vbCrLf
        If ActMemberNum.Text <> "" AndAlso Not TIMS.Check123(ActMemberNum.Text) Then Errmsg += "「勞工保險加保人數-會員人數」請輸入數字" & vbCrLf
        If ActStaffNum.Text <> "" AndAlso Not TIMS.Check123(ActStaffNum.Text) Then Errmsg += "「勞工保險加保人數-員工人數」請輸入數字" & vbCrLf

        '共用-不顯示 (新增／修改)顯示
        Select Case Rq_ProcessType
            Case cst_Insert ', cst_Update
                If sTPlanID <> "" AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                    If dl_typeid1.SelectedValue = "" Then Errmsg += "請選擇 訓練機構屬性設定-計畫別" & vbCrLf
                    If dl_typeid2.SelectedValue = "" Then Errmsg += "請選擇 訓練機構屬性設定-機構別" & vbCrLf
                End If
        End Select

        If Errmsg = "" Then Call Chk_DoubleX(Errmsg)
        If Errmsg <> "" Then blnRst = False
        Return blnRst
    End Function

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        '前端驗證是否有錯誤。
        'Dim rstPageIsValid As Boolean=True
        'If Not rstPageIsValid Then
        '    'args.Value'Summary.HeaderText 
        '    Common.MessageBox(Page, "(前端驗證資料錯誤)儲存失敗!!")
        '    Return 'Exit Sub
        'End If

        Dim gflag_test As Boolean = True 'true:測試環境。(false:正式環境) / TestStr
        gflag_test = False
        If TIMS.sUtl_ChkTest() Then gflag_test = True '測試用--取消測試環境參數，即可啟用

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return ' Exit Sub
        End If

        'Dim slogMsg1 As String=""
        'slogMsg1="##TC_01_002_add"
        'slogMsg1 &= ",ViewState(vs_RSID):" & ViewState(vs_RSID) & vbCrLf
        'slogMsg1 &= ",gFlagEnv:" & gFlagEnv & vbCrLf
        'TIMS.writeLog(Me, slogMsg1)
        Dim save_flag_ok As Boolean = False '儲存狀況: true:正常/false:異常

        '儲存
        'If gflag_test Then
        '    Common.MessageBox(Page, "(測試環境)--停止儲存!!")
        '    Return 'Exit Sub
        'End If
        'If Not gflag_test Then Call Insert_Auth_Relship(save_flag_ok) '儲存-正式環境。(測試用) / TestStr
        Dim lockobj As New Object
        SyncLock lockobj
            Call Insert_Auth_Relship(save_flag_ok) '儲存-正式環境。
        End SyncLock

        If Not save_flag_ok Then Return '儲存產生異常，直接離開

        Dim thisLock As New Object
        SyncLock thisLock
            Call SaveACCTORG()
        End SyncLock
    End Sub

    ''' <summary>
    ''' 分署管理者，處理業務序號
    ''' </summary>
    Sub SaveACCTORG()
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then Return
        'sm.UserInfo.LID=0/1 
        If ViewState(vs_RSID) Is Nothing Then Return
        '從字串 "" 至類型 'Double' 的轉換是無效的
        If Convert.ToString(ViewState(vs_RSID)) = "" Then Return

        Const Cst_二級 As Integer = 4
        Const Cst_承辦人 As Integer = 5
        Dim i_RSID As Integer = Val(ViewState(vs_RSID))
        'Dim slogMsg1 As String=""
        'slogMsg1="##TC_01_002_add-SaveACCTORG"
        'slogMsg1 &= ",ViewState(vs_RSID):" & ViewState(vs_RSID).ToString() & vbCrLf
        'slogMsg1 &= ",i_RSID:" & i_RSID.ToString() & vbCrLf
        'slogMsg1 &= ",PlanIDValue.Value:" & PlanIDValue.Value & vbCrLf
        'slogMsg1 &= ",sm.UserInfo.TPlanID:" & sm.UserInfo.TPlanID & vbCrLf
        'TIMS.writeLog(Me, slogMsg1)
        TIMS.Update_AUTH_ACCTORG(PlanIDValue.Value, i_RSID, Cst_承辦人, objconn)
        TIMS.Update_AUTH_ACCTORG(PlanIDValue.Value, i_RSID, Cst_二級, objconn)
    End Sub

    ''' <summary>
    ''' 階層選擇
    ''' </summary>
    Sub SUtl_levellistSel()
        Dim v_level_list As String = TIMS.GetListValue(level_list)
        Dim iv_level_list As Integer = Val(v_level_list)

        If iv_level_list = 1 OrElse iv_level_list = 0 Then
            Dim myValue As String = TIMS.Get_DistName1("000")
            Me.choice_button.Disabled = True
            'Me.TBplan.ReadOnly=True
            Me.TBplan.Text = myValue '"職訓局"
            Me.PlanIDValue.Value = "0"
            Me.RIDValue.Value = "A"
            TIMS.Tooltip(choice_button, myValue)
            TIMS.Tooltip(TBplan, myValue)
        ElseIf iv_level_list = 2 Then
            Me.choice_button.Disabled = False
            'Me.TBplan.ReadOnly=True
            Me.TBplan.Text = ""
            Me.PlanIDValue.Value = ""
            Me.RIDValue.Value = ""
        End If
        btn_clear.Disabled = choice_button.Disabled
    End Sub

    ''' <summary>
    ''' 指定的數值是否通過驗證。
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="args"></param>
    Private Sub CustomValidator1_ServerValidate(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles CustomValidator1.ServerValidate
        If Me.TBID.Text = "00000000" Then
            args.IsValid = False
            Common.MessageBox(Page, source.errormessage)
            Exit Sub
        End If
        args.IsValid = True
    End Sub

    ''' <summary>
    ''' 回上一頁
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim url1 As String = ""
        If Me.ViewState(vs_Redirect) Is Nothing Then
            If Session("_Search") IsNot Nothing Then Session("_Search") = Session("_Search")
            url1 = "TC_01_002.aspx?ID=" & Request("ID") & ""
        Else
            url1 = Me.ViewState(vs_Redirect) & "?ID=" & Request("ID") & ""
        End If
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    ''' <summary>
    ''' 計算年度核定人次執行率 ExeRate_Click
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ExeRate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExeRate.Click
        Dim Errmsg As String = ""
        If TBID.Text = "" Then Errmsg += "統一編號不可為空" & vbCrLf
        If LastYear.Value = "" Then Errmsg += "年度不可為空" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        txtLastYearExeRate.Text = FindExeRate(TBID.Text, LastYear.Value)
        If txtLastYearExeRate.Text = "" Then
            Common.MessageBox(Me, LabLastYear2.Text & "，無此資料，請確認!!")
            txtLastYearExeRate.Text = "0"
        End If
        Call Set_GWOrgKind2(Re_orgid)
        'Set_GWOrgKind2(sm.UserInfo.OrgID.ToString)
    End Sub

    ''' <summary>
    ''' 計算年度核定人次執行率(產投專用)
    ''' </summary>
    ''' <param name="ComIDNO"></param>
    ''' <param name="PlanYear"></param>
    ''' <returns></returns>
    Function FindExeRate(ByVal ComIDNO As String, ByVal PlanYear As String) As String
        Dim sRate As String = ""
        Dim sql As String = ""
        sql &= " SELECT SUM(CS.StudentCount) ClosedStdCnt ,SUM(P1.TNum) TNum" & vbCrLf
        sql &= "  ,CASE WHEN SUM(P1.TNum)>0 THEN ROUND(SUM(CS.StudentCount)/SUM(P1.TNum)*100,4) ELSE ROUND(0,4) END ExeRate "
        sql &= " FROM CLASS_CLASSINFO c1" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO p1 ON c1.Planid=p1.Planid AND c1.ComIDNO=p1.ComIDNO AND c1.SeqNo=p1.SeqNo AND p1.TPlanID IN (" & TIMS.Cst_TPlanID28_2 & ")" & vbCrLf
        sql &= " JOIN (" & vbCrLf
        sql &= "  SELECT OCID, COUNT(1) StudentCount" & vbCrLf
        sql &= "  FROM Class_StudentsOfClass" & vbCrLf
        sql &= "  WHERE StudStatus IN (5)" & vbCrLf
        sql &= "  GROUP BY ocid) cs ON c1.ocid=cs.ocid" & vbCrLf
        sql &= " WHERE c1.COMIDNO='" & ComIDNO & "'" & vbCrLf
        sql &= "  AND p1.PlanYear='" & PlanYear & "'" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr("ExeRate").ToString <> "" And IsNumeric(dr("ExeRate").ToString) Then
            '100.00
            sRate = TIMS.ROUND(dr("ExeRate").ToString, 2)
        Else
            sRate = dr("ExeRate").ToString
        End If
        Return sRate
    End Function

    ''' <summary>
    ''' 20080818 andy 可修改機構名稱
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Bt_chgOrg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_CHGORG.Click
        Dim MyValue1 As String = ""
        MyValue1 = "ProcessType=Share"
        MyValue1 &= "&orgid=" & TIMS.ClearSQM(Request("orgid"))
        MyValue1 &= "&distid=" & TIMS.ClearSQM(Request("distid"))
        MyValue1 &= "&planid=" & TIMS.ClearSQM(Request("planid"))
        MyValue1 &= "&rid=" & TIMS.ClearSQM(Request("rid"))
        MyValue1 &= "&ID=" & TIMS.ClearSQM(Request("ID"))

        Dim strScript As String
        strScript = String.Format("<script>window.open('../../Common/Chg_OrgName.aspx?{0}&BackPage=../TC/01/TC_01_002.aspx','','scrollbars=yes,width=1040,height=555'); </script>", MyValue1)
        Me.Page.RegisterStartupScript("", strScript)
        CustomValidator2.Enabled = True
        RequiredFieldValidator10.Enabled = True
    End Sub

    Private Sub Bt_chgOrg_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles BT_CHGORG.PreRender
        CustomValidator2.Enabled = False
        'RequiredFieldValidator10.Enabled=False
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DataGrid2.PageSize * DataGrid2.CurrentPageIndex
                If flag_ROC Then e.Item.Cells(2).Text = (Convert.ToInt32(drv("PlanYear")) - 1911).ToString  'edit，by:20181001
        End Select
    End Sub

    ''' <summary>
    ''' 顯示開班歷史。
    ''' </summary>
    ''' <param name="ComIDNO_Val"></param>
    ''' <param name="TPlanID_Val"></param>
    ''' <param name="Years_Val"></param>
    Sub SearchHistory(ByVal ComIDNO_Val As String, Optional ByVal TPlanID_Val As String = "", Optional ByVal Years_Val As String = "")
        DataGrid2.CurrentPageIndex = 0
        If ComIDNO_Val = "" Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.OCID ,ip.DistID ,pp.TPlanID ,kp.PlanName ,pp.ComIDNO " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,pp.PlanYear ,pp.STDate ,pp.FDDate ,pp.TMID ,idt.Name DistName " & vbCrLf
        sql &= " ,CASE when K2.JobID IS NULL THEN K2.TrainName ELSE K2.JobName END TrainName" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, pp.STDate, 111) + '<BR>|<BR>' + CONVERT(VARCHAR, pp.FDDate, 111) TRound" & vbCrLf
        sql &= " FROM Class_ClassInfo cc" & vbCrLf
        sql &= " JOIN plan_planinfo pp ON cc.planid=pp.planid AND cc.comidno=pp.comidno AND cc.seqno=pp.seqno AND cc.rid=pp.rid" & vbCrLf
        sql &= " JOIN ID_Plan ip ON pp.PlanID=ip.PlanID" & vbCrLf
        sql &= " JOIN Key_Plan kp ON pp.TPlanID=kp.TPlanID" & vbCrLf
        sql &= " JOIN Key_TrainType K2 ON pp.TMID=K2.TMID" & vbCrLf
        sql &= " JOIN ID_District idt ON ip.DistID=idt.DistID" & vbCrLf
        sql &= " WHERE pp.AppliedResult='Y'" & vbCrLf
        sql &= " AND pp.IsApprPaper='Y'" & vbCrLf
        sql &= " AND cc.notopen='N'" & vbCrLf
        If ComIDNO_Val <> "" Then
            'SearchStr1="" & vbCrLf
            sql &= " AND pp.ComIDNO='" & ComIDNO_Val & "'" & vbCrLf
        End If
        If TPlanID_Val <> "" Then
            sql &= " AND ip.TPlanID='" & TPlanID_Val & "'" & vbCrLf
        End If
        If Years_Val <> "" Then
            sql &= " AND pp.PlanYear='" & Years_Val & "'" & vbCrLf
        End If

        'sql &= SearchStr1
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!"
        'HistoryTable.Style.Item("display")="none"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            'RecordCount.Text=dt.Rows.Count
            'HistoryTable.Style.Item("display")="inline"
            If Me.ViewState(vs_sort) Is Nothing Then Me.ViewState(vs_sort) = "PlanYear,TRound,ClassName,DistID"
            dt.DefaultView.Sort = Me.ViewState(vs_sort)
            'PageControler1.PageDataTable=dt
            'PageControler1.Sort=Me.ViewState(vs_sort) ' "IDNO,Birthday,TRound"
            'PageControler1.ControlerLoad()
            DataGrid2.DataSource = dt
            DataGrid2.DataBind()
            dt.Dispose()
            dt = Nothing
        End If
    End Sub

    ''' <summary>
    ''' 檢查RID正常性
    ''' </summary>
    ''' <param name="RID"></param>
    ''' <param name="CNT"></param>
    ''' <param name="oConn"></param>
    ''' <param name="oTrans"></param>
    ''' <returns></returns>
    Function ChkRID_ERROR(ByVal RID As String, ByVal CNT As Integer, ByVal oConn As SqlConnection, ByRef oTrans As SqlTransaction) As Boolean
        Dim rst As Boolean = False '預設為正常
        Dim sql As String = ""
        sql = " SELECT * FROM AUTH_RELSHIP WHERE RID=@RID"
        Dim sCmd As New SqlCommand(sql, oConn, oTrans)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = RID
            dt.Load(.ExecuteReader())
        End With
        '不等同傳入數字
        If dt.Rows.Count <> CNT Then rst = True '異常
        Return rst
    End Function

    '2018 merge TC_01_017 訓練機構屬性設定(for 產投 & 充飛計畫用)

    ''' <summary>
    ''' 查詢取得訓練機構屬性設定結果
    ''' </summary>
    ''' <param name="orgTypeID1"></param>
    ''' <returns></returns>
    Private Function GetOrgType1ByPK(ByVal ORGTYPEID1 As String) As DataRow
        Dim parms As New Hashtable From {{"ORGTYPEID1", ORGTYPEID1}}
        Dim sql As String = "SELECT TYPEID1 ,TYPEID2 FROM KEY_ORGTYPE1 WHERE ORGTYPEID1=@ORGTYPEID1"
        Return DbAccess.GetOneRow(sql, objconn, parms)
    End Function

    ''' <summary>
    ''' 依設定結果回查機構屬性設定key值(key_orgtype1.orgtypeid1)
    ''' </summary>
    ''' <param name="typeID1"></param>
    ''' <param name="typeID2"></param>
    ''' <returns></returns>
    Private Function GetOrgType1ByTypeID(TYPEID1 As String, TYPEID2 As String) As DataRow
        Dim parms As New Hashtable From {{"TYPEID1", TYPEID1}, {"TYPEID2", TYPEID2}}
        Dim sql As String = " SELECT ORGTYPEID1 ,TYPEID1 ,TYPEID2 FROM KEY_ORGTYPE1 WHERE TYPEID1=@TYPEID1 AND TYPEID2=@TYPEID2 "
        Return DbAccess.GetOneRow(sql, objconn, parms)
    End Function

    ''' <summary>
    ''' 代入訓練機構屬性設定-計畫別DropDownList資料
    ''' </summary>
    ''' <param name="TypeID1"></param>
    Private Sub Get_ddl_typeid2(ByVal TypeID1 As String)
        If TypeID1 = "" OrElse TypeID1 = "0" Then Exit Sub
        Dim dt As New DataTable
        Dim sSql As String = ""
        sSql &= " SELECT TypeID1 ,TypeID2 , concat(TypeID2,'-',TypeID2Name) TypeID2Name" & vbCrLf
        sSql &= " FROM Key_OrgType1" & vbCrLf
        sSql &= " WHERE TypeID1=@TypeID1" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sSql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("@TypeID1", SqlDbType.Int).Value = CInt(TypeID1)
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            dl_typeid2.Items.Clear()
            dl_typeid2.DataSource = dt
            dl_typeid2.DataValueField = "TypeID2"
            dl_typeid2.DataTextField = "TypeID2Name"
            dl_typeid2.DataBind()
            dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
            dl_typeid2.SelectedIndex = 0
            dl_typeid2.Enabled = True
        Else
            dl_typeid2.Items.Clear()
            dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
            dl_typeid2.SelectedIndex = 0
            dl_typeid2.Enabled = False
        End If
    End Sub

    ''' <summary>顯示機構屬性設定結果</summary>
    ''' <param name="orgKind1"></param>
    Private Sub SetOrgKind1(ByVal orgKind1 As String)
        If orgKind1 = "" Then Exit Sub
        Dim orgtype1Dr As DataRow = GetOrgType1ByPK(orgKind1)
        If orgtype1Dr IsNot Nothing Then
            Common.SetListItem(dl_typeid1, Convert.ToString(orgtype1Dr("typeid1")))
            Get_ddl_typeid2(Convert.ToString(orgtype1Dr("typeid1")))
            Common.SetListItem(dl_typeid2, Convert.ToString(orgtype1Dr("typeid2")))
        End If
    End Sub

    ''' <summary>
    ''' 計畫別下拉連動顯示機構別下拉選項
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Dl_typeid1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dl_typeid1.SelectedIndexChanged
        If dl_typeid1.SelectedIndex = 0 Then
            dl_typeid2.Items.Clear()
            dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
            dl_typeid2.SelectedIndex = 0
            dl_typeid2.Enabled = False
        Else
            Get_ddl_typeid2(dl_typeid1.SelectedValue)
        End If
    End Sub

#Region "SAVE LOG"'2018 寫入SYS_TRANS_LOG'

    ''' <summary>
    ''' 儲存org_orginfo交易log (insert/update/delete)
    ''' </summary>
    ''' <param name="objTrans"></param>
    ''' <param name="dr"></param>
    ''' <param name="processType"></param>
    Private Sub SaveOrgOrgInfoLog(ByVal objTrans As SqlTransaction, ByVal dr As DataRow, ByVal processType As String)
        Dim BeforeValues As String = ""
        Dim AfterValues As String = ""
        Dim t_iSql As String = ""
        Dim myParam As Hashtable = New Hashtable
        Select Case processType
            Case cst_Insert, cst_InsertChk
                '新增作業
                '==========
                BeforeValues = ""
                BeforeValues += "ORGID=" + Convert.ToString(dr("ORGID"))
                BeforeValues += ",ORGKIND=" + Convert.ToString(dr("OrgKind"))
                BeforeValues += ",ORGNAME=" + Convert.ToString(dr("OrgName"))
                BeforeValues += ",COMIDNO=" + Convert.ToString(dr("ComIDNO"))
                BeforeValues += ",COMCIDNO=" + Convert.ToString(dr("ComCIDNO"))
                BeforeValues += ",ISCONUNIT=" + Convert.ToString(dr("IsConUnit"))
                BeforeValues += ",ORGKIND2=" + Convert.ToString(dr("OrgKind2"))
                BeforeValues += ",LASTYEAREXERATE=" + Convert.ToString(dr("LastYearExeRate"))
                BeforeValues += ",ISCONTTQS=" + Convert.ToString(dr("IsConTTQS"))
                BeforeValues += ",BANKNAME=" + Convert.ToString(dr("BankName"))
                BeforeValues += ",EXBANKNAME=" + Convert.ToString(dr("ExBankName"))
                BeforeValues += ",ACCNO=" + Convert.ToString(dr("AccNo"))
                BeforeValues += ",ACCNAME=" + Convert.ToString(dr("AccName"))
                '訓練機構屬性設定
                BeforeValues += ",ORGKIND1=" + Convert.ToString(dr("orgkind1"))
                BeforeValues += ",ORGZIPCODE=" + Convert.ToString(dr("orgzipcode"))
                BeforeValues += ",ORGZIPCODE6W=" + Convert.ToString(dr("ORGZIPCODE6W"))
                BeforeValues += ",ORGADDRESS=" + Convert.ToString(dr("orgaddress"))
                BeforeValues += ",MODIFYACCT=" + Convert.ToString(dr("ModifyAcct"))
                BeforeValues += ",MODIFYDATE=" + Convert.ToDateTime(Convert.ToString(dr("ModifyDate"))).ToString("yyyy-MM-dd HH:mm:ss.fff")
                '==========
                t_iSql &= " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
                t_iSql &= " VALUES(@SessionID, @TransTime, '/TC/01/TC_01_002_add.aspx', @UserID, 'Insert', 'ORG_ORGINFO', '', @BeforeValues, '') "

                myParam.Clear()
                myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                myParam.Add("SessionID", sm.SessionID.ToString)
                myParam.Add("UserID", sm.UserInfo.UserID)
                myParam.Add("BeforeValues", BeforeValues)
                DbAccess.ExecuteNonQuery(t_iSql, objTrans, myParam)
            Case cst_Update
                '修改作業
                Dim t_BeforeSql As String = " SELECT * FROM ORG_ORGINFO WHERE ORGID=@ORGID "
                Dim t_BeforeTB As DataTable = Nothing
                Dim t_BeforeRow As DataRow = Nothing
                myParam.Add("ORGID", dr("orgid"))
                t_BeforeTB = DbAccess.GetDataTable(t_BeforeSql, objconn, myParam)
                If t_BeforeTB IsNot Nothing Then
                    t_BeforeRow = t_BeforeTB.Rows(0)
                    '異動前資料
                    BeforeValues = ""
                    BeforeValues += "ORGKIND=" + Convert.ToString(t_BeforeRow("OrgKind"))
                    BeforeValues += ",ORGNAME=" + Convert.ToString(t_BeforeRow("OrgName"))
                    BeforeValues += ",COMIDNO=" + Convert.ToString(t_BeforeRow("ComIDNO"))
                    BeforeValues += ",COMCIDNO=" + Convert.ToString(t_BeforeRow("ComCIDNO"))
                    BeforeValues += ",ISCONUNIT=" + Convert.ToString(t_BeforeRow("IsConUnit"))
                    BeforeValues += ",ORGKIND2=" + Convert.ToString(t_BeforeRow("OrgKind2"))
                    BeforeValues += ",LASTYEAREXERATE=" + Convert.ToString(t_BeforeRow("LastYearExeRate"))
                    BeforeValues += ",ISCONTTQS=" + Convert.ToString(t_BeforeRow("IsConTTQS"))
                    BeforeValues += ",BANKNAME=" + Convert.ToString(t_BeforeRow("BankName"))
                    BeforeValues += ",EXBANKNAME=" + Convert.ToString(t_BeforeRow("ExBankName"))
                    BeforeValues += ",ACCNO=" + Convert.ToString(t_BeforeRow("AccNo"))
                    BeforeValues += ",ACCNAME=" + Convert.ToString(t_BeforeRow("AccName"))
                    '訓練機構屬性設定
                    BeforeValues += ",ORGKIND1=" + Convert.ToString(t_BeforeRow("orgkind1"))
                    BeforeValues += ",ORGZIPCODE=" + Convert.ToString(t_BeforeRow("orgzipcode"))
                    BeforeValues += ",ORGZIPCODE6W=" + Convert.ToString(t_BeforeRow("ORGZIPCODE6W"))
                    BeforeValues += ",ORGADDRESS=" + Convert.ToString(t_BeforeRow("orgaddress"))
                    BeforeValues += ",MODIFYACCT=" + Convert.ToString(t_BeforeRow("ModifyAcct"))
                    BeforeValues += ",MODIFYDATE=" + Convert.ToDateTime(Convert.ToString(t_BeforeRow("MODIFYDATE"))).ToString("yyyy-MM-dd HH:mm:ss.fff")
                    '異動後資料
                    AfterValues = ""
                    AfterValues += "ORGKIND=" + Convert.ToString(dr("OrgKind"))
                    AfterValues += ",ORGNAME=" + Convert.ToString(dr("OrgName"))
                    AfterValues += ",COMIDNO=" + Convert.ToString(dr("ComIDNO"))
                    AfterValues += ",COMCIDNO=" + Convert.ToString(dr("ComCIDNO"))
                    AfterValues += ",ISCONUNIT=" + Convert.ToString(dr("IsConUnit"))
                    AfterValues += ",ORGKIND2=" + Convert.ToString(dr("OrgKind2"))
                    AfterValues += ",LASTYEAREXERATE=" + Convert.ToString(dr("LastYearExeRate"))
                    AfterValues += ",ISCONTTQS=" + Convert.ToString(dr("IsConTTQS"))
                    AfterValues += ",BANKNAME=" + Convert.ToString(dr("BankName"))
                    AfterValues += ",EXBANKNAME=" + Convert.ToString(dr("ExBankName"))
                    AfterValues += ",ACCNO=" + Convert.ToString(dr("AccNo"))
                    AfterValues += ",AccName=" + Convert.ToString(dr("AccName"))
                    '訓練機構屬性設定
                    AfterValues += ",ORGKIND1=" + Convert.ToString(dr("orgkind1"))
                    AfterValues += ",ORGZIPCODE=" + Convert.ToString(dr("orgzipcode"))
                    AfterValues += ",ORGZIPCODE6W=" + Convert.ToString(dr("ORGZIPCODE6W"))
                    AfterValues += ",ORGADDRESS=" + Convert.ToString(dr("orgaddress"))
                    AfterValues += ",ModifyAcct=" + Convert.ToString(dr("ModifyAcct"))
                    AfterValues += ",MODIFYDATE=" + Convert.ToDateTime(Convert.ToString(dr("ModifyDate"))).ToString("yyyy-MM-dd HH:mm:ss.fff")
                    Dim Conditions As String = ""
                    Conditions += ("ORGID=" + Convert.ToString(dr("orgid")))
                    '==========
                    t_iSql &= " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
                    t_iSql &= " VALUES(@SessionID, @TransTime, '/TC/01/TC_01_002_add.aspx', @UserID, 'Update', 'ORG_ORGINFO',  @Conditions, @BeforeValues, @AfterValues) "

                    myParam.Clear()
                    myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                    myParam.Add("SessionID", sm.SessionID.ToString)
                    myParam.Add("UserID", sm.UserInfo.UserID)
                    myParam.Add("Conditions", Conditions)
                    myParam.Add("BeforeValues", BeforeValues)
                    myParam.Add("AfterValues", AfterValues)
                    DbAccess.ExecuteNonQuery(t_iSql, objTrans, myParam)
                End If
        End Select
    End Sub

    ''' <summary>
    ''' 記錄org_orgplaninfo交易log (insert/update/delete)
    ''' </summary>
    ''' <param name="objTrans"></param>
    ''' <param name="dr"></param>
    ''' <param name="processType"></param>
    Sub SaveOrgOrgPlanInfoLog(ByVal objTrans As SqlTransaction, ByVal dr As DataRow, ByVal processType As String)
        Dim s_TransType As String = TIMS.cst_TRANS_LOG_Insert 'insert:cst_TRANS_LOG_Insert/update:cst_TRANS_LOG_Update
        Dim s_TargetTable As String = "ORG_ORGPLANINFO"
        Dim s_FuncPath As String = "/TC/01/TC_01_002_add"
        Const cst_fWHERE As String = "RSID={0}"
        Dim s_WHERE As String = "" 'insert省略/update必要'String.Format(cst_fWHERE, pkVALUE)

        Select Case processType
            Case cst_Insert, cst_InsertChk, cst_Share
                '新增作業
                Dim htPP As New Hashtable From {
                    {"TransType", s_TransType},
                    {"TargetTable", s_TargetTable},
                    {"FuncPath", s_FuncPath}
                }
                'htPP.Add("s_WHERE", s_WHERE)
                TIMS.SaveTRANSLOG(sm, Nothing, objTrans, dr, htPP)

            Case cst_Update
                '修改作業
                Dim s_RSID As String = dr("rsid").ToString()
                s_TransType = TIMS.cst_TRANS_LOG_Update 'insert:cst_TRANS_LOG_Insert/update:cst_TRANS_LOG_Update
                s_WHERE = String.Format(cst_fWHERE, s_RSID) 'insert省略/update必要

                Dim htPP As New Hashtable From {
                    {"TransType", s_TransType},
                    {"TargetTable", s_TargetTable},
                    {"FuncPath", s_FuncPath},
                    {"s_WHERE", s_WHERE}
                }
                TIMS.SaveTRANSLOG(sm, Nothing, objTrans, dr, htPP)

        End Select
    End Sub

    ''' <summary>
    ''' 記錄 auth_relship 交易log (insert/update/delete)
    ''' </summary>
    ''' <param name="objTrans"></param>
    ''' <param name="dr"></param>
    ''' <param name="processType"></param>
    Sub SaveAuthRelshipLog(ByVal objTrans As SqlTransaction, ByVal dr As DataRow, ByVal processType As String)
        Dim BeforeValues As String = ""
        Dim AfterValues As String = ""
        Dim t_iSql As String = ""
        Dim myParam As Hashtable = New Hashtable
        Select Case processType
            Case cst_Insert, cst_InsertChk, cst_Share
                '新增作業
                BeforeValues = ""
                BeforeValues += "RSID=" + Convert.ToString(dr("rsid"))
                BeforeValues += ",PLANID=" + Convert.ToString(dr("planid"))
                BeforeValues += ",RID=" + Convert.ToString(dr("rid"))
                BeforeValues += ",ORGID=" + Convert.ToString(dr("OrgID"))
                BeforeValues += ",RELSHIP=" + Convert.ToString(dr("relship"))
                BeforeValues += ",ORGLEVEL=" + Convert.ToString(dr("OrgLevel"))
                BeforeValues += ",DISTID=" + Convert.ToString(dr("DistID"))
                BeforeValues += ",ModifyAcct=" + Convert.ToString(dr("ModifyAcct"))
                BeforeValues += ",MODIFYDATE=" + Convert.ToDateTime(dr("ModifyDate")).ToString("yyyy-MM-dd HH:mm:ss.fff")
                '==========
                t_iSql &= " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
                t_iSql &= " VALUES(@SessionID, @TransTime, '/TC/01/TC_01_002_add.aspx', @UserID, 'Insert', 'AUTH_RELSHIP', '', @BeforeValues, '') "
                myParam.Clear()
                myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                myParam.Add("SessionID", sm.SessionID.ToString)
                myParam.Add("UserID", sm.UserInfo.UserID)
                myParam.Add("BeforeValues", BeforeValues)
                '==========
                DbAccess.ExecuteNonQuery(t_iSql, objTrans, myParam)
            Case cst_Update
                '修改作業
                Dim t_BeforeSql As String = " SELECT * FROM AUTH_RELSHIP WHERE RSID=@RSID "
                Dim t_BeforeTB As DataTable = Nothing
                Dim t_BeforeRow As DataRow = Nothing
                myParam.Add("RSID", dr("rsid"))
                t_BeforeTB = DbAccess.GetDataTable(t_BeforeSql, objconn, myParam)
                If t_BeforeTB IsNot Nothing AndAlso t_BeforeTB.Rows.Count > 0 Then
                    t_BeforeRow = t_BeforeTB.Rows(0)
                    '異動前資料
                    BeforeValues = ""
                    BeforeValues += "PLANID=" + Convert.ToString(t_BeforeRow("planid"))
                    BeforeValues += ",RID=" + Convert.ToString(t_BeforeRow("rid"))
                    BeforeValues += ",ORGID=" + Convert.ToString(t_BeforeRow("OrgID"))
                    BeforeValues += ",RELSHIP=" + Convert.ToString(t_BeforeRow("relship"))
                    BeforeValues += ",ORGLEVEL=" + Convert.ToString(t_BeforeRow("OrgLevel"))
                    BeforeValues += ",DISTID=" + Convert.ToString(t_BeforeRow("DistID"))
                    BeforeValues += ",MODIFYACCT=" + Convert.ToString(t_BeforeRow("ModifyAcct"))
                    BeforeValues += ",MODIFYDATE=" + Convert.ToDateTime(t_BeforeRow("ModifyDate")).ToString("yyyy-MM-dd HH:mm:ss.fff")
                    '異動後資料
                    AfterValues = ""
                    AfterValues += "PLANID=" + Convert.ToString(dr("planid"))
                    AfterValues += ",RID=" + Convert.ToString(dr("rid"))
                    AfterValues += ",ORGID=" + Convert.ToString(dr("OrgID"))
                    AfterValues += ",RELSHIP=" + Convert.ToString(dr("relship"))
                    AfterValues += ",ORGLEVEL=" + Convert.ToString(dr("OrgLevel"))
                    AfterValues += ",DISTID=" + Convert.ToString(dr("DistID"))
                    AfterValues += ",MODIFYACCT=" + Convert.ToString(dr("ModifyAcct"))
                    AfterValues += ",MODIFYDATE=" + Convert.ToDateTime(dr("ModifyDate")).ToString("yyyy-MM-dd HH:mm:ss.fff")
                    Dim Conditions As String = ""
                    Conditions += ("RSID=" + Convert.ToString(dr("rsid")))
                    '==========
                    t_iSql &= " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
                    t_iSql &= " VALUES(@SessionID, @TransTime, '/TC/01/TC_01_002_add.aspx', @UserID, 'Update', 'AUTH_RELSHIP', @Conditions, @BeforeValues, @AfterValues) "
                    myParam.Clear()
                    myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                    myParam.Add("SessionID", sm.SessionID.ToString)
                    myParam.Add("UserID", sm.UserInfo.UserID)
                    myParam.Add("Conditions", Conditions)
                    myParam.Add("BeforeValues", BeforeValues)
                    myParam.Add("AfterValues", AfterValues)
                    DbAccess.ExecuteNonQuery(t_iSql, objTrans, myParam)
                End If
        End Select
    End Sub

    ''' <summary>
    ''' 記錄 auth_apply 交易log 
    ''' </summary>
    ''' <param name="objTrans"></param>
    ''' <param name="dr"></param>
    Sub SaveAuthApplyLog(ByVal objTrans As SqlTransaction, ByVal dr As DataRow)
        'Dim s_TransType As String=TIMS.cst_TRANS_LOG_Insert 'insert:cst_TRANS_LOG_Insert/update:cst_TRANS_LOG_Update
        'Dim s_TargetTable As String="AUTH_APPLY"
        'Dim s_FuncPath As String="/TC/01/TC_01_002_add"
        'Const cst_fWHERE As String="OrgID=-1 and ComIDNO={0}"
        'Dim s_WHERE As String="" 'insert省略/update必要'String.Format(cst_fWHERE, pkVALUE)
        'Dim htPP As New Hashtable
        Dim BeforeValues As String = ""
        Dim AfterValues As String = ""
        Dim t_iSql As String = ""
        Dim myParam As Hashtable = New Hashtable
        Dim t_BeforeSql As String = " SELECT * FROM AUTH_APPLY WHERE OrgID=-1 and ComIDNO=@ComIDNO "
        Dim t_BeforeTB As DataTable = Nothing
        Dim t_BeforeRow As DataRow = Nothing
        myParam.Add("ComIDNO", Convert.ToString(dr("ComIDNO")))
        t_BeforeTB = DbAccess.GetDataTable(t_BeforeSql, objconn, myParam)
        If t_BeforeTB IsNot Nothing AndAlso t_BeforeTB.Rows.Count > 0 Then
            '新增作業
            'htPP.Clear()
            'htPP.Add("TransType", s_TransType)
            'htPP.Add("TargetTable", s_TargetTable)
            'htPP.Add("FuncPath", s_FuncPath)
            ''htPP.Add("s_WHERE", s_WHERE)
            'TIMS.SaveTRANSLOG(sm, objconn, objTrans, dr, htPP)
            t_BeforeRow = t_BeforeTB.Rows(0)
            '異動前資料
            BeforeValues = ""
            BeforeValues &= String.Concat("ORGID=", t_BeforeRow("OrgID"))
            'BeforeValues += ",MODIFYACCT=" + Convert.ToString(t_BeforeRow("ModifyAcct"))
            'BeforeValues += ",MODIFYDATE=" + Convert.ToDateTime(t_BeforeRow("ModifyDate")).ToString("yyyy-MM-dd HH:mm:ss.fff")

            '異動後資料
            AfterValues = ""
            AfterValues &= String.Concat("ORGID=", dr("OrgID"))
            Dim Conditions As String = ""
            Conditions &= String.Concat("OrgID=-1 and ComIDNO=", dr("ComIDNO"))
            '==========
            t_iSql = " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
            t_iSql &= " VALUES(@SessionID, @TransTime, '/TC/01/TC_01_002_add.aspx', @UserID, 'Update', 'AUTH_APPLY', @Conditions, @BeforeValues, @AfterValues) "

            myParam.Clear()
            myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
            myParam.Add("SessionID", sm.SessionID.ToString)
            myParam.Add("UserID", sm.UserInfo.UserID)
            myParam.Add("Conditions", Conditions)
            myParam.Add("BeforeValues", BeforeValues)
            myParam.Add("AfterValues", AfterValues)
            DbAccess.ExecuteNonQuery(t_iSql, objTrans, myParam)
        End If
    End Sub

    ''' <summary>
    ''' 記錄 org_apply 交易log
    ''' </summary>
    ''' <param name="objTrans"></param>
    ''' <param name="dr"></param>
    Sub SaveOrgApplyLog(ByVal objTrans As SqlTransaction, ByVal dr As DataRow)
        Dim BeforeValues As String = ""
        Dim AfterValues As String = ""
        Dim t_iSql As String = ""
        Dim myParam As Hashtable = New Hashtable
        Dim t_BeforeSql As String = " SELECT * FROM ORG_APPLY WHERE RESULT IS NULL AND COMIDNO=@ComIDNO "
        Dim t_BeforeTB As DataTable = Nothing
        Dim t_BeforeRow As DataRow = Nothing
        myParam.Add("COMIDNO", Convert.ToString(dr("ComIDNO")))
        t_BeforeTB = DbAccess.GetDataTable(t_BeforeSql, objconn, myParam)
        If t_BeforeTB IsNot Nothing AndAlso t_BeforeTB.Rows.Count > 0 Then
            t_BeforeRow = t_BeforeTB.Rows(0)
            '異動前資料
            BeforeValues = ""
            BeforeValues &= String.Concat("RESULT=", t_BeforeRow("Result"))
            'BeforeValues += ",MODIFYACCT=" + Convert.ToString(t_BeforeRow("ModifyAcct"))
            'BeforeValues += ",MODIFYDATE=" + Convert.ToDateTime(t_BeforeRow("ModifyDate")).ToString("yyyy-MM-dd HH:mm:ss.fff")
            '異動後資料
            AfterValues = ""
            AfterValues &= String.Concat("RESULT=", dr("Result"))

            Dim Conditions As String = ""
            Conditions &= String.Concat("ComIDNO=", dr("ComIDNO"))
            '==========
            t_iSql = " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
            t_iSql &= " VALUES(@SessionID, @TransTime, '/TC/01/TC_01_002_add.aspx', @UserID, 'Update', 'ORG_APPLY', @Conditions, @BeforeValues, @AfterValues) "
            myParam.Clear()
            myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
            myParam.Add("SessionID", sm.SessionID.ToString)
            myParam.Add("UserID", sm.UserInfo.UserID)
            myParam.Add("Conditions", Conditions)
            myParam.Add("BeforeValues", BeforeValues)
            myParam.Add("AfterValues", AfterValues)
            DbAccess.ExecuteNonQuery(t_iSql, objTrans, myParam)
        End If
    End Sub

    '階層選擇
    'Protected Sub level_list_SelectedIndexChanged(sender As Object, e As EventArgs) Handles level_list.SelectedIndexChanged
    '    Call sUtl_levellistSel()
    'End Sub

#End Region

End Class