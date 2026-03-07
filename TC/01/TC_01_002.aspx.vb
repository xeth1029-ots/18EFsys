Imports System.Web.Mvc

Partial Class TC_01_002
    Inherits AuthBasePage
    'Protected WithEvents PageControler1 As PageControler

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    'ProcessType : 修改(年度對應功能 ) Update/ 共用 Share /審核 InsertChk /新增 Insert
    Dim ProcessType As String
    'Dim objconn As SqlConnection
    'Dim objreader As SqlDataReader
    'Dim FunDr As DataRow
    Dim dtKey_Years As DataTable

    Const Cst_ComIDNO As Integer = 0
    Const Cst_Years As Integer = 1
    Const Cst_GradeDate As Integer = 2
    Const Cst_FreeComments As Integer = 3
    Const Cst_Point01A As Integer = 4
    Const Cst_Point01B As Integer = 5
    Const Cst_Point02A As Integer = 6
    Const Cst_Point02B As Integer = 7
    Const Cst_Point03A As Integer = 8
    Const Cst_Point03B As Integer = 9
    Const Cst_Point04A As Integer = 10
    Const Cst_Point04B As Integer = 11
    Const Cst_ClassCNames As Integer = 12
    Const cst_filedNum As Integer = 13

    Dim G_OrgID As String = ""
    Dim G_Years As String = ""
    Dim G_GradeDate As String = ""
    Dim G_FreeComments As String = ""
    Dim G_Point01A As String = ""
    Dim G_Point01B As String = ""
    Dim G_Point02A As String = ""
    Dim G_Point02B As String = ""
    Dim G_Point03A As String = ""
    Dim G_Point03B As String = ""
    Dim G_Point04A As String = ""
    Dim G_Point04B As String = ""
    Dim G_ClassCNames As String = ""

    Dim str_superuser1 As String = "snoopy" '(預設)(吃管理者權限)
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '2018 動態連動產生下拉-訓練機構屬性
        If Request("OP") = "Ajax" And Request("TYPEID1") <> "" Then
            ' Ajax 載入計畫清單
            Call ResponseTypeID2(Request("TYPEID1"))
            Return
        End If

        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        TIMS.OpenDbConn(objconn)

        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
            str_superuser1 = CStr(sm.UserInfo.UserID)
        End If

        'tr_IsApply.Visible = False '暫無此機制
        tr_IsApply.Style.Add("display", "none") '暫無此機制
        orglevelTR.Visible = False
        If flgROLEIDx0xLIDx0 Then orglevelTR.Visible = True

        ProcessType = TIMS.ClearSQM(Request("ProcessType"))
        bt_search.Attributes("onclick") = "return Search();"
        'Call check_bt_add() '限定計畫啟用。 '2018-09-25 mark： 不用再呼叫此function了

        '分頁設定 Start
        PageControler1.PageDataGrid = DG_Org
        '分頁設定 End

#Region "(No Use)"

        '2018 todo:按鈕權限檢核尚未做，先開啟權限檢測功能
        'check_add.Value = "1"
        'Me.bt_add.Enabled = True
        'check_del.Value = "1"
        'check_mod.Value = "1"

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        If FunDr("Adds") = "1" Then
        '            check_add.Value = "1"
        '            Me.bt_add.Enabled = True
        '        Else
        '            check_add.Value = "0"
        '            Me.bt_add.Enabled = False
        '            TIMS.Tooltip(Me.bt_add, "登入者無權新增")
        '        End If
        '        If FunDr("Sech") = "1" Then
        '            bt_search.Enabled = True
        '        Else
        '            bt_search.Enabled = False
        '            TIMS.Tooltip(bt_search, "登入者無查詢權限")
        '        End If
        '        If FunDr("Del") = "1" Then
        '            check_del.Value = "1"
        '        Else
        '            check_del.Value = "0"
        '        End If
        '        If FunDr("Mod") = "1" Then
        '            check_mod.Value = "1"
        '        Else
        '            check_mod.Value = "0"
        '        End If
        '    End If
        'End If

#End Region

        If Not Me.IsPostBack Then
            Call Create1()

            '取得查詢條件
            If Session("_Search") IsNot Nothing Then
                Dim MyValue As String = ""
                Dim str1 As String = Convert.ToString(Session("_Search"))
                TB_OrgName.Text = TIMS.GetMyValue(str1, "TB_OrgName")
                TB_ComIDNO.Text = TIMS.GetMyValue(str1, "TB_ComIDNO")
                TBCity.Text = TIMS.GetMyValue(str1, "TBCity")
                city_code.Value = TIMS.GetMyValue(str1, "city_code")
                zip_code.Value = TIMS.GetMyValue(str1, "zip_code")
                Common.SetListItem(DistID, TIMS.GetMyValue(str1, "DistID"))
                Common.SetListItem(OrgKindList, TIMS.GetMyValue(str1, "OrgKindList"))
                Common.SetListItem(Yearlist, TIMS.GetMyValue(str1, "Years"))
                Common.SetListItem(drpPlan, TIMS.GetMyValue(str1, "drpPlan"))
                Common.SetListItem(IsApply, TIMS.GetMyValue(str1, "IsApply"))
                'PageControler1.PageIndex = TIMS.GetMyValue(str1, "PageIndex")
                'MyValue = TIMS.GetMyValue(str1, "PageIndex")
                MyValue = TIMS.GetMyValue(str1, "Button1")
                If MyValue = "True" Then
                    MyValue = TIMS.GetMyValue(str1, "PageIndex")
                    'Me.ViewState("PageIndex") = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
                    If IsNumeric(MyValue) Then PageControler1.PageIndex = Val(MyValue)
                    'bt_search_Click(sender, e)
                    Call SetViewStateVal1()
                    Call Search1()
                    'If IsNumeric(Me.ViewState("PageIndex")) Then
                    '    PageControler1.PageIndex = Me.ViewState("PageIndex")
                    '    PageControler1.CreateData()
                    'End If
                End If
                Session("_Search") = Nothing
            End If

        End If

        hidLID.Value = sm.UserInfo.LID

        '只顯示正式資料
        IsApply.Enabled = False
        Select Case sm.UserInfo.LID
            Case "0"
                '署。
                IsApply.Enabled = True
                'Me.IsApply.Attributes("onclick") = "IsApply_display(this);"
            Case "1" '分署 
                '顯示 正式與審核 資料
                IsApply.Enabled = True
                'Me.IsApply.Attributes("onclick") = "IsApply_display(this);"
            Case Else '其他。
        End Select
        If Not IsApply.Enabled Then
            IsApply.Visible = False
            TIMS.Tooltip(IsApply, "登入者無權選擇")
        End If
        drpPlan.Attributes("onchange") = "chgPlan();"

        'Page.RegisterStartupScript("window_onload", "<script language=""javascript"">IsApply_display(document.getElementById('IsApply'));chgPlan();</script>")
        Page.RegisterStartupScript("window_onload", "<script language=""javascript"">chgPlan();</script>") '2018 改版：先 mark 資料狀態（org_apply無訓練機構別資料欄位）
    End Sub

    Sub Create1()
        'TIMS.Get_Years()
        'Years.Items.Clear()
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " SELECT DISTINCT years, dbo.FN_CYEAR2b(years) roc_year" & vbCrLf
        'sql &= " FROM dbo.ID_Plan WITH(NOLOCK)" & vbCrLf
        'sql &= " WHERE 1=1 AND years!=' ' AND years <= CONVERT(VARCHAR, DATEPART(year,DATEADD(YEAR,10,GETDATE())))" & vbCrLf
        'sql &= " ORDER BY years" & vbCrLf
        'Dim dt As DataTable
        'dt = DbAccess.GetDataTable(sql, objconn)
        'If dt.Rows.Count > 0 Then
        '    With Years
        '        .DataSource = dt
        '        .DataTextField = If(flag_ROC, "roc_year", "Years")
        '        .DataValueField = "Years"
        '        .DataBind()
        '        If TypeOf Years Is DropDownList Then .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        '    End With
        '    Common.SetListItem(Years, "")
        'End If
        Yearlist = TIMS.Get_Years(Yearlist, objconn)
        Yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(Yearlist, sm.UserInfo.Years)

        drpPlan = TIMS.Get_TPlan(drpPlan)
        Common.SetListItem(drpPlan, sm.UserInfo.TPlanID)

        If TIMS.Cst_DistID78.IndexOf(sm.UserInfo.DistID) > -1 Then
            DistID = TIMS.Get_DistID78(Me, DistID, objconn)
        Else
            DistID = TIMS.Get_DistID(DistID)
        End If
        OrgKindList = TIMS.Get_OrgType(OrgKindList, objconn)

    End Sub

    ''' <summary>
    ''' 設定查詢值
    ''' </summary>
    Sub SetViewStateVal1()
        ViewState("TB_OrgName") = ""
        ViewState("TB_ComIDNO") = ""
        ViewState("TB_ComIDNO_lk") = ""
        ViewState("TypeID1") = ""
        ViewState("TypeID2") = ""

        TB_OrgName.Text = TIMS.ClearSQM(TB_OrgName.Text)
        TB_ComIDNO.Text = TIMS.ClearSQM(TB_ComIDNO.Text)
        ViewState("TB_OrgName") = TB_OrgName.Text
        ViewState("TB_ComIDNO") = TB_ComIDNO.Text
        If Len(ViewState("TB_ComIDNO")) >= 8 Then
            '超過8碼使用like 查詢 清空 ViewState("TB_ComIDNO")
            ViewState("TB_ComIDNO_lk") = ViewState("TB_ComIDNO")
            ViewState("TB_ComIDNO") = ""
        End If
        '2018 add (機構屬性)機構別
        ViewState("TypeID1") = TIMS.GetListValue(rblPlanPoint) '.SelectedValue
        ViewState("TypeID2") = TIMS.GetListValue(dl_typeid2) '.SelectedValue
    End Sub

    ''' <summary>
    ''' 資料狀態為「正式」時的查詢sql
    ''' </summary>
    ''' <param name="blnIsSuperUser"></param>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function SearchSql1(ByVal blnIsSuperUser As Boolean, ByRef parms As Hashtable) As String
        Dim sqlstr As String = ""
        Dim v_drpPlan As String = TIMS.GetListValue(drpPlan)
        Dim blChkPlan28 As Boolean = (TIMS.Cst_TPlanID28AppPlan.IndexOf(v_drpPlan) > -1)  '是否為查詢產投or充飛計畫

        '整理變數
        Dim v_zip_code As String = ""
        Dim v_city_code As String = ""
        Dim v_TB_OrgName As String = ""
        Dim v_TB_ComIDNO As String = ""
        Dim v_TB_ComIDNO_lk As String = ""
        Dim v_OrgKindList_SelV As String = ""
        Dim v_DistID_SelV As String = ""
        Dim v_rblOrgLevel_SelV As String = ""
        Dim v_TypeID1 As String = ""
        Dim v_TypeID2 As String = ""
        Dim v_Years_SelV As String = ""
        Dim v_drpPlan_SelV As String = ""

        If zip_code.Value <> "" Then v_zip_code = TIMS.ClearSQM(zip_code.Value)
        If city_code.Value <> "" Then v_city_code = TIMS.ClearSQM(city_code.Value)
        If ViewState("TB_OrgName") <> "" Then v_TB_OrgName = TIMS.ClearSQM(ViewState("TB_OrgName"))
        If ViewState("TB_ComIDNO") <> "" Then v_TB_ComIDNO = TIMS.ClearSQM(ViewState("TB_ComIDNO"))
        If ViewState("TB_ComIDNO_lk") <> "" Then v_TB_ComIDNO_lk = TIMS.ClearSQM(ViewState("TB_ComIDNO_lk"))

        v_OrgKindList_SelV = TIMS.GetListValue(OrgKindList) '.SelectedValue)
        v_DistID_SelV = TIMS.GetListValue(DistID) '.SelectedValue)
        v_rblOrgLevel_SelV = TIMS.GetListValue(rblOrgLevel) '.SelectedValue)
        v_TypeID1 = TIMS.ClearSQM(ViewState("TypeID1"))
        v_TypeID2 = TIMS.ClearSQM(ViewState("TypeID2"))
        v_Years_SelV = TIMS.GetListValue(Yearlist) '.SelectedValue)
        v_drpPlan_SelV = TIMS.GetListValue(drpPlan) '.SelectedValue)

        sqlstr = ""
        sqlstr &= " SELECT a.orgid ,b.RSID ,b.relship ,a.OrgName" & vbCrLf
        sqlstr &= " ,c.name DISTNAME,a.ComIDNO" & vbCrLf
        sqlstr &= " ,a.OrgKind,k1.NAME OrgKindNAME" & vbCrLf
        ',f.Address 
        sqlstr &= " ,dbo.FN_ADDR2(f.ZIPCODE,f.ZIPCODE6W,'',dbo.FN_ADDR1(f.Address,dbo.FN_GET_ZIPNAME(f.ZIPCODE))) Address" & vbCrLf

        sqlstr &= " ,b.distid ,b.PlanID ,b.RID" & vbCrLf
        sqlstr &= " ,d.PlanName,f.ActNo,f.ContactName,f.ContactEmail,f.CONTACTCELLPHONE,f.MasterName,f.modifyAcct" & vbCrLf
        sqlstr &= " ,ISNULL(r3.orgname2,c.name) orgname2" & vbCrLf
        sqlstr &= " FROM ORG_ORGINFO a" & vbCrLf
        sqlstr &= " JOIN AUTH_RELSHIP b ON a.orgid = b.orgid" & vbCrLf
        sqlstr &= " JOIN ID_DISTRICT c ON b.distid = c.distid" & vbCrLf
        If blnIsSuperUser Then '系統管理者 (ID_PLAN)
            sqlstr &= " LEFT JOIN VIEW_LOGINPLAN d ON d.PlanID = b.PlanID" & vbCrLf
        Else
            sqlstr &= " JOIN VIEW_LOGINPLAN d ON d.PlanID = b.PlanID" & vbCrLf
        End If
        sqlstr &= " JOIN ORG_ORGPLANINFO f ON f.RSID = b.RSID" & vbCrLf
        sqlstr &= " LEFT JOIN VIEW_RELSHIP23 r3 ON r3.RSID3 = b.RSID" & vbCrLf 'view_relship23
        sqlstr &= " LEFT JOIN KEY_ORGTYPE k1 ON A.ORGKIND = k1.ORGTYPEID" & vbCrLf
        '2018 add:merge TC_01_017 機構屬性設定
        sqlstr &= " LEFT JOIN KEY_ORGTYPE1 G ON A.ORGKIND1 = G.ORGTYPEID1" & vbCrLf
        sqlstr &= " WHERE b.OrgLevel>=@OrgLevel" & vbCrLf

        If v_zip_code <> "" Then
            sqlstr &= " AND f.ZipCode = @ZipCode" & vbCrLf
        End If
        If v_zip_code = "" AndAlso v_city_code <> "" Then
            sqlstr &= " AND f.ZipCode IN (SELECT zipcode FROM ID_Zip WHERE ctid=@ctid)" & vbCrLf
        End If
        If v_TB_OrgName <> "" Then sqlstr &= " AND a.OrgName LIKE '%' + @OrgName + '%'" & vbCrLf
        If v_TB_ComIDNO <> "" Then sqlstr &= " AND a.ComIDNO = @ComIDNO" & vbCrLf
        '超過8碼使用like 查詢 清空 ViewState("TB_ComIDNO")
        If v_TB_ComIDNO_lk <> "" Then sqlstr &= " AND a.ComIDNO LIKE '%' + @ComIDNOlk + '%'" & vbCrLf
        If v_OrgKindList_SelV <> "" Then sqlstr &= " AND a.OrgKind = @OrgKind" & vbCrLf
        If v_DistID_SelV <> "" Then sqlstr &= " AND b.DistID = @DistID" & vbCrLf
        Select Case v_rblOrgLevel_SelV
            Case "2", "3"
                sqlstr &= " AND b.OrgLevel = @OrgLevel23" & vbCrLf
        End Select
        '2018 add 機構屬性-計畫別( 0 不區分, 1 產投計畫, 2 提升勞工學習自主計畫)
        If blChkPlan28 AndAlso v_TypeID1 <> "" AndAlso v_TypeID1 <> "0" Then sqlstr &= " AND g.TYPEID1 = @TYPEID1" & vbCrLf
        '2018 add 機構屬性-機構別(2018 add)
        If blChkPlan28 AndAlso v_TypeID2 <> "" Then sqlstr &= " AND g.TYPEID2 = @TYPEID2" & vbCrLf
        If v_Years_SelV <> "" Then sqlstr &= " AND d.Years = @Years" & vbCrLf
        If v_drpPlan_SelV <> "" Then sqlstr &= " AND d.TPlanID = @TPlanID" & vbCrLf

        ' sql 參數設定
        parms.Add("OrgLevel", sm.UserInfo.OrgLevel)
        If v_zip_code <> "" Then parms.Add("ZipCode", v_zip_code)
        If v_zip_code = "" AndAlso v_city_code <> "" Then parms.Add("ctid", v_city_code)
        If v_TB_OrgName <> "" Then parms.Add("OrgName", v_TB_OrgName)
        If v_TB_ComIDNO <> "" Then parms.Add("ComIDNO", v_TB_ComIDNO)
        '超過8碼使用like 查詢 清空 ViewState("TB_ComIDNO")
        If v_TB_ComIDNO_lk <> "" Then parms.Add("ComIDNOlk", v_TB_ComIDNO_lk)
        If v_OrgKindList_SelV <> "" Then parms.Add("OrgKind", v_OrgKindList_SelV)
        If v_DistID_SelV <> "" Then parms.Add("DistID", v_DistID_SelV)
        Select Case v_rblOrgLevel_SelV
            Case "2", "3"
                parms.Add("OrgLevel23", v_rblOrgLevel_SelV)
        End Select
        '機構屬性-計畫別 '2018 add 機構屬性-計畫別( 0 不區分, 1 產投計畫, 2 提升勞工學習自主計畫)
        If blChkPlan28 AndAlso v_TypeID1 <> "" AndAlso v_TypeID1 <> "0" Then parms.Add("TYPEID1", v_TypeID1)
        '機構屬性-機構別 '2018 add 機構屬性-機構別(2018 add)
        If blChkPlan28 AndAlso v_TypeID2 <> "" Then parms.Add("TYPEID2", v_TypeID2)
        If v_Years_SelV <> "" Then parms.Add("Years", v_Years_SelV)
        If v_drpPlan_SelV <> "" Then parms.Add("TPlanID", v_drpPlan_SelV)

        Return sqlstr
    End Function

    ''' <summary>
    ''' 資料狀態為「審核中」時的查詢sql
    ''' </summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function SearchSql2(ByRef parms As Hashtable) As String
        '整理變數
        Dim v_zip_code As String = ""
        Dim v_city_code As String = ""
        Dim v_TB_OrgName As String = ""
        Dim v_TB_ComIDNO As String = ""
        Dim v_OrgKindList_SelV As String = ""
        Dim v_Years_SelV As String = ""
        Dim v_drpPlan_SelV As String = ""

        If zip_code.Value <> "" Then v_zip_code = TIMS.ClearSQM(zip_code.Value)
        If city_code.Value <> "" Then v_city_code = TIMS.ClearSQM(city_code.Value)
        If ViewState("TB_OrgName") <> "" Then v_TB_OrgName = TIMS.ClearSQM(ViewState("TB_OrgName"))
        If ViewState("TB_ComIDNO") <> "" Then v_TB_ComIDNO = TIMS.ClearSQM(ViewState("TB_ComIDNO"))
        v_OrgKindList_SelV = TIMS.GetListValue(OrgKindList)
        v_Years_SelV = TIMS.GetListValue(Yearlist)
        v_drpPlan_SelV = TIMS.GetListValue(drpPlan)

        Dim sqlstr As String = ""
        sqlstr &= " SELECT b.OrgName ,a.ComIDNO" & vbCrLf
        sqlstr &= " ,b.planID ,b.ActNo ,b.ContactName ,b.ContactEmail,b.MasterName" & vbCrLf
        ',b.Address 
        sqlstr &= " ,dbo.FN_ADDR2(b.ZIPCODE,'','',dbo.FN_ADDR1(b.Address,dbo.FN_GET_ZIPNAME(b.ZIPCODE))) Address" & vbCrLf
        sqlstr &= " ,d.distid ,0 AS orgid ,0 AS RSID" & vbCrLf
        sqlstr &= " ,c.name ,d.PlanName ,NULL orgname2" & vbCrLf
        sqlstr &= " FROM ORG_APPLY b" & vbCrLf
        sqlstr &= " JOIN VIEW_LOGINPLAN d ON d.PlanID = b.PlanID" & vbCrLf
        sqlstr &= " JOIN AUTH_APPLY a ON a.ComIDNO = b.ComIDNO AND a.OrgID = '-1'" & vbCrLf
        sqlstr &= " JOIN ID_DISTRICT c ON c.distid = d.distid" & vbCrLf
        sqlstr &= " WHERE d.distid = @distid" & vbCrLf
        If v_zip_code <> "" Then
            sqlstr &= " AND b.ZipCode = @ZipCode" & vbCrLf
        End If
        If v_zip_code = "" AndAlso v_city_code <> "" Then
            sqlstr &= " AND b.ZipCode IN (SELECT zipcode FROM ID_Zip WHERE ctid = @ctid)" & vbCrLf
        End If
        If v_TB_OrgName <> "" Then sqlstr &= " AND b.OrgName LIKE '%' + @OrgName + '%'" & vbCrLf
        If v_TB_ComIDNO <> "" Then sqlstr &= " AND a.ComIDNO = @ComIDNO" & vbCrLf
        If v_OrgKindList_SelV <> "" Then sqlstr &= " AND b.OrgKind = @OrgKind" & vbCrLf
        If v_Years_SelV <> "" Then sqlstr &= " AND d.Years = @Years" & vbCrLf
        If v_drpPlan_SelV <> "" Then sqlstr &= " AND d.TPlanID = @TPlanID" & vbCrLf

        ' sql 參數設定
        parms.Add("distid", sm.UserInfo.DistID)
        If v_zip_code <> "" Then parms.Add("ZipCode", v_zip_code)
        If v_zip_code = "" AndAlso v_city_code <> "" Then parms.Add("ctid", v_city_code)
        If v_TB_OrgName <> "" Then parms.Add("OrgName", v_TB_OrgName)
        If v_TB_ComIDNO <> "" Then parms.Add("ComIDNO", v_TB_ComIDNO)
        If v_OrgKindList_SelV <> "" Then parms.Add("OrgKind", v_OrgKindList_SelV)
        If v_Years_SelV <> "" Then parms.Add("Years", v_Years_SelV)
        If v_drpPlan_SelV <> "" Then parms.Add("TPlanID", v_drpPlan_SelV)

        Return sqlstr
    End Function

    Function Search1_dt() As DataTable
        Dim blnIsSuperUser As Boolean = False
        If TIMS.IsSuperUser(Me, 1) Then blnIsSuperUser = True

        Dim sqlstr As String = ""
        Dim parms As Hashtable = New Hashtable()

        Dim v_IsApply As String = TIMS.GetListValue(IsApply)
        Select Case v_IsApply 'Me.IsApply.SelectedValue
            Case "Y"
                sqlstr = SearchSql1(blnIsSuperUser, parms)
            Case "N"
                sqlstr = SearchSql2(parms)
        End Select
        Session("exp_tc_01_002_sql") = sqlstr
        Session("exp_tc_01_002_parms") = parms
        'Dim sqlAdapter As SqlDataAdapter
        Dim dtOrgInfo As DataTable
        dtOrgInfo = DbAccess.GetDataTable(sqlstr, objconn, parms)
        Return dtOrgInfo
    End Function

    ''' <summary>
    ''' 查詢
    ''' </summary>
    Sub Search1()
        'Dim sqlAdapter As SqlDataAdapter
        Dim dtOrgInfo As DataTable = Search1_dt()

        Panel.Visible = False
        DG_Org.Visible = False
        msg.Text = "查無資料!!"
        bt_EXPORT.Visible = False

        If dtOrgInfo.Rows.Count > 0 Then
            bt_EXPORT.Visible = True
            Panel.Visible = True
            msg.Text = ""
            DG_Org.Visible = True
            PageControler1.PageDataTable = dtOrgInfo
            PageControler1.PrimaryKey = "RSID"
            PageControler1.Sort = "OrgID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    ''' <summary>
    ''' 查詢鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call SetViewStateVal1()
        Call Search1()
    End Sub

    Private Sub Bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_add.Click
        KeepSearch()
        'Response.Redirect("TC_01_002_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "")
        '20100208 按新增時代查詢之 機構名稱 & 統一編號 新增 Insert
        TIMS.Utl_Redirect1(Me, "TC_01_002_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "&OrgName=" & TB_OrgName.Text & "&ComIDNO=" & TB_ComIDNO.Text & "")
    End Sub

    ''' <summary>
    ''' 檢核有無帳號 有回傳1
    ''' </summary>
    ''' <param name="drv"></param>
    ''' <returns></returns>
    Function Check_account(ByRef drv As DataRowView) As String
        Dim rst As String = ""
        TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        Dim sql As String = ""
        sql &= " SELECT b.AccFunRecord" & vbCrLf
        sql &= " FROM AUTH_RELSHIP a" & vbCrLf
        sql &= " JOIN VIEW_ACCFUNCOUNT b ON a.RID=b.RID" & vbCrLf
        sql &= " WHERE a.ORGID=@ORGID AND a.RID=@RID" & vbCrLf
        Using sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Add("ORGID", SqlDbType.BigInt).Value = Convert.ToInt32(drv("orgid"))
                .Parameters.Add("RID", SqlDbType.VarChar).Value = Convert.ToString(drv("RID"))
                dt1.Load(.ExecuteReader())
            End With
        End Using
        If TIMS.dtNODATA(dt1) Then Return rst
        If Val(dt1.Rows(0)("AccFunRecord")) > 0 Then rst = "1" '有帳號
        Return rst 'account = ""'If DbAccess.ExecuteScalar(sql, objconn) > 0 Then account = "1" '有帳號
    End Function

    ''' <summary>
    ''' 檢核有無訓練計畫，有回傳1
    ''' </summary>
    ''' <param name="drv"></param>
    ''' <returns></returns>
    Function Check_ppinfo(ByRef drv As DataRowView) As String
        Dim rst As String = ""
        TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        Dim sql As String = ""
        sql &= " SELECT 'x' FROM AUTH_RELSHIP a" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO b ON a.rid = b.rid AND A.PLANID = B.PLANID" & vbCrLf
        sql &= " WHERE a.ORGID=@ORGID AND a.RID=@RID" & vbCrLf
        Using sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Add("ORGID", SqlDbType.BigInt).Value = Convert.ToInt32(drv("orgid"))
                .Parameters.Add("RID", SqlDbType.VarChar).Value = Convert.ToString(drv("RID"))
                dt1.Load(.ExecuteReader())
            End With
        End Using
        If TIMS.dtNODATA(dt1) Then Return rst
        rst = "1"
        Return rst
    End Function

    ''' <summary>
    ''' 檢核有無開班計畫，有回傳1
    ''' </summary>
    ''' <returns></returns>
    Function Check_ccinfo(ByRef drv As DataRowView) As String
        Dim rst As String = ""
        TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        Dim sql As String = ""
        sql &= " SELECT a.RID FROM AUTH_RELSHIP a JOIN CLASS_CLASSINFO b ON a.RID = b.RID AND A.PLANID = B.PLANID"
        sql &= " WHERE a.ORGID=@ORGID AND a.RID=@RID" & vbCrLf
        Using sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Add("ORGID", SqlDbType.BigInt).Value = Convert.ToInt32(drv("orgid"))
                .Parameters.Add("RID", SqlDbType.VarChar).Value = Convert.ToString(drv("RID"))
                dt1.Load(.ExecuteReader())
            End With
        End Using
        If TIMS.dtNODATA(dt1) Then Return rst
        rst = "1"
        Return rst
    End Function

    ''' <summary>
    ''' 檢核有無子單位，有回傳true
    ''' </summary>
    ''' <param name="drv"></param>
    ''' <returns></returns>
    Function Check_suborg(ByRef drv As DataRowView) As Boolean
        Dim rst As Boolean = False
        TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        Dim sqlstr_A As String = ""
        sqlstr_A &= " SELECT a.orgid FROM ORG_ORGINFO a JOIN Auth_Relship b ON a.ORGID=b.ORGID"
        sqlstr_A &= " WHERE b.RELSHIP LIKE @RELSHIP+'%'"
        Using sCmd As New SqlCommand(sqlstr_A, objconn)
            With sCmd
                .Parameters.Add("RELSHIP", SqlDbType.VarChar).Value = Convert.ToString(drv("relship"))
                dt1.Load(.ExecuteReader())
            End With
        End Using
        'If dt1.Rows.Count = 0 Then Return rst
        If dt1.Rows.Count > 1 Then rst = True '有子單位
        Return rst
    End Function

    Private Sub DG_Org_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Org.ItemCommand
        'Dim lbtShare As LinkButton = e.Item.Cells(6).FindControl("lbtShare") '共用
        'Dim but_edit As LinkButton = e.Item.Cells(6).FindControl("lbtEdit") '修改
        'Dim but_del As LinkButton = e.Item.Cells(6).FindControl("lbtDel") '刪除
        'Dim but_chk As LinkButton = e.Item.Cells(6).FindControl("lbtChk") '審核
        'Dim but_year As LinkButton = e.Item.Cells(6).FindControl("lbtYear") '年度對應功能

        Select Case e.CommandName
            Case "edit" '修改 Update
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "TC_01_002_add.aspx?ProcessType=Update&" & e.CommandArgument & "")
            Case "share" '共用 Share
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "TC_01_002_add.aspx?ProcessType=Share&" & e.CommandArgument & "")
            Case "del" '刪除
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "TC_01_002_del.aspx?" & e.CommandArgument & "")
            Case "chk" '審核確認 InsertChk
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "TC_01_002_add.aspx?ProcessType=InsertChk&" & e.CommandArgument & "")
            Case "year" '年度對應功能 Update
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "TC_01_002_year.aspx?ProcessType=Update&" & e.CommandArgument & "")
        End Select
    End Sub

    Private Sub DG_Org_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Org.ItemDataBound
        Dim strScript1 As String = ""
        Select Case e.Item.ItemType
            Case ListItemType.Header, ListItemType.Footer
            Case Else
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DG_Org.PageSize * DG_Org.CurrentPageIndex
                Dim lbtShare As LinkButton = e.Item.Cells(6).FindControl("lbtShare") '共用
                Dim but_edit As LinkButton = e.Item.Cells(6).FindControl("lbtEdit") '修改
                Dim but_del As LinkButton = e.Item.Cells(6).FindControl("lbtDel") '刪除
                Dim lbtChk As LinkButton = e.Item.Cells(6).FindControl("lbtChk") '審核
                Dim but_year As LinkButton = e.Item.Cells(6).FindControl("lbtYear") '年度對應功能

                Dim plan_Name As String = SGetPlanName(drv("PlanID"), drv("RID"))
                Dim vTitle As String = ""
                vTitle = ""
                vTitle &= "[" & drv("RSID") & "] " & plan_Name & drv("orgname2") & vbCrLf
                vTitle &= "/業務ID:" & Convert.ToString(drv("RID"))
                TIMS.Tooltip(e.Item.Cells(1), vTitle)
                TIMS.Tooltip(e.Item.Cells(2), vTitle)
                TIMS.Tooltip(e.Item.Cells(3), vTitle)

                Dim sql As String = ""
                Dim v_IsApply As String = TIMS.GetListValue(IsApply) '資料狀態 Y:正式 / N:審核中
                Select Case v_IsApply'Me.IsApply.SelectedValue '目前已經在使用的 ORG_ORGINFO
                    Case "Y" '資料狀態 Y:正式 / N:審核中
                        lbtChk.Visible = False
                        but_edit.CommandArgument = "orgid=" & drv("orgid") & "&planid=" & drv("PlanID") & "&rid=" & drv("RID") & "&RSID=" & drv("RSID") & "&distid=" & drv("distid") & "&ID=" & Request("ID") & ""
                        but_year.CommandArgument = "orgid=" & drv("orgid") & "&planid=" & drv("PlanID") & "&rid=" & drv("RID") & "&RSID=" & drv("RSID") & "&distid=" & drv("distid") & "&ID=" & Request("ID") & ""
                        but_del.CommandArgument = "orgid=" & drv("orgid") & "&rid=" & drv("RID") & "&planid=" & drv("PlanID") & "&RSID=" & drv("RSID") & "&ID=" & Request("ID") & ""
                        lbtShare.CommandArgument = "orgid=" & drv("orgid") & "&distid=" & drv("distid") & "&planid=" & drv("PlanID") & "&rid=" & drv("RID") & "&RSID=" & drv("RSID") & "&ID=" & Request("ID") & ""
                        If Val(drv("PlanID")) = 0 Then
                            lbtShare.Enabled = False '無計畫代碼，不可共用
                        End If

                        'Dim is_parent, classid, account, plan_sql, plan_list As String
                        'Dim account_sql As String = "select AccFunRecord from Auth_Relship a   join VIEW_AccFunCount b on a.rid=b.rid where a.orgid='" & drv("orgid") & "' and a.rid='" & drv("RID") & "'"
                        'Dim sql As String = ""
                        '檢核有無帳號
                        Dim account As String = Check_account(drv)
                        '檢核有無訓練計畫
                        Dim plan_list As String = Check_ppinfo(drv)
                        '檢核有無開班計畫
                        Dim classid As String = Check_ccinfo(drv)
                        '檢核有無子單位
                        Dim flag_is_parent As Boolean = Check_suborg(drv)

                        Dim NotdelFlag1 As Boolean = False '不可刪除
                        If plan_list <> "" Then
                            NotdelFlag1 = True '不可刪除
                            but_del.Attributes("onclick") = "javascript:alert('此機構已有計畫資料，不可以刪除!!');return false;"
                        ElseIf classid <> "" Then
                            NotdelFlag1 = True '不可刪除
                            but_del.Attributes("onclick") = "javascript:alert('此機構已有開班資料，不可以刪除!!');return false;"
                        ElseIf flag_is_parent Then
                            NotdelFlag1 = True '不可刪除
                            but_del.Attributes("onclick") = "javascript:alert('此機構尚有下層單位,不可刪除!!');return false;"
                        ElseIf classid = "" And account = "" And plan_list = "" And (Not flag_is_parent) Then
                            but_del.Attributes("onclick") = "javascript:return confirm('此動作會刪除機構資料，是否確定刪除?');"
                        End If

                        If Not NotdelFlag1 Then
                            '沒有帳號資料
                            If account = "1" Then
                                '有帳號資料
                                but_del.Attributes("onclick") = "javascript:return confirm('此機構已有帳號計畫資料，是否確定刪除?');"
                                'but_del.Attributes("onclick") = "javascript:alert('此機構已有帳號資料，不可以刪除!!(業務ID:" & Convert.ToString(drv("RID")) & ")');return false;"
                            End If
                        End If
                        'but_del.Attributes.Add("onclick", "but_del(" & drv("orgid") & ",'" & account & "','" & classid & "','" & drv("RID") & "'," & drv("PlanID") & "," & is_parent & "," & Request("ID") & ");return false;")

                        If Convert.ToString(drv("DISTID")) <> sm.UserInfo.DistID Then '不同轄區
                            'lbtShare.Enabled = False
                            '10	新興科技人才培訓
                            '11 資訊軟體人才培訓
                            'If sm.UserInfo.TPlanID = "10" Or sm.UserInfo.TPlanID = "11" Then
                            '    'mark by nick 20060407  'If sm.UserInfo.DistID <> "002" Then lbtShare.Enabled = False 'Else
                            '    lbtShare.Enabled = True
                            'End If
                            but_edit.Enabled = False '不可修改
                            but_del.Enabled = False '不可刪除
                            TIMS.Tooltip(but_edit, "不同轄區")
                            TIMS.Tooltip(but_del, "不同轄區")
                        Else
                            'If sm.UserInfo.TPlanID = "10" Or sm.UserInfo.TPlanID = "11" Then lbtShare.Enabled = True
                            '同轄區
                            '39B8D0C19B534B3156E731DD70917DBE61084CC024A31970D365C6BF0E146BFD5CDA8AC85C6FA208E455D89753023C0065D2DCC54E08C3561E6B32
                            'F979D784B3017B14E3371075104A7D88DE16329AC41B52049A8C281B21C97727B60E9BAD791A9270A8569E
                            '2F504A690E44696512554C40278855A8685B475816D06A6D664B78BFEF1F15EED121579E170C8857E247E5BEB0
                            '4A42730B9FFB46BA9BBF2630BD05EDFE88CD38240660CD8AE28517C62CAECC8C1D7131800BC2E1815046A5A97EE38AFE2BD53A961EAF702C6E4154
                            'F06C01409162E613DE421DA26FD4779CA1A2DF424995C2462C9C3A2AD3A0C1037AFA8293D0663FDF8A7DA8
                            '1143C1BA8FE0FC6C02976901EE12E37062280D47ED937B568831254EBF09EDB8E3E4B28D3ADE6A80BB39A3CE78419F869083D8A7749412A5252489
                            'DF568BF123B62F324B1BFCE8DD47F497A4235FC9C53C9E37A1C016520AAFFFC57071C447996233FC530FBE
                            'FB11CB9836CA20F6BF49167B153A06F781EAEB13BAAF62B781D5BB460048E98033E3ED17F93F0DB01FBC7C4F26
                            'D6249B0A081A66647681705C2D916438E621425880E5AB431D34E835E34BBA7ED589CAB67A6B2FD20B0D50B0EA
                            '28791E05326B4834089026DFB960210120F29FABF8FCEDEF7E50BE8D9FF417C65AAE6C4BA1524F394CE1B363F9
                            '44852E4B22E1DF648CE5DEBBF45479E0841D5AD591E812E4A056102BF209C4EE77E98D96DB9D8EB1EEDB6BE099
                            'F2840B9745C6FA656AC588772AC8F74362DA4E1680829AC97122B8AAA3E20E27F8158FC396A57C234ED20D4B9BDB0E8F69CC5460A3538DB2A805273E
                            '7810C63D20BA3C990557A8163D5372FA4AA0F53AF074476BBB1373D7B28030A9A57E1054D0913554FC6A108E7D63BDA5BB557C8CBDC93719B842AF223
                            'A30C948EE612FA9D6939CA706F03FC080029C2BEBA0AE103EF5D03D0EFE15BD27CE9277DD8CE1E8885E4D019E7
                            '673A4BE7CBC85325960E050560A8DE678BEB5203515E3E42F6AA3F534FE0D8B8A65938265AF662D45C6C53C9A7F1D9B97FE44C619E0C74F97D09AC47F282C9C
                            '3DA4B5FC65D19FB45B52A93776D056772AC5CD23D3616D204547646015C6B687555BFCD47A60C7AA48C2B02B78CD8B4F49E3
                            '3062A0F9365531708FF34B20F797282828BCE8242C2316A59B98AF77A1CE62623D4A61CF022DA11D450FF973F665B85230DBDFEFABBEFDAADC6FBCB7335FE6DF9BE7CA57A0F
                            '0F37F65BF9471564A71FB8420C93B5E4F699DB97610AB50F47A586DA884A6ADE17A8503C833E97EA636A345CC8BD4914FA05D82A5E49AAB
                            'CC320CACE8CFD3E7F949B9C6AA7B08A349A52292BA9EDFD300155D28B9AA8FAFD70A3676B0F529EA64847A44C3450ACE4C8D8CCF60A597AB89F26230D48FE8E7259AB9530439F05F8E4
                            'C5136C8ED77E4705CC0FDB7E399A10136423D1847E36E95A5F0548A419FAFDCFC3AFD502DDF69179ADAABB60AA6362499E74ABA9DE3EE7C658702403FE4
                            '49FA2B144BE59F248FCF3C5EBCD830DC6C87C70A98C102B4D9003D33BA8F1676083AA28F8144D8184F6FD1A228DB8CEC7C975F2D
                            '57DC128C1174D836C289F10125CFEF225A84D68D4A782AF7D7D55A8E73C3344DB003F3A7B458B291185B9A5B840B
                        End If
                        If sm.UserInfo.UserID = str_superuser1 Then '系統管理者
                            If Not but_edit.Enabled Then
                                but_edit.Enabled = True
                                TIMS.Tooltip(but_edit, str_superuser1 & "權限啟動")
                            End If
                            If Not but_del.Enabled Then
                                but_del.Enabled = True
                                TIMS.Tooltip(but_del, str_superuser1 & "權限啟動")
                            End If
                        End If
                        If Convert.ToString(drv("PlanID")) = "0" Then
                            'lbtShare.Enabled = False '不可共用
                            'but_edit.Enabled = False '不可修改
                            but_del.Enabled = False '不可刪除
                            TIMS.Tooltip(but_del, "系統建立不可刪除")
                        End If

                    Case "N" '資料狀態 Y:正式 / N:審核中
                        but_edit.Visible = False
                        but_del.Visible = False
                        lbtShare.Visible = False
                        lbtChk.CommandArgument = "comidno=" & drv("comidno") & "&distid=" & drv("distid") & "&planid=" & drv("PlanID") & "&RSID=" & drv("RSID") & "&ID=" & Request("ID") & ""
                End Select
        End Select
    End Sub

    Sub KeepSearch()
        'TIMS.GetListValue()'
        Dim sSearch As String = ""
        sSearch = "TB_OrgName=" & TB_OrgName.Text
        sSearch += "&TB_ComIDNO=" & TB_ComIDNO.Text
        sSearch += "&TBCity=" & TBCity.Text
        sSearch += "&city_code=" & city_code.Value
        sSearch += "&zip_code=" & zip_code.Value
        sSearch += "&DistID=" & TIMS.GetListValue(DistID) '.SelectedValue
        sSearch += "&OrgKindList=" & TIMS.GetListValue(OrgKindList) '.SelectedValue
        sSearch += "&Years=" & TIMS.GetListValue(Yearlist) '.SelectedValue
        sSearch += "&drpPlan=" & TIMS.GetListValue(drpPlan) '.SelectedValue
        sSearch += "&IsApply=" & TIMS.GetListValue(IsApply) '.SelectedValue
        sSearch += "&PageIndex=" & DG_Org.CurrentPageIndex + 1
        sSearch += "&Button1=" & DG_Org.Visible
        sSearch += "&rblPlanPoint=" & TIMS.GetListValue(rblPlanPoint) '.SelectedValue
        sSearch += "&dl_typeid2=" & TIMS.GetListValue(dl_typeid2) '.SelectedValue
        Session("_Search") = sSearch
    End Sub

    ''' <summary>
    ''' 限定計畫啟用。
    ''' </summary>
    Sub Check_bt_add()
        '95年度各區新興 & 資軟由各區自行新增 by nick
        '2005/0615訓練計畫為新興科技人才培訓 or 資訊軟體人才培訓 and 轄區不為泰山者,不可新增機構
        If sm.UserInfo.TPlanID = "10" OrElse sm.UserInfo.TPlanID = "11" Then bt_add.Enabled = True
    End Sub

    ''' <summary>
    ''' 匯出
    ''' </summary>
    Sub Utl_EXPORT1()
        Dim flag_no_ok_sess As Boolean = False
        If Session("exp_tc_01_002_sql") Is Nothing Then flag_no_ok_sess = True
        If Session("exp_tc_01_002_parms") Is Nothing Then flag_no_ok_sess = True
        If flag_no_ok_sess Then
            msg.Text = "查無資料!!"
            Return
        End If
        Dim sqlstr As String = Session("exp_tc_01_002_sql")
        Dim parms As Hashtable = Session("exp_tc_01_002_parms")
        If String.IsNullOrEmpty(sqlstr) Then flag_no_ok_sess = True
        If parms Is Nothing Then flag_no_ok_sess = True
        If flag_no_ok_sess Then
            msg.Text = "查無資料!!"
            Return
        End If
        Dim objtable As DataTable
        objtable = DbAccess.GetDataTable(sqlstr, objconn, parms)
        If objtable Is Nothing Then
            msg.Text = "查無資料!!"
            Return
        End If
        If objtable.Rows.Count = 0 Then
            msg.Text = "查無資料!!"
            Return
        End If
        msg.Text = ""

        '編號,
        Const s_title1 As String = "轄區分署,機構別,機構名稱,計畫名稱,統編,地址,聯絡電話,保險證號,聯絡人姓名,聯絡人E-Mail,負責人姓名"
        Const s_data1 As String = "DISTNAME,OrgKindNAME,OrgName,PlanName,ComIDNO,Address,CONTACTCELLPHONE,ActNo,ContactName,ContactEmail,MasterName"
        Dim As_title1() As String = s_title1.Split(",")
        Dim As_data1() As String = s_data1.Split(",")

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("OrgPlanInfo", System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        Dim sFileName1 As String = "訓練機構檔"

        '套CSS值
        'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String '建立輸出文字
        'ExportStr = "編號" & vbTab & "轄區中心" & vbTab & "機構名稱" & vbTab & "計畫名稱" & vbTab & "統編" & vbTab & "地址" & vbTab & "保險證號" & vbTab & "聯絡人姓名" & vbTab & "聯絡人E-Mail" & vbTab
        ExportStr = "<tr>"
        ExportStr &= "<td>編號</td>" '& vbTab '& "轄區分署" & vbTab & "機構名稱" & vbTab & "計畫名稱" & vbTab & "統編" & vbTab & "地址" & vbTab & "保險證號" & vbTab & "聯絡人姓名" & vbTab & "聯絡人E-Mail" & vbTab
        For Each s_T1 As String In As_title1
            ExportStr &= "<td>" & s_T1 & "</td>"   '& vbTab
        Next
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim i_num As Integer = 0
        For Each oDr1 As DataRow In objtable.Rows
            i_num += 1
            ExportStr = "<tr>"
            ExportStr &= "<td>" & CStr(i_num) & "</td>"
            For Each s_D1 As String In As_data1
                ExportStr &= "<td>" & TIMS.ClearSQM(oDr1(s_D1)) & "</td>"
            Next
            ExportStr &= "</tr>"
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")
        objtable = Nothing

        Dim parmsExp As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)},
            {"FileName", sFileName1},
            {"strSTYLE", strSTYLE},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '匯出
    Private Sub Bt_EXPORT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_EXPORT.Click
        Utl_EXPORT1()
    End Sub

    Function ChangeImportDate(ByVal colArray As Array) As Array
        Const cst_000 As String = "00000000"
        colArray(Cst_ComIDNO) = Right(cst_000 & colArray(Cst_ComIDNO).ToString, 8) '廠商統一編號
        Return colArray
    End Function

    Function CheckImportData(ByVal colArray As Array) As String
        'Const cst_filedNum = 8
        Const cst_必須填寫 As String = "必須填寫"
        Dim Reason As String = ""
        Dim sql As String = ""
        Dim dr As DataRow = Nothing

        If colArray.Length < cst_filedNum Then
            'Reason += "欄位數量不正確(應該為" & cst_filedNum & "個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
        Else
            Dim ComIDNO As String = colArray(Cst_ComIDNO).ToString '統一編號
            Dim aYears As String = colArray(Cst_Years).ToString '年度
            Dim GradeDate As String = colArray(Cst_GradeDate).ToString '評鑑日期
            Dim FreeComments As String = colArray(Cst_FreeComments).ToString '是否免評
            Dim Point01A As String = colArray(Cst_Point01A).ToString '星等
            Dim Point01B As String = colArray(Cst_Point01B).ToString '分數
            Dim Point02A As String = colArray(Cst_Point02A).ToString '星等
            Dim Point02B As String = colArray(Cst_Point02B).ToString '分數
            Dim Point03A As String = colArray(Cst_Point03A).ToString '星等
            Dim Point03B As String = colArray(Cst_Point03B).ToString '分數
            Dim Point04A As String = colArray(Cst_Point04A).ToString '星等
            Dim Point04B As String = colArray(Cst_Point04B).ToString '分數
            Dim ClassCNames As String = colArray(Cst_ClassCNames).ToString '班級

            If ComIDNO = "" Then
                Reason += cst_必須填寫 & "廠商統編<Br>"
            Else
                If ComIDNO.Length <> 8 Then Reason += "廠商統一編號必須為8碼<BR>"
                If Not IsNumeric(ComIDNO) Then Reason += "廠商統一編號必須為數字<BR>"
                Dim pms As New Hashtable From {{"ComIDNO", ComIDNO}}
                sql = " SELECT * FROM ORG_ORGINFO WHERE ComIDNO=@ComIDNO"
                dr = DbAccess.GetOneRow(sql, objconn, pms)
                If dr Is Nothing Then Reason += "廠商統一編號必須存在於系統<BR>"
            End If
            If Trim(aYears) <> "" Then
                If Not IsNumeric(aYears) Then
                    Reason += "評鑑年度必須為正確的數字格式<BR>"
                Else
                    If dtKey_Years.Select("Years='" & aYears & "'").Length = 0 Then Reason += "評鑑年度" & "有錯，不符合鍵詞<BR>"
                End If
            Else
                Reason += cst_必須填寫 & "評鑑年度<Br>"
            End If
            If Trim(GradeDate) <> "" Then
                If Not IsDate(GradeDate) Then Reason += "評鑑日期必須為正確的日期格式<BR>"
            Else
                'Reason += cst_必須填寫 & "評鑑日期<Br>"
            End If

            Select Case FreeComments
                Case "Y"
                Case ""
                    If Trim(Point01A) <> "" Then
                        If IsNumeric(Point01A) = False Then Reason += "單位能力指標星等" & "必須要是數字<BR>"
                    Else
                        Reason += cst_必須填寫 & "單位能力指標星等<Br>"
                    End If
                    If Trim(Point02A) <> "" Then
                        If IsNumeric(Point02A) = False Then Reason += "就業表現指標星等" & "必須要是數字<BR>"
                    Else
                        Reason += cst_必須填寫 & "就業表現指標星等<Br>"
                    End If
                    If Trim(Point03A) <> "" Then
                        If IsNumeric(Point03A) = False Then Reason += "學員問卷滿意度指標星等" & "必須要是數字<BR>"
                    Else
                        Reason += cst_必須填寫 & "學員問卷滿意度指標星等<Br>"
                    End If
                    If Trim(Point04A) <> "" Then
                        If IsNumeric(Point04A) = False Then Reason += "總分指標星等" & "必須要是數字<BR>"
                    Else
                        Reason += cst_必須填寫 & "總分指標星等<Br>"
                    End If
                    If Trim(Point01B) <> "" Then
                        If Point01B.Length > 50 Then Reason += "單位能力指標分數" & "長度過長，必須要小於50 <BR>"
                    Else
                        Reason += cst_必須填寫 & "單位能力指標分數<Br>"
                    End If
                    If Trim(Point02B) <> "" Then
                        If Point01B.Length > 50 Then Reason += "就業表現指標分數" & "長度過長，必須要小於50 <BR>"
                    Else
                        Reason += cst_必須填寫 & "就業表現指標分數<Br>"
                    End If
                    If Trim(Point03B) <> "" Then
                        If Point01B.Length > 50 Then Reason += "學員問卷滿意度指標分數" & "長度過長，必須要小於50 <BR>"
                    Else
                        Reason += cst_必須填寫 & "學員問卷滿意度指標分數<Br>"
                    End If
                    If Trim(Point04B) <> "" Then
                        If Point01B.Length > 50 Then Reason += "總分指標分數" & "長度過長，必須要小於50 <BR>"
                    Else
                        Reason += cst_必須填寫 & "總分指標分數<Br>"
                    End If
                    If Trim(ClassCNames) <> "" Then
                        If ClassCNames.Length > 1000 Then Reason += "評鑑班級資料" & "字數超過長度1000<BR>"
                    Else
                        'Reason += cst_必須填寫 & "評鑑班級資料<Br>"
                    End If
                Case Else
                    Reason += "免評欄位只可以是Y或是不填寫<Br>"
            End Select
        End If
        Return Reason
    End Function

    Private Sub Btn_XlsImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_XlsImport.Click
        Const Cst_FileSavePath As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        Const Cst_FileExt1 As String = ".xls"
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "xls") Then Return

        Dim sql As String = " SELECT DISTINCT Years FROM ID_Plan WITH(NOLOCK) ORDER BY 1"
        dtKey_Years = DbAccess.GetDataTable(sql, objconn)

        sql = ""
        Dim MyFileName As String = ""
        'Dim MyFileType As String = ""
        If File1.Value <> "" Then
            ' 2. 取得原始檔名 (注意：某些舊版瀏覽器會傳回完整路徑，例如 C:\Users\Test\image.jpg)
            'Dim fileName As String = File1.PostedFile.FileName
            ' 3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
            '檢查檔案格式與大小----------   Start
            If File1.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                Exit Sub
            End If
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            If MyFileName.IndexOf(".") = -1 Then
                Common.MessageBox(Me, "檔案類型錯誤!")
                Exit Sub
            End If
            If fileNM_Ext <> Cst_FileExt1 Then
                Common.MessageBox(Me, $"檔案類型錯誤，必須為{Cst_FileExt1}檔!,'{fileNM_Ext }")
                Exit Sub
            End If

            Dim dt As DataTable = Nothing
            Dim Reason As String = "" '儲存錯誤的原因
            File1.PostedFile.SaveAs(Server.MapPath($"{Cst_FileSavePath}{MyFileName}"))  '上傳檔案

            '取得內容
            dt = TIMS.GetDataTable_XlsFile(Server.MapPath($"{Cst_FileSavePath}{MyFileName}"), "", Reason, "年度")
            'dt = TIMS.GetDataTable_XlsFile(Server.MapPath(Cst_FileSavePath & MyFileName).ToString, "PlanInfo", Reason, "廠商統一編號", "講師姓名", "課程名稱(單元)")
            IO.File.Delete(Server.MapPath($"{Cst_FileSavePath}{MyFileName}")) '刪除檔案
            If Reason <> "" Then
                Common.MessageBox(Me, Reason)
                Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝")
                Exit Sub
            End If

            'xls 方式 讀取寫入資料庫
            If dt.Rows.Count > 0 Then '有資料
                '將檔案讀出放入記憶體
                Dim RowIndex As Integer = 1
                'Dim OneRow As String
                Dim colArray As Array

                '取出資料庫的所有欄位--------   Start
                'Dim sql As String
                'Dim dr As DataRow
                Dim da As SqlDataAdapter = Nothing
                'Dim trans As SqlTransaction
                'Dim conn As SqlConnection = DbAccess.GetConnection
                'Dim dt As DataTable
                'Dim Reason As String            '儲存錯誤的原因
                Dim dtWrong As New DataTable     '儲存錯誤資料的DataTable
                Dim drWrong As DataRow

                '建立錯誤資料格式Table----------------Start
                dtWrong.Columns.Add(New DataColumn("Index"))
                dtWrong.Columns.Add(New DataColumn("ComIDNO"))
                dtWrong.Columns.Add(New DataColumn("GradeDate"))
                dtWrong.Columns.Add(New DataColumn("OrgName"))
                dtWrong.Columns.Add(New DataColumn("Reason"))
                '建立錯誤資料格式Table----------------End

                For i As Integer = 0 To dt.Rows.Count - 1
                    Reason = ""
                    colArray = dt.Rows(i).ItemArray
                    colArray = ChangeImportDate(colArray) '轉換正確欄位值
                    Reason += CheckImportData(colArray) '檢查正確欄位值

                    '通過檢查，開始輸入資料---------------------Start
                    If Reason = "" Then
                        G_OrgID = TIMS.Get_OrgIDforComIDNO(objconn, colArray(Cst_ComIDNO).ToString)
                        'G_Years = CDate(colArray(Cst_GradeDate).ToString).Year
                        G_Years = colArray(Cst_Years).ToString
                        If colArray(Cst_GradeDate).ToString.Trim <> "" Then
                            G_GradeDate = Common.FormatDate(colArray(Cst_GradeDate).ToString)
                        Else
                            G_GradeDate = ""
                        End If
                        G_FreeComments = colArray(Cst_FreeComments).ToString.Trim

                        Select Case G_FreeComments
                            Case "Y"
                                G_Point01A = ""
                                G_Point01B = ""
                                G_Point02A = ""
                                G_Point02B = ""
                                G_Point03A = ""
                                G_Point03B = ""
                                G_Point04A = ""
                                G_Point04B = ""
                                G_ClassCNames = ""
                            Case Else
                                G_Point01A = colArray(Cst_Point01A).ToString
                                G_Point01B = colArray(Cst_Point01B).ToString
                                G_Point02A = colArray(Cst_Point02A).ToString
                                G_Point02B = colArray(Cst_Point02B).ToString
                                G_Point03A = colArray(Cst_Point03A).ToString
                                G_Point03B = colArray(Cst_Point03B).ToString
                                G_Point04A = colArray(Cst_Point04A).ToString
                                G_Point04B = colArray(Cst_Point04B).ToString
                                G_ClassCNames = colArray(Cst_ClassCNames).ToString
                        End Select

                        If Not Save_Org_Comments(Reason) Then
                            '錯誤資料，填入錯誤資料表
                            drWrong = dtWrong.NewRow
                            dtWrong.Rows.Add(drWrong)
                            drWrong("Index") = RowIndex
                            If colArray.Length > 5 Then
                                drWrong("ComIDNO") = colArray(Cst_ComIDNO).ToString
                                If IsDate(colArray(Cst_GradeDate).ToString) Then
                                    drWrong("GradeDate") = Common.FormatDate(colArray(Cst_GradeDate).ToString)
                                Else
                                    drWrong("GradeDate") = "" & colArray(Cst_GradeDate).ToString
                                End If
                                drWrong("OrgName") = "--"
                                'If IsNothing(TIMS.GetOCIDDate(TIMS.Get_OrgIDforComIDNO(colArray(Cst_ComIDNO).ToString))) Then
                                '    drWrong("OrgName") = "--"
                                'Else
                                '    drWrong("OrgName") = "" & TIMS.GetOCIDDate(TIMS.Get_OrgIDforComIDNO(colArray(Cst_ComIDNO).ToString))("OrgName")
                                'End If
                                drWrong("Reason") = Reason
                            End If
                        End If
                    Else
                        '錯誤資料，填入錯誤資料表
                        drWrong = dtWrong.NewRow
                        dtWrong.Rows.Add(drWrong)
                        drWrong("Index") = RowIndex
                        If colArray.Length > 5 Then
                            drWrong("ComIDNO") = colArray(Cst_ComIDNO).ToString
                            If IsDate(colArray(Cst_GradeDate).ToString) Then
                                drWrong("GradeDate") = Common.FormatDate(colArray(Cst_GradeDate).ToString)
                            Else
                                drWrong("GradeDate") = "" & colArray(Cst_GradeDate).ToString
                            End If
                            drWrong("OrgName") = "--"
                            'If IsNothing(TIMS.GetOCIDDate(TIMS.Get_OrgIDforComIDNO(colArray(Cst_ComIDNO).ToString))) Then
                            '    drWrong("OrgName") = "--"
                            'Else
                            '    drWrong("OrgName") = "" & TIMS.GetOCIDDate(TIMS.Get_OrgIDforComIDNO(colArray(Cst_ComIDNO).ToString))("OrgName")
                            'End If
                            drWrong("Reason") = Reason
                        End If
                    End If
                    RowIndex += 1
                Next

                '判斷匯出資料是否有誤
                Dim explain, explain2 As String
                explain = ""
                explain += "匯入資料共" & dt.Rows.Count & "筆" & vbCrLf
                explain += "成功：" & (dt.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
                explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
                explain2 = ""
                explain2 += "匯入資料共" & dt.Rows.Count & "筆\n"
                explain2 += "成功：" & (dt.Rows.Count - dtWrong.Rows.Count) & "筆\n"
                explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

                '開始判別欄位存入------------   End
                If dtWrong.Rows.Count = 0 Then
                    Common.MessageBox(Me, explain)
                Else
                    Session("MyWrongTable") = dtWrong
                    Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視失敗原因?')){window.open('TC_01_002_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
                End If
            End If
            'MyFile.Delete(Server.MapPath(Cst_FileSavePath & MyFileName))
        End If
    End Sub

    ''' <summary>
    ''' 儲存 (機構評鑑)
    ''' </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function Save_Org_Comments(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = False 'True:正常結束 False:異常結束
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim Trans As SqlTransaction = Nothing
        Dim sql As String = ""
        'Dim conn As SqlConnection
        Try
            'conn = DbAccess.GetConnection()
            Trans = DbAccess.BeginTrans(objconn)
            sql = " SELECT * FROM Org_Comments WHERE OrgID = '" & G_OrgID & "' AND Years = '" & G_Years & "'" & vbCrLf
            dt = DbAccess.GetDataTable(sql, da, Trans)
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
            Else
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("OrgID") = G_OrgID
                dr("Years") = G_Years
            End If
            If G_GradeDate <> "" Then
                dr("GradeDate") = G_GradeDate
            Else
                dr("GradeDate") = Convert.DBNull
            End If
            Select Case G_FreeComments
                Case "Y"
                    dr("FreeComments") = G_FreeComments
                    dr("Point01A") = Convert.DBNull
                    dr("Point01B") = Convert.DBNull
                    dr("Point02A") = Convert.DBNull
                    dr("Point02B") = Convert.DBNull
                    dr("Point03A") = Convert.DBNull
                    dr("Point03B") = Convert.DBNull
                    dr("Point04A") = Convert.DBNull
                    dr("Point04B") = Convert.DBNull
                    dr("ClassCNames") = Convert.DBNull
                Case Else
                    dr("FreeComments") = Convert.DBNull
                    dr("Point01A") = G_Point01A
                    dr("Point01B") = G_Point01B
                    dr("Point02A") = G_Point02A
                    dr("Point02B") = G_Point02B
                    dr("Point03A") = G_Point03A
                    dr("Point03B") = G_Point03B
                    dr("Point04A") = G_Point04A
                    dr("Point04B") = G_Point04B
                    dr("ClassCNames") = G_ClassCNames
            End Select
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now()
            DbAccess.UpdateDataTable(dt, da, Trans)
            DbAccess.CommitTrans(Trans)
            Errmsg = ""
            Rst = True
            'Return True
        Catch ex As Exception
            DbAccess.RollbackTrans(Trans)
            Errmsg = "機構評鑑新增異常!!:" & ex.Message
            'Return False
        End Try
        Return Rst
    End Function

    ''' <summary>取得PlanName 前綴詞。</summary>
    Function SGetPlanName(ByVal PLANID As String, ByVal RID As String) As String
        Dim rst As String = ""
        Dim pms1 As New Hashtable From {{"PLANID", PLANID}, {"RID", RID}}
        Dim Sqls As String = ""
        Sqls &= " SELECT concat(c.YEARS,d.NAME,e.PLANNAME,c.SEQ,'_') PlanName" & vbCrLf
        Sqls &= " FROM AUTH_RELSHIP a" & vbCrLf
        Sqls &= " JOIN ORG_ORGINFO b ON a.ORGID=b.ORGID" & vbCrLf
        Sqls &= " JOIN ID_PLAN c ON c.PlanID=a.PlanID" & vbCrLf
        Sqls &= " JOIN ID_DISTRICT d ON d.DistID=c.DistID" & vbCrLf
        Sqls &= " JOIN KEY_PLAN e ON c.TPlanID=e.TPlanID" & vbCrLf
        Sqls &= " WHERE a.PLANID=@PLANID AND a.RID=@RID" & vbCrLf
        rst = $"{DbAccess.ExecuteScalar(Sqls, objconn, pms1)}"
        Return rst
    End Function

#Region "2018 機構屬性設定(merge from TC_01_017)"

    ''' <summary>
    ''' 產投計畫子計畫類別單選項目連動(機構屬性)機構別下拉事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RblPlanPoint_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rblPlanPoint.SelectedIndexChanged
        Dim v_rblPlanPoint As String = TIMS.GetListValue(rblPlanPoint)
        Call TIMS.GET_DDL_TYPEID12(dl_typeid2, objconn, v_rblPlanPoint)
    End Sub

#Region "(No Use)"
    'Select Case v_rblPlanPoint'.SelectedValue
    '    Case "1", "2"
    '        get_ddl_typeid2(v_rblPlanPoint) '.SelectedValue)
    '    Case Else
    '        dl_typeid2.Items.Clear()
    '        dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
    '        dl_typeid2.SelectedIndex = 0
    '        dl_typeid2.Enabled = False
    'End Select

    ' 依產投計畫子計畫類別單選項目選擇結果連動查詢顯示(機構屬性)機構別下拉
    '    Private Sub get_ddl_typeid2(ByVal TypeID1 As String)
    '        If TypeID1 = "" OrElse TypeID1 = "0" Then Exit Sub

    '        Dim sql As String = ""
    '        sql &= " SELECT TypeID1 ,TypeID2 ,concat(TypeID2,'-',TypeID2Name) TypeID2Name" & vbCrLf
    '        sql &= " FROM KEY_ORGTYPE1" & vbCrLf
    '        sql &= " WHERE TypeID1=@TypeID1" & vbCrLf
    '        Call TIMS.OpenDbConn(objconn)
    '        Dim parms As New Hashtable From {{"TypeID1", CInt(TypeID1)}}
    '        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

    '        If dt.Rows.Count > 0 Then
    '            dl_typeid2.Items.Clear()
    '            dl_typeid2.DataSource = dt
    '            dl_typeid2.DataValueField = "TypeID2"
    '            dl_typeid2.DataTextField = "TypeID2Name"
    '            dl_typeid2.DataBind()
    '            dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
    '            dl_typeid2.SelectedIndex = 0
    '            dl_typeid2.Enabled = True
    '        Else
    '            dl_typeid2.Items.Clear()
    '            dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
    '            dl_typeid2.SelectedIndex = 0
    '            dl_typeid2.Enabled = False
    '        End If
    '    End Sub
#End Region

    ''' <summary>
    ''' ajax 動態產生 訓練機構屬性設定-機構別下拉選項
    ''' </summary>
    ''' <param name="typeid1">1.產業人才投資計畫 2.提升勞工自主學習計畫</param>
    Public Sub ResponseTypeID2(ByVal typeid1 As String)

        Dim selTag As TagBuilder = New TagBuilder("select")

        Dim sql As String = ""
        sql &= " SELECT TypeID1 ,TypeID2 ,TypeID2 + '-' + TypeID2Name TypeID2Name" & vbCrLf
        sql &= " FROM Key_OrgType1" & vbCrLf
        sql &= " WHERE TypeID1=@TypeID1" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim parms As New Hashtable From {{"TypeID1", CInt(typeid1)}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        Dim optTag As TagBuilder = Nothing
        Dim hasData As Boolean = False

        If Not IsNothing(dt) Then
            Dim dr As DataRow
            If dt.Rows.Count > 0 Then
                optTag = New TagBuilder("option")
                optTag.Attributes.Add("value", "== 請選擇 ==")
                optTag.InnerHtml = ""
                selTag.InnerHtml += optTag.ToString()
            End If
            For Each dr In dt.Rows
                hasData = True
                optTag = New TagBuilder("option")
                optTag.Attributes.Add("value", dr("TypeID2"))
                optTag.InnerHtml = dr("TypeID2Name")
                selTag.InnerHtml += optTag.ToString()
            Next
        End If

        If Not hasData Then
            optTag = New TagBuilder("option")
            optTag.Attributes.Add("value", "")
            optTag.InnerHtml = "== 請選擇 =="
            selTag.InnerHtml += optTag.ToString()
        End If

        Response.Clear()
        Response.Write(selTag.InnerHtml)
        Response.End()
    End Sub
#End Region
End Class