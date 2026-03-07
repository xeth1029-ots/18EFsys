Partial Class CP_04_003
    Inherits AuthBasePage

    'CP_04_003_add.aspx?yearlist=" & Me.yearlist.SelectedValue & "&ID=" & Request("ID") '查詢
    'CP_04_003_add.aspx?export=Y&yearlist=" & Me.yearlist.SelectedValue & "&ID=" & Request("ID") '匯出

    Const cst_search As String = "_search" ' Session(cst_search) = strSession 'CP_04_003_add.aspx 使用
    Const cst_search2 As String = "_search2" 'KeepSearch
    'Dim sqlstr As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)

        If Not Page.IsPostBack Then
            Call SUtl_Create1()
        End If

        '20180920 依承辦人需求,將[訓練計畫]欄位先隱藏起來
        TPlan_item_TR.Visible = False
        TIMS.SetCblValue(PlanList, sm.UserInfo.TPlanID)
    End Sub

    Sub SUtl_Create1()
        '檢查日期格式
        SSTDate.Attributes("onchange") = "check_date();"
        ESTDate.Attributes("onchange") = "check_date();"
        SFTDate.Attributes("onchange") = "check_date();"
        EFTDate.Attributes("onchange") = "check_date();"

        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');" '選擇全部轄區
        CityList.Attributes("onclick") = "SelectAll('CityList','CityHidden');" '選擇全部縣市
        PlanList.Attributes("onclick") = "SelectAll('PlanList','TPlanHidden');" '選擇全部訓練計畫

        '取得訓練計畫
        'TPlan = TIMS.Get_TPlan(TPlan, , 1)
        'Call TIMS.Get_TPlan2(chkTPlanID0, chkTPlanID1, chkTPlanIDX, objconn, 1)
        'chkTPlanID0.Attributes("onclick") = "SelectAll('chkTPlanID0','TPlanID0HID');"
        'chkTPlanID1.Attributes("onclick") = "SelectAll('chkTPlanID1','TPlanID1HID');"
        'chkTPlanIDX.Attributes("onclick") = "SelectAll('chkTPlanIDX','TPlanIDXHID');"

        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)

        DistrictList = TIMS.Get_DistID(DistrictList, Nothing, objconn)
        If (DistrictList.Items.FindByValue("") IsNot Nothing) Then DistrictList.Items.Remove(DistrictList.Items.FindByValue(""))
        ' DistrictList.Items.Remove( DistrictList.Items.FindByValue(""))
        DistrictList.Items.Insert(0, New ListItem("全部", ""))

        PlanList = TIMS.Get_TPlan(PlanList, , 1, "Y", , objconn) '計畫
        'PlanList.Items.FindByValue(sm.UserInfo.TPlanID).Selected = True

        CityList = TIMS.Get_CityName(CityList, TIMS.dtNothing) '縣市

        Dim sqlstr As String = " SELECT TMID, concat( BUSID,'.',BUSNAME) BUSNAME FROM KEY_TRAINTYPE WHERE LEVELS=0 ORDER BY TMID"
        Dim dtKT As DataTable = DbAccess.GetDataTable(sqlstr, objconn)
        With TMID
            .DataSource = dtKT
            .DataTextField = "BusName"
            .DataValueField = "TMID"
            .DataBind()
            .Items.Insert(0, New ListItem("全部", ""))
        End With

        '預算來源(Key_Budget、view_Budget)
        'sqlstr = " SELECT * FROM VIEW_BUDGET ORDER BY BUDID "
        'dt = DbAccess.GetDataTable(sqlstr, objconn)
        ' BudgetList.DataSource = dt
        ' BudgetList.DataTextField = "BudName"
        ' BudgetList.DataValueField = "BudID"
        ' BudgetList.DataBind()
        ' BudgetList.Items.Remove( BudgetList.Items.FindByValue("99")) '移除不補助

        Call UseKeepSession2()

        'Dim flagLID23 As Boolean = TIMS.Chk_Relship23(Me, objconn)
        '當分署(中心)使用者使用時,轄區應該都要鎖死該轄區,不可選擇其它轄區
        HIDOrgID.Value = ""
        Select Case sm.UserInfo.LID '階層代碼【0:署(局) 1:分署(中心) 2:委訓/補助地方政府】
            Case "0"
                '是本署(本局) '完全不鎖定
            Case "1"
                '是分署(中心)
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False
            Case "2"
                '是一般機構
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False '轄區
                Common.SetListItem(yearlist, sm.UserInfo.Years)
                yearlist.Enabled = False '年度
                TIMS.SetCblValue(PlanList, sm.UserInfo.TPlanID)
                PlanList.Enabled = False '計畫
                OrgName.Text = sm.UserInfo.OrgName
                HIDOrgID.Value = sm.UserInfo.OrgID
                OrgName.Enabled = False '計畫
            Case Else
                '異常情況處理。
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False '轄區
                Common.SetListItem(yearlist, sm.UserInfo.Years)
                yearlist.Enabled = False '年度
                TIMS.SetCblValue(PlanList, sm.UserInfo.TPlanID)
                PlanList.Enabled = False '計畫
                OrgName.Text = sm.UserInfo.OrgName
                HIDOrgID.Value = sm.UserInfo.OrgID
                OrgName.Enabled = False '計畫
                'DistrictList.Style.Item("display") = "none"
        End Select
    End Sub

    Private Sub UseKeepSession2()
        Dim MyValue As String = ""
        Dim sSession_search2 As String = ""
        If Session(cst_search2) IsNot Nothing AndAlso Convert.ToString(Session(cst_search2)) <> "" Then
            sSession_search2 = Convert.ToString(Session(cst_search2))
            ClearKeepSearch12() 'Session(cst_search2) = Nothing
        End If

        MyValue = TIMS.GetMyValue(sSession_search2, "yearlist")
        If MyValue <> "" Then Common.SetListItem(yearlist, MyValue)

        MyValue = TIMS.GetMyValue(sSession_search2, "DistrictList")
        If MyValue <> "" Then TIMS.SetCblValue(DistrictList, MyValue)

        MyValue = TIMS.GetMyValue(sSession_search2, "CityList")
        If MyValue <> "" Then TIMS.SetCblValue(CityList, MyValue)

        MyValue = TIMS.GetMyValue(sSession_search2, "PlanList")
        If MyValue <> "" Then TIMS.SetCblValue(PlanList, MyValue)

        SSTDate.Text = TIMS.GetMyValue(sSession_search2, "SSTDate")
        ESTDate.Text = TIMS.GetMyValue(sSession_search2, "ESTDate")
        SFTDate.Text = TIMS.GetMyValue(sSession_search2, "SFTDate")
        EFTDate.Text = TIMS.GetMyValue(sSession_search2, "EFTDate")
        MyValue = TIMS.GetMyValue(sSession_search2, "NotOpenStaus")
        If MyValue <> "" Then TIMS.SetCblValue(NotOpenStaus, MyValue)

        OrgName.Text = TIMS.GetMyValue(sSession_search2, "OrgName")
        HIDOrgID.Value = TIMS.GetMyValue(sSession_search2, "OrgID")
        ClassCName.Text = TIMS.GetMyValue(sSession_search2, "ClassCName")

        MyValue = TIMS.GetMyValue(sSession_search2, "TMID")
        If MyValue <> "" Then Common.SetListItem(TMID, MyValue)

        'MyValue = TIMS.GetMyValue(sSession_search2, "BudgetList")
        'If MyValue <> "" Then TIMS.SetCblValue(BudgetList, MyValue)
        'Session(cst_search2) = Nothing
    End Sub

    Sub KeepSearch12()
        Call ClearKeepSearch12()
        '選擇轄區
        Dim itemstr As String = ""
        For Each objitem As ListItem In DistrictList.Items
            If objitem.Selected = True AndAlso objitem.Value <> "" Then
                itemstr &= String.Concat(If(itemstr <> "", ",", ""), "'", objitem.Value, "'")
            End If
        Next

        '報表要用的轄區參數
        Dim DistID As String = ""
        Dim DistName As String = ""
        For i As Integer = 1 To DistrictList.Items.Count - 1
            If DistrictList.Items(i).Selected AndAlso DistrictList.Items(i).Value <> "" Then
                DistID &= String.Concat(If(DistID <> "", ",", ""), "\'" & DistrictList.Items(i).Value & "\'")
                DistName &= String.Concat(If(DistName <> "", ",", ""), DistrictList.Items(i).Text)
            End If
        Next
        If DistID <> "" AndAlso DistrictList.Items(0).Selected Then DistName = "全部"

        '選擇縣市
        Dim itemcity As String = ""
        For Each objitem As ListItem In CityList.Items
            If objitem.Selected = True AndAlso objitem.Value <> "" Then
                itemcity &= String.Concat(If(itemcity <> "", ",", ""), "'" & objitem.Value & "'")
            End If
        Next

        '報表要用的縣市參數
        Dim ICity As String = ""
        Dim ICityName As String = ""
        'ICity = ""
        'ICityName = ""
        For i As Integer = 1 To CityList.Items.Count - 1
            If CityList.Items(i).Selected AndAlso CityList.Items(i).Value <> "" Then
                ICity &= String.Concat(If(ICity <> "", ",", ""), "\'" & CityList.Items(i).Value & "\'")
                ICityName &= String.Concat(If(ICityName <> "", ",", ""), CityList.Items(i).Text)
            End If
        Next
        If ICity <> "" AndAlso CityList.Items(0).Selected Then ICityName = "全部"

        '選擇訓練計畫
        Dim itemplan As String = ""
        For Each objitem As ListItem In PlanList.Items
            If objitem.Selected = True AndAlso objitem.Value <> "" Then
                itemplan &= String.Concat(If(itemplan <> "", ",", ""), "'" & objitem.Value & "'")
            End If
        Next

        '報表要用的訓練計畫參數
        'Dim TPlanID As String = ""
        'Dim TPlanID3 As String = ""
        'TPlanID3 = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 3)
        'Dim TPlanName5 As String = ""
        'TPlanName5 = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 5)
        'If TPlanName5 = "" Then TPlanName5 = "全部"

        Dim TPlanID As String = ""
        Dim TPlanName As String = ""
        For i As Integer = 1 To PlanList.Items.Count - 1
            If PlanList.Items(i).Selected AndAlso PlanList.Items(i).Value <> "" Then
                TPlanID &= String.Concat(If(TPlanID <> "", ",", ""), "\'" & PlanList.Items(i).Value & "\'")
                TPlanName &= String.Concat(If(TPlanName <> "", ",", ""), PlanList.Items(i).Text)
            End If
        Next
        If TPlanID <> "" AndAlso PlanList.Items(0).Selected Then TPlanName = "全部"

        '選擇開班狀態
        Dim sNotOpenStaus As String = ""
        Dim sNotOpenStausStr As String = ""
        If NotOpenStaus.Items(0).Selected AndAlso NotOpenStaus.Items(1).Selected Then
            sNotOpenStausStr = "開班,不開班"
        Else
            If NotOpenStaus.SelectedIndex = 0 Then '0 開班
                sNotOpenStaus = "N"
                sNotOpenStausStr = "開班"
            ElseIf NotOpenStaus.SelectedIndex = 1 Then '1 不開班
                sNotOpenStaus = "Y"
                sNotOpenStausStr = "不開班"
            End If
        End If

        ''選擇預算來源
        'Dim itembudget As String = ""
        'For Each objitem As ListItem In BudgetList.Items
        '    If objitem.Selected = True AndAlso objitem.Value.ToString <> "" Then
        '        If itembudget <> "" Then itembudget &= ","
        '        itembudget &= "'" & objitem.Value.ToString & "'"
        '    End If
        'Next

        ''報表要用的預算來源參數
        'Dim BudgetID As String = ""
        'Dim BudgetName As String = ""
        'BudgetID = ""
        'BudgetName = ""
        'For i As Integer = 0 To BudgetList.Items.Count - 1
        '    If BudgetList.Items(i).Selected AndAlso BudgetList.Items(i).Value <> "" Then
        '        If BudgetID <> "" Then BudgetID &= ","
        '        BudgetID &= "\'" & BudgetList.Items(i).Value & "\'"
        '        If BudgetName <> "" Then BudgetName &= ","
        '        BudgetName &= BudgetList.Items(i).Text
        '    End If
        'Next

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim v_TMID As String = TIMS.GetListValue(TMID)
        Dim v_TMID_txt As String = TIMS.GetListText(TMID)

        Dim strSession As String
        strSession = ""
        TIMS.SetMyValue(strSession, "yearlist", v_yearlist) '下拉
        TIMS.SetMyValue(strSession, "DistrictList", TIMS.GetCblValue(DistrictList))    'cbl
        TIMS.SetMyValue(strSession, "CityList", TIMS.GetCblValue(CityList)) 'cbl
        TIMS.SetMyValue(strSession, "PlanList", TIMS.GetCblValue(PlanList)) 'cbl
        TIMS.SetMyValue(strSession, "SSTDate", SSTDate.Text)
        TIMS.SetMyValue(strSession, "ESTDate", ESTDate.Text)
        TIMS.SetMyValue(strSession, "SFTDate", SFTDate.Text)
        TIMS.SetMyValue(strSession, "EFTDate", EFTDate.Text)
        TIMS.SetMyValue(strSession, "NotOpenStaus", TIMS.GetCblValue(NotOpenStaus))  'cbl
        TIMS.SetMyValue(strSession, "OrgName", OrgName.Text)
        TIMS.SetMyValue(strSession, "OrgID", HIDOrgID.Value)
        TIMS.SetMyValue(strSession, "ClassCName", ClassCName.Text)
        TIMS.SetMyValue(strSession, "TMID", v_TMID) '下拉
        'TIMS.SetMyValue(strSession, "BudgetList", TIMS.GetCblValue(BudgetList))  'cbl
        Session(cst_search2) = strSession '查詢條件儲存(本頁條件使用)

        strSession = ""
        TIMS.SetMyValue(strSession, "itemstr", itemstr) '選擇轄區
        TIMS.SetMyValue(strSession, "itemplan", itemplan) '選擇訓練計畫
        'TIMS.SetMyValue(strSession, "itemcity", itemcity) '選擇縣市
        TIMS.SetMyValue(strSession, "SSTDate", SSTDate.Text)
        TIMS.SetMyValue(strSession, "ESTDate", ESTDate.Text)
        TIMS.SetMyValue(strSession, "SFTDate", SFTDate.Text)
        TIMS.SetMyValue(strSession, "EFTDate", EFTDate.Text)
        TIMS.SetMyValue(strSession, "NotOpenStaus", sNotOpenStaus) '(N,Y,"")
        TIMS.SetMyValue(strSession, "newDistID", DistID) '選擇轄區 (報表)
        TIMS.SetMyValue(strSession, "newTPlanID", TPlanID) '選擇訓練計畫 (報表)
        TIMS.SetMyValue(strSession, "TMID", v_TMID) '下拉
        'TIMS.SetMyValue(strSession, "itembudget", itembudget) '選擇預算來源
        'TIMS.SetMyValue(strSession, "newBudgetID", BudgetID) '選擇預算來源(報表)
        TIMS.SetMyValue(strSession, "OrgName", TIMS.ChangeSQM(OrgName.Text))
        TIMS.SetMyValue(strSession, "OrgID", HIDOrgID.Value)
        TIMS.SetMyValue(strSession, "ClassCName", TIMS.ChangeSQM(ClassCName.Text))
        TIMS.SetMyValue(strSession, "newICityName", ICityName) '報表
        TIMS.SetMyValue(strSession, "newTPlanIDName", TPlanName) '中文
        TIMS.SetMyValue(strSession, "NotOpenStausStr", sNotOpenStausStr) '中文
        TIMS.SetMyValue(strSession, "TMIDName", v_TMID_txt) '中文
        'TIMS.SetMyValue(strSession, "newBudgetName", BudgetName) '中文
        Dim v_RBListExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS
        TIMS.SetMyValue(strSession, "RBListExpType", v_RBListExpType)
        Session(cst_search) = strSession 'CP_04_003_add.aspx 使用

        Session("itemcity") = itemcity '選擇縣市
        Session("newICity") = ICity '選擇縣市(報表用)
    End Sub

    '查詢
    Private Sub Bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call KeepSearch12()
        Dim url1 As String = "CP_04_003_add.aspx?ID=" & TIMS.Get_MRqID(Me) & "&yearlist=" & TIMS.GetListValue(yearlist) 'yearlist.SelectedValue
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '重新設定
    Private Sub Bt_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reset.Click
        'Reset
        Common.SetListItem(yearlist, sm.UserInfo.Years)
        TIMS.SetCblValue(DistrictList, "")
        TIMS.SetCblValue(CityList, "")
        TIMS.SetCblValue(PlanList, "")

        SSTDate.Text = ""
        ESTDate.Text = ""
        SFTDate.Text = ""
        EFTDate.Text = ""
        OrgName.Text = ""
        ClassCName.Text = ""
        NotOpenStaus.SelectedIndex = 0
        TMID.SelectedIndex = 0

        Call ClearKeepSearch12()
    End Sub

    Private Sub ClearKeepSearch12()
        If Session(cst_search2) IsNot Nothing Then Session(cst_search2) = Nothing 'strSession '查詢條件儲存(本頁條件使用)
        If Session(cst_search) IsNot Nothing Then Session(cst_search) = Nothing 'strSession 'CP_04_003_add.aspx 使用
        If Session("itemcity") IsNot Nothing Then Session("itemcity") = Nothing 'itemcity '選擇縣市
        If Session("newICity") IsNot Nothing Then Session("newICity") = Nothing 'ICity '選擇縣市(報表用)
    End Sub

    ''' <summary> 匯出 CP_04_003_add.SUB_EXPORT </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Bt_export_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_export.Click
        Call KeepSearch12()
        Dim url1 As String = $"CP_04_003_add.aspx?ID={TIMS.Get_MRqID(Me)}&export=Y&yearlist={TIMS.GetListValue(yearlist)}" 'Me.yearlist.SelectedValue
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class