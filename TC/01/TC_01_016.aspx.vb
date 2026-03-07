Partial Class TC_01_016
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DG_Org
        '分頁設定 End

        Dim flagP As Boolean = CHECK_Acct_Permission()
        If Not flagP Then Exit Sub

        If Not Me.IsPostBack Then
            Call cCreate1()

            Common.SetListItem(Years, sm.UserInfo.Years)
            Common.SetListItem(drpPlan, sm.UserInfo.TPlanID)
            Common.SetListItem(DistID, sm.UserInfo.DistID)
            '取得查詢條件
            If Not Session("_Search") Is Nothing Then
                TB_OrgName.Text = TIMS.GetMyValue(Session("_Search"), "TB_OrgName")
                TB_ComIDNO.Text = TIMS.GetMyValue(Session("_Search"), "TB_ComIDNO")
                TBCity.Text = TIMS.GetMyValue(Session("_Search"), "TBCity")
                city_code.Value = TIMS.GetMyValue(Session("_Search"), "city_code")
                zip_code.Value = TIMS.GetMyValue(Session("_Search"), "zip_code")
                Common.SetListItem(DistID, TIMS.GetMyValue(Session("_search"), "DistID"))
                Common.SetListItem(OrgKindList, TIMS.GetMyValue(Session("_search"), "OrgKindList"))
                Common.SetListItem(Years, TIMS.GetMyValue(Session("_search"), "Years"))
                Common.SetListItem(drpPlan, TIMS.GetMyValue(Session("_search"), "drpPlan"))
                'PageControler1.PageIndex = TIMS.GetMyValue(Session("_search"), "PageIndex")
                Me.ViewState("PageIndex") = TIMS.GetMyValue(Session("_search"), "PageIndex")
                If TIMS.GetMyValue(Session("_search"), "Button1") = "True" Then
                    Me.ViewState("PageIndex") = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
                    If IsNumeric(Me.ViewState("PageIndex")) Then PageControler1.PageIndex = Me.ViewState("PageIndex")
                    bt_search_Click(sender, e)
                End If
                Session("_Search") = Nothing
            End If
        End If
        bt_search.Attributes("onclick") = "return Search();"  '搜尋時Client端檢查
    End Sub

    Sub cCreate1()
        Dim sql As String = ""
        'Dim dt As DataTable

        'Years.Items.Clear()
        ''Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " SELECT DISTINCT YEARS, dbo.FN_CYEAR2B(years) ROC_YEARS" & vbCrLf
        'sql &= " FROM dbo.ID_Plan WITH(NOLOCK)" & vbCrLf
        'sql &= " ORDER BY YEARS DESC" & vbCrLf
        'dt = DbAccess.GetDataTable(sql, objconn)
        'If dt.Rows.Count > 0 Then
        '    With Years
        '        .DataSource = dt
        '        .DataTextField = If(flag_ROC, "ROC_YEARS", "YEARS") 'edit，by:20181002
        '        .DataValueField = "Years"
        '        .DataBind()
        '        If TypeOf Years Is DropDownList Then .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        '    End With
        '    Common.SetListItem(Years, "")
        'End If

        Years = TIMS.Get_Years(Years, objconn)
        '找不到空值放一個
        Dim im As ListItem = Years.Items.FindByValue("")
        If im Is Nothing Then Years.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

        drpPlan = TIMS.Get_TPlan(drpPlan)

        ''移除下列計畫
        'drpPlan.Items.Remove(drpPlan.Items.FindByValue("17")) '補助地方政府訓練
        'drpPlan.Items.Remove(drpPlan.Items.FindByValue("28")) '產業人才投資方案
        'drpPlan.Items.Remove(drpPlan.Items.FindByValue("36")) '青年職涯啟動計畫

        DistID = TIMS.Get_DistID(DistID, TIMS.dtNothing(), objconn)
        OrgKindList = TIMS.Get_OrgType(OrgKindList, objconn)
        drpAppliedResult = Get_AppliedResult(drpAppliedResult)
        'Common.SetListItem(drpPlan, sm.UserInfo.TPlanID)
    End Sub

    '取得鍵值-計畫審核狀態
    Public Shared Function Get_AppliedResult(ByVal obj As DropDownList) As DropDownList
        With obj
            .Items.Clear()
            '.Items.Add(New ListItem("==請選擇==", ""))
            .Items.Add(New ListItem("通過", "Y"))
            .Items.Add(New ListItem("不通過", "N"))
            .Items.Add(New ListItem("未審核", "X"))
        End With
        Return obj
    End Function

    ''' <summary>
    ''' 權限檢核
    ''' </summary>
    ''' <returns></returns>
    Function CHECK_Acct_Permission() As Boolean
        '0:署(局)  1:分署(中心)  2:委訓
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                bt_search.Enabled = False
                Common.MessageBox(Me, "開放給勞動力發展署階層以上者，權限不足，無法使用!!")
                Return False
        End Select
        '0:超級使用者
        '1:系統管理者
        '2:一級以上
        '3:一級
        '4:二級
        '5:承辦人
        '99:一般使用者
        Select Case sm.UserInfo.RoleID
            Case "0", "1"
            Case Else
                bt_search.Enabled = False
                Common.MessageBox(Me, "非系統管理者以上權限者，權限不足，無法使用!!")
                Return False
        End Select
        Return True
    End Function

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Me.Panel.Visible = True
        Dim flagP As Boolean = CHECK_Acct_Permission()
        If Not flagP Then Exit Sub

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " SELECT b.* " & vbCrLf
        sqlstr &= " ,idt.Name DistName" & vbCrLf
        sqlstr &= " ,op.Address ,op.ActNo ,op.ContactName ,op.ContactEmail " & vbCrLf
        sqlstr &= " FROM dbo.VIEW_RWPLANRID b " & vbCrLf
        sqlstr &= " JOIN dbo.ORG_ORGPLANINFO op ON op.RSID = b.RSID " & vbCrLf
        sqlstr &= " JOIN dbo.ID_DISTRICT idt ON idt.distid = b.distid " & vbCrLf
        sqlstr &= " WHERE 1=1 " & vbCrLf
        '不限定
        'sqlstr += " AND b.PlanID !=0 " & vbCrLf'限定委訓單位
        sqlstr &= " AND b.OrgLevel >= '" & sm.UserInfo.OrgLevel & "' " & vbCrLf
        If zip_code.Value <> "" Then
            sqlstr &= " AND op.ZipCode = '" & Me.zip_code.Value & "' " & vbCrLf
        ElseIf city_code.Value <> "" Then
            sqlstr &= " AND op.ZipCode IN (SELECT zipcode FROM ID_Zip WHERE ctid = '" & Me.city_code.Value & "') " & vbCrLf
        End If
        If Me.TB_OrgName.Text <> "" Then sqlstr &= " AND b.OrgName LIKE '%" & Me.TB_OrgName.Text & "%' " & vbCrLf
        If Me.TB_ComIDNO.Text <> "" Then sqlstr &= " AND b.ComIDNO = '" & Me.TB_ComIDNO.Text & "' " & vbCrLf

        Dim v_OrgKindList As String = TIMS.GetListValue(OrgKindList)
        Dim v_DistID As String = TIMS.GetListValue(DistID)
        Dim v_Years As String = TIMS.GetListValue(Years)
        Dim v_drpPlan As String = TIMS.GetListValue(drpPlan)
        If v_OrgKindList <> "" Then sqlstr &= " AND b.OrgKind = '" & v_OrgKindList & "' " & vbCrLf
        If v_DistID <> "" Then sqlstr &= " AND b.DistID = '" & v_DistID & "' " & vbCrLf
        If v_Years <> "" Then sqlstr &= " AND b.Years = '" & v_Years & "' " & vbCrLf
        If v_drpPlan <> "" Then sqlstr &= " AND b.TPlanID = '" & v_drpPlan & "' " & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        If dt.Rows.Count = 0 Then
            Panel.Visible = False
            DG_Org.Visible = False
            msg.Text = "查無資料!!"
            Return
        End If

        '重整 PlanName
        Dim sub_plan_name As String = ""
        Dim sub_orgName As String = ""
        Dim sub_sql As String = ""
        Dim Parent_list_RID As String = ""

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim list As DataRow = dt.Rows(i)
            If list("PlanID") <> 0 Then
                'Dim myarray As Array
                'myarray = list("Relship").Split("/")
                'Dim range As Integer = myarray.Length - 3
                'If range >= 0 Then
                '    Parent_list = myarray(range)
                'Else
                '    Common.MessageBox(Me, "計畫權限取得有誤，請重新輸入查詢值!!")
                '    Exit Sub
                'End If
                If Convert.ToString(list("ORGLEVEL")) <> "2" Then
                    Common.MessageBox(Me, "計畫權限取得有誤，請重新輸入查詢值!!")
                    Exit Sub
                End If
                Parent_list_RID = Left(list("RID"), 1)
                sub_sql = "SELECT b.OrgName AS PlanName FROM AUTH_RELSHIP a JOIN ORG_ORGINFO b ON a.orgid = b.orgid WHERE a.RID = '" & Parent_list_RID & "' "
                sub_orgName = Convert.ToString(DbAccess.ExecuteScalar(sub_sql, objconn))

                sub_sql = "" & vbCrLf
                sub_sql &= " SELECT c.Years + d.Name + e.PlanName + c.seq + '_' PlanName " & vbCrLf
                sub_sql &= " FROM Auth_Relship a " & vbCrLf
                sub_sql &= " JOIN org_orginfo b ON a.orgid = b.orgid " & vbCrLf
                sub_sql &= " JOIN ID_Plan c ON c.PlanID = a.PlanID " & vbCrLf
                sub_sql &= " JOIN ID_District d ON d.DistID = c.DistID " & vbCrLf
                sub_sql &= " JOIN Key_Plan e ON c.TPlanID = e.TPlanID " & vbCrLf
                sub_sql &= " WHERE a.planid = '" & list("RWPlanID") & "' " & vbCrLf
                sub_sql &= " AND a.RID = '" & list("RID") & "' " & vbCrLf
                sub_plan_name = Convert.ToString(DbAccess.ExecuteScalar(sub_sql, objconn))
                list("PlanName") = sub_plan_name + sub_orgName
            End If
        Next

        'Me.bt_save.Visible = True
        Panel.Visible = True
        msg.Text = ""
        DG_Org.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Sub KeepSearch()
        Dim v_OrgKindList As String = TIMS.GetListValue(OrgKindList)
        Dim v_DistID As String = TIMS.GetListValue(DistID)
        Dim v_Years As String = TIMS.GetListValue(Years)
        Dim v_drpPlan As String = TIMS.GetListValue(drpPlan)

        Dim v_Search1 As String = ""
        v_Search1 = "TB_OrgName=" & TB_OrgName.Text
        v_Search1 &= "&TB_ComIDNO=" & TB_ComIDNO.Text
        v_Search1 &= "&TBCity=" & TBCity.Text
        v_Search1 &= "&city_code=" & city_code.Value
        v_Search1 &= "&zip_code=" & zip_code.Value
        v_Search1 &= "&DistID=" & v_DistID 'DistID.SelectedValue
        v_Search1 &= "&OrgKindList=" & v_OrgKindList 'OrgKindList.SelectedValue
        v_Search1 &= "&Years=" & v_Years 'Years.SelectedValue
        v_Search1 &= "&drpPlan=" & v_drpPlan 'drpPlan.SelectedValue
        v_Search1 &= "&PageIndex=" & DG_Org.CurrentPageIndex + 1
        v_Search1 &= "&Button1=" & DG_Org.Visible
        Session("_Search") = v_Search1
    End Sub

    Private Sub DG_Org_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Org.ItemCommand
        Select Case e.CommandName
            Case "modify"
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "../01/TC_01_016_add.aspx?ProcessType=modify&" & e.CommandArgument & "")
        End Select
    End Sub

    Private Sub DG_Org_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Org.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim but_modify As LinkButton = e.Item.FindControl("lbtModify")
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex
                TIMS.Tooltip(but_modify, "計畫權限代碼：" & drv("RWPlanRID"))
                'sCmdArg1 = "RWPlanRID=" & drv("RWPlanRID") & "&orgid=" & drv("orgid") & "&planid=" & drv("RWPlanID") & 
                '"&rid=" & drv("RID") & "&RSID=" & drv("RSID") & "&distid=" & drv("distid") & "&ID=" & Request("ID") & ""  '不使用 PlanID 使用 RWPlanID

                Dim v_drpAppliedResult As String = TIMS.GetListValue(drpAppliedResult)

                Dim sCmdArg1 As String = ""
                sCmdArg1 = "ID=" & Request("ID") & ""
                sCmdArg1 += "&RWPlanRID=" & drv("RWPlanRID")
                sCmdArg1 += "&orgid=" & drv("orgid")
                sCmdArg1 += "&planid=" & drv("RWPlanID")
                sCmdArg1 += "&rid=" & drv("RID")
                sCmdArg1 += "&RSID=" & drv("RSID")
                sCmdArg1 += "&distid=" & drv("distid")
                sCmdArg1 += "&AppliedResult=" & v_drpAppliedResult 'drpAppliedResult.SelectedValue

                but_modify.CommandArgument = sCmdArg1
        End Select
    End Sub

    Protected Sub DG_Org_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DG_Org.SelectedIndexChanged
    End Sub
End Class