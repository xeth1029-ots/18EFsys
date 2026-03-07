Partial Class TC_01_014
    Inherits AuthBasePage

    '參數/變數 設定
    'Plan_VerReport
    'Const Cst_編號 As Integer = 0
    'Const cst_func As Integer = 6 '功能欄位在(5)
    'Const Cst_審核狀態 As Integer = 8
    Const cst_PlanID As String = "PlanID"
    Const cst_ComIDNO As String = "ComIDNO"
    Const cst_SeqNO As String = "SeqNO"
    Const cst_ProcessType As String = "ProcessType" 'ProcessType @Insert/Update/View
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印
    'Dim au As New cAUTH
    Dim gFlagEnv As Boolean = True '正式環境。(測試用) / TestStr
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DG_Org
        '分頁設定 End

        If TIMS.sUtl_ChkTest() Then gFlagEnv = False
        tr_center.Visible = False
        If TIMS.IsSuperUser(Me, 1) Then tr_center.Visible = True

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            bt_search.Attributes("onclick") = "return Search();"

            '取得查詢條件
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            yearlist = TIMS.GetSyear(yearlist, 0, sm.UserInfo.Years, True)
            Common.SetListItem(yearlist, sm.UserInfo.Years)

            '取得查詢條件
            If Not Session("_Search") Is Nothing Then
                Dim MyValue As String = ""
                MyValue = TIMS.GetMyValue(Session("_Search"), "PlanYear")
                If MyValue <> "" Then Common.SetListItem(yearlist, MyValue)
                MyValue = TIMS.GetMyValue(Session("_Search"), "TxtPageSize")
                If MyValue <> "" Then TxtPageSize.Text = MyValue
                MyValue = TIMS.GetMyValue(Session("_Search"), "PageIndex")
                If MyValue <> "" Then PageControler1.PageIndex = MyValue
                MyValue = TIMS.GetMyValue(Session("_Search"), "Button1")
                If MyValue = "True" Then Call sSearch1()
                Session("_Search") = Nothing
            End If
        End If

        '#End Region
    End Sub

    Sub sSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DG_Org)
        Dim sDistID As String = ""
        If RIDValue.Value <> "" Then
            sDistID = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ' ' times" & vbCrLf
        sql &= " ,CONCAT(ppi.PlanID,'x',ppi.ComIDNO,'x',ppi.SeqNo) PCS" & vbCrLf
        sql &= " ,ppi.PlanYear" & vbCrLf
        sql &= " ,ppi.CYCLTYPE" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(ppi.CLASSNAME,ppi.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,ISNULL(ppi.ClassCate, ' ') ClassCate" & vbCrLf
        sql &= " ,ISNULL(pvr.PlanID, 0) pvrPlanID" & vbCrLf 'NULL為0
        sql &= " ,ISNULL(pvr.ClassID,'0') ClassID" & vbCrLf 'NULL為"0"
        sql &= " ,ppi.TPlanID" & vbCrLf
        sql &= " ,ppi.PlanID" & vbCrLf
        sql &= " ,ppi.TMID" & vbCrLf
        sql &= " ,ppi.ComIDNO" & vbCrLf
        sql &= " ,ppi.SeqNO" & vbCrLf
        sql &= " ,ppi.RID" & vbCrLf
        sql &= " ,ppi.TNum" & vbCrLf
        sql &= " ,ppi.THours" & vbCrLf
        sql &= " ,ppi.STDate" & vbCrLf
        sql &= " ,ppi.FDDate" & vbCrLf
        sql &= " ,ppi.PointYN" & vbCrLf
        sql &= " ,pvr.FirResult" & vbCrLf
        sql &= " ,pvr.SecResult" & vbCrLf
        sql &= " ,ppi.DefGovCost" & vbCrLf
        sql &= " ,ppi.DefStdCost" & vbCrLf
        sql &= " ,ppi.CapDegree" & vbCrLf
        sql &= " ,kcc.CCName ClassCateText" & vbCrLf
        sql &= " ,ooi.OrgName" & vbCrLf
        sql &= " FROM PLAN_PLANINFO ppi" & vbCrLf
        sql &= " JOIN ORG_ORGINFO ooi ON ppi.ComIDNO = ooi.ComIDNO" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid = ppi.planid" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP rr ON rr.RID = ppi.RID" & vbCrLf
        sql &= " LEFT JOIN PLAN_VERREPORT pvr ON ppi.PlanID = pvr.PlanID AND ppi.ComIDNO = pvr.ComIDNO AND ppi.SeqNO = pvr.SeqNo" & vbCrLf
        sql &= " LEFT JOIN KEY_CLASSCATELOG kcc ON ppi.ClassCate = kcc.CCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND ppi.IsApprPaper='Y'" & vbCrLf '正式送出
        sql &= " AND ip.TPLANID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        If sm.UserInfo.LID = 0 Then
            sql &= " AND ip.DISTID='" & sDistID & "'" & vbCrLf
        Else
            sql &= " AND ip.DISTID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If

        Select Case IsApprPaper.SelectedIndex
            Case 0 'SelectedIndex: 正式:0 草稿:1 SelectedValue: Y/N
                sql &= " AND pvr.IsApprPaper = 'Y'" & vbCrLf '正式
            Case 1  'SelectedIndex: 正式:0 草稿:1 SelectedValue: Y/N
                sql &= " AND pvr.IsApprPaper = 'N'" & vbCrLf '草稿
            Case Else
                'sqlstr += " AND pvr.IsApprPaper='N'" & vbCrLf '草稿
        End Select

        center.Text = TIMS.ClearSQM(center.Text)
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)

        Dim myParam As Hashtable = New Hashtable

        '訓練機構
        If center.Text <> "" Then
            sql &= " AND ooi.OrgName LIKE @orgName" & vbCrLf
            myParam.Add("orgName", "%" + center.Text.Trim + "%")
        End If

        '年度
        If yearlist.SelectedValue <> "" Then
            sql &= " AND ppi.PlanYear = @PlanYear" & vbCrLf
            myParam.Add("PlanYear", yearlist.SelectedValue.Trim)
        End If

        '班級名稱
        If ClassName.Text <> "" Then
            sql &= " AND ISNULL(ppi.ClassName, '') LIKE @className" & vbCrLf
            myParam.Add("className", "%" + ClassName.Text.Trim + "%")
        End If

        '期別
        If CyclType.Text <> "" Then
            If IsNumeric(CyclType.Text) Then
                If Int(CyclType.Text) < 10 Then
                    sql &= " AND ppi.CyclType= '0" & Int(CyclType.Text) & "'" & vbCrLf
                Else
                    sql &= " AND ppi.CyclType= '" & CyclType.Text & "'" & vbCrLf
                End If
            End If
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, myParam)
        Session("TC_table") = dt

        Panel.Visible = False
        DG_Org.Visible = False
        msg.Text = "查無資料!!(請確認登入帳號及班級申請，謝謝)"
        If dt.Rows.Count = 0 Then Exit Sub

        Panel.Visible = True
        DG_Org.Visible = True
        msg.Text = ""

        PageControler1.PrimaryKey = "PCS"
        PageControler1.Sort = "PCS"
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call sSearch1()
    End Sub

    Sub KeepSearch()
        Dim ssSearch As String = ""
        ssSearch = "PlanYear=" & yearlist.SelectedValue
        ssSearch += "&TxtPageSize=" & TxtPageSize.Text
        ssSearch += "&PageIndex=" & DG_Org.CurrentPageIndex + 1
        ssSearch += "&Button1=" & DG_Org.Visible
        Session("_Search") = ssSearch
    End Sub

    Private Sub DG_Org_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Org.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Select Case e.CommandName
            Case "add"
                Call KeepSearch()
                Dim url1 As String = "TC_01_014_add.aspx?ID=" & Request("ID") & "&ProcessType=Insert&" & e.CommandArgument & ""
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "edit"
                Call KeepSearch()
                Dim url1 As String = "TC_01_014_add.aspx?ID=" & Request("ID") & "&ProcessType=Update&" & e.CommandArgument & "&IsApprPaper=" & IsApprPaper.SelectedValue & ""
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "del"
                Dim rPlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
                Dim rComIDNO As String = TIMS.GetMyValue(sCmdArg, "ComIDNO")
                Dim rSeqNO As String = TIMS.GetMyValue(sCmdArg, "SeqNO")
                If rPlanID = "" Then Exit Sub
                If rComIDNO = "" Then Exit Sub
                If rSeqNO = "" Then Exit Sub
                Dim sqlstr As String = ""
                sqlstr = " DELETE Plan_VerReport WHERE PlanID=" & rPlanID & " and ComIDNO='" & rComIDNO & "' and SeqNo=" & rSeqNO
                DbAccess.ExecuteNonQuery(sqlstr, objconn)
                Call sSearch1()
            Case "view"
                Call KeepSearch()
                Dim url1 As String = "TC_01_014_add.aspx?ID=" & Request("ID") & "&ProcessType=View&" & e.CommandArgument & ""
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    Private Sub DG_Org_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Org.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim add_but As LinkButton = e.Item.FindControl("add_but") '新增
                Dim edit_but As LinkButton = e.Item.FindControl("edit_but") '修改
                Dim del_but As LinkButton = e.Item.FindControl("del_but") '刪除
                Dim view_but As LinkButton = e.Item.FindControl("view_but") '查詢
                Dim LSecResult As Label = e.Item.FindControl("LSecResult") '審核狀態
                Dim HidPCS As HiddenField = e.Item.FindControl("HidPCS")
                Dim HidClassCate As HiddenField = e.Item.FindControl("HidClassCate") '課程類別id
                Dim HidClassID As HiddenField = e.Item.FindControl("HidClassID") '課程班別ID

                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DG_Org.PageSize * DG_Org.CurrentPageIndex

                HidPCS.Value = Convert.ToString(drv("PCS"))
                HidClassCate.Value = Convert.ToString(drv("ClassCate"))
                HidClassID.Value = Convert.ToString(drv("ClassID"))
                Dim sCmdArg As String = ""
                sCmdArg = ""
                sCmdArg &= "PlanYear=" & Convert.ToString(drv("PlanYear"))
                sCmdArg &= "&PlanID=" & Convert.ToString(drv("PlanID"))
                sCmdArg &= "&TPlanID=" & Convert.ToString(drv("TPlanID"))
                sCmdArg &= "&TMID=" & Convert.ToString(drv("TMID"))
                sCmdArg &= "&RID=" & Convert.ToString(drv("RID"))
                sCmdArg &= "&ComIDNO=" & Convert.ToString(drv("ComIDNO"))
                sCmdArg &= "&SeqNO=" & Convert.ToString(drv("SeqNO"))
                add_but.CommandArgument = sCmdArg '新增
                edit_but.CommandArgument = sCmdArg '修改
                del_but.CommandArgument = sCmdArg '刪除
                del_but.Attributes("onclick") = "javascript:return confirm('此動作會刪除開班計畫表資料，是否確定刪除?');"
                view_but.CommandArgument = sCmdArg '查詢
                Dim SecResult As String = Convert.ToString(drv("SecResult"))
                Dim sSecResult2 As String = "尚未審核"
                Select Case SecResult 'null/R/Y/N
                    Case "Y"
                        sSecResult2 = "審核通過"
                    Case "N"
                        sSecResult2 = "<font color=red>審核不通過</font>"
                    Case "R"
                        sSecResult2 = "退件修正"
                    Case Else
                        If CStr(drv("pvrPlanID")) = "0" Then
                            sSecResult2 = "<font color=red>尚未新增</font>"
                        End If
                End Select
                LSecResult.Text = sSecResult2
                edit_but.Visible = False '修改
                del_but.Visible = False '刪除
                view_but.Visible = False '查詢
                add_but.Visible = False '新增
                Select Case Convert.ToString(sm.UserInfo.LID)
                    Case "2"
                        If CStr(drv("pvrPlanID")) = "0" Then
                            add_but.Visible = True '可新增(委訓單位)
                        End If
                        If CStr(drv("pvrPlanID")) <> "0" Then
                            '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
                            Select Case SecResult 'null/R/Y/N
                                Case "Y", "N"
                                    view_but.Visible = True
                                Case "R"
                                    '委訓單位可重新編輯
                                    edit_but.Visible = True
                                Case Else
                                    view_but.Visible = True
                            End Select
                        End If
                    Case Else
                        If CStr(drv("pvrPlanID")) <> "0" Then
                            '分署(中心)與署(局)只能查看，無法修改
                            view_but.Visible = True '查詢
                        End If
                End Select
        End Select
    End Sub

    Private Sub IsApprPaper_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IsApprPaper.SelectedIndexChanged
        Call sSearch1()
    End Sub
End Class