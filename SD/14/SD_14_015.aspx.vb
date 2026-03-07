Partial Class SD_14_015
    Inherits AuthBasePage

    'SD_14_015_1_2010 'W(勞工)
    'SD_14_015_2010 'G(產投)
    '檢查 TC_03_006.CheckCostItemTable (function)

    'Dim iPYNum14 As Integer = 1 'TIMS.sUtl_GetPYNum14(Me)
    Dim prtFilename As String = "" '列印表件名稱

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        'iPYNum14 = TIMS.sUtl_GetPYNum14(Me)

        msg.Text = ""
        PageControler1.PageDataGrid = DataGrid1
        PageControler2.PageDataGrid = DataGrid2

        If Not IsPostBack Then
            CCreate1()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        If Me.Radio1.SelectedIndex = 1 Then
            TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
            If HistoryTable.Rows.Count <> 0 Then
                OCID1.Attributes("onclick") = "showObj('HistoryList');"
                OCID1.Style("CURSOR") = "hand"
            End If
        End If

        Years.Value = sm.UserInfo.Years - 1911
        AllPrint.Visible = False
    End Sub

    Private Sub CCreate1()
        DataGridTable.Visible = False
        ClassTR.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        Me.Radio1.SelectedIndex = 0

        PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
        Common.SetListItem(PlanPoint, "1")

        AppStage = TIMS.Get_AppStage(AppStage)
        TIMS.SET_MY_APPSTAGE_LIST_VAL(Me, AppStage)

        Dim s_javascript_btn2 As String = ""
        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
        Button2.Attributes("onclick") = s_javascript_btn2
        '列印
        Button3.Attributes("onclick") = "return CheckPrint();"
        Button4.Attributes("onclick") = "ClearData();"
        'AllPrint.Attributes("onclick") = "SelectAll3(this.checked);"
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Radio1.SelectedIndex = 0 Then
            'If Me.TxtPageSize.Text <> Me.DataGrid2.PageSize Then Me.DataGrid2.PageSize = Me.TxtPageSize.Text
            TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid2)
            Call CreatePlan()
        Else
            'If Me.TxtPageSize.Text <> Me.DataGrid1.PageSize Then Me.DataGrid1.PageSize = Me.TxtPageSize.Text
            TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
            Call CreateClass()
        End If
    End Sub

    '(未轉班)依計畫
    Private Sub CreatePlan()

        OCIDValue.Value = ""
        SeqNoValue.Value = ""
        PlanIDValue.Value = ""
        ComIDNOValue.Value = ""

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim SearchStr As String = ""
        If sm.UserInfo.RID = "A" Then
            SearchStr += " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' and Years='" & sm.UserInfo.Years & "')"
        Else
            SearchStr += " and PlanID='" & sm.UserInfo.PlanID & "'"
        End If

        Dim v_AppStage As String = TIMS.GetListValue(AppStage)
        If (v_AppStage <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage
        If AppStage.SelectedIndex <> 0 AndAlso v_AppStage <> "" Then SearchStr &= " and AppStage ='" & v_AppStage & "'"

        SearchStr += " and RID IN (SELECT RID FROM Auth_Relship  WHERE RelShip like '" & RelShip & "%')"

        Dim sql As String = ""
        sql &= " SELECT a.PlanID,a.ComIDNO,a.SeqNo" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) ClassCName" & vbCrLf
        sql &= " ,a.STDate ,a.FDDate FTDate ,b.OrgName" & vbCrLf
        sql &= " FROM (SELECT * FROM Plan_PlanInfo WHERE TransFlag ='N' AND IsApprPaper='Y'" & SearchStr & ") a" & vbCrLf
        sql &= " JOIN view_RIDName b ON a.RID=b.RID" & vbCrLf

        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If PlanPoint.SelectedValue = "1" Then
                '產業人才投資計畫
                sql &= " AND b.OrgKind <> '10'"
                orgid.Value = "G"
            Else
                '提升勞工自主學習計畫
                sql &= " AND b.OrgKind = '10'"
                orgid.Value = "W"
            End If
            'sql += " ) b ON a.RID=b.RID "
        Else
            orgid.Value = ""
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        DataGrid1.Visible = False
        PageControler1.Visible = False
        DataGrid2.Visible = False
        PageControler2.Visible = False

        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            DataGrid2.Visible = True
            PageControler2.Visible = True

            PageControler2.PageDataTable = dt
            PageControler2.ControlerLoad()
        End If


    End Sub

    '(已轉班)依班別
    Private Sub CreateClass()

        OCIDValue.Value = ""
        SeqNoValue.Value = ""
        PlanIDValue.Value = ""
        ComIDNOValue.Value = ""

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)
        Dim SearchStr As String = ""
        If sm.UserInfo.RID = "A" Then
            SearchStr += " and cc.PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' and Years='" & sm.UserInfo.Years & "')"
        Else
            SearchStr += " and cc.PlanID='" & sm.UserInfo.PlanID & "'"
        End If

        Dim v_AppStage As String = TIMS.GetListValue(AppStage)
        If (v_AppStage <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage
        If AppStage.SelectedIndex <> 0 AndAlso v_AppStage <> "" Then SearchStr &= " and pp.AppStage ='" & v_AppStage & "'"

        SearchStr += " and cc.RID IN (SELECT RID FROM Auth_Relship WHERE RelShip like '" & RelShip & "%')"

        If OCIDValue1.Value <> "" Then
            SearchStr += " and cc.OCID='" & OCIDValue1.Value & "'"
        End If

        'sql = "SELECT orgid FROM Auth_Relship WHERE RID='" & RIDValue.Value & "'" '判斷產學訓的計畫別是否為勞工團體
        'orgidkind = DbAccess.ExecuteScalar(sql)
        'orgid.Value = TIMS.Get_OrgKind(orgidkind.ToString)
        'sql = "SELECT  a.OCID,a.ClassCName+case CyclType when '00' then '' else '第'+convert(varchar,convert(int,CyclType))+'期' end as ClassCName,a.STDate,a.FTDate,b.OrgName FROM "
        'sql += "(SELECT cc.* FROM Class_ClassInfo cc join Plan_PlanInfo pp on cc.planid = pp.planid and cc.rid = pp.rid and cc.seqno = pp.seqno  WHERE 1=1" & SearchStr & ") a "
        'sql += " JOIN view_RIDName b ON a.RID=b.RID "

        Dim sql As String = "" ' & vbCrLf
        sql &= " SELECT  a.OCID ,dbo.FN_GET_CLASSCNAME(a.ClassCName,a.CyclType) ClassCName" & vbCrLf
        sql &= " ,a.STDate,a.FTDate,b.OrgName" & vbCrLf
        sql &= " FROM (  SELECT cc.* FROM Class_ClassInfo cc" & vbCrLf
        sql &= "  join Plan_PlanInfo pp on cc.planid = pp.planid and cc.rid = pp.rid and cc.seqno = pp.seqno" & vbCrLf
        sql &= "  WHERE pp.TransFlag ='Y' AND pp.IsApprPaper='Y'" & vbCrLf
        sql &= SearchStr
        sql &= " ) a" & vbCrLf
        sql &= " JOIN view_RIDName b   ON a.RID=b.RID" & vbCrLf

        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'sql += "JOIN (select * from view_RIDName where "
            If PlanPoint.SelectedValue = "1" Then
                '產業人才投資計畫
                sql &= " AND b.OrgKind <> '10'"
                orgid.Value = "G"
            Else
                '提升勞工自主學習計畫
                sql &= " AND b.OrgKind = '10'"
                orgid.Value = "W"
            End If
            'sql += " ) b ON a.RID=b.RID "
        Else
            orgid.Value = ""
            'KindValue.Value = TIMS.GetTPlanName(sm.UserInfo.TPlanID)
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        DataGrid1.Visible = False
        PageControler1.Visible = False
        DataGrid2.Visible = False
        PageControler2.Visible = False

        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            DataGrid1.Visible = True
            PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If


    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SD_TD2"
                End If
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HtmlInputCheckBox = e.Item.FindControl("OCID")
                OCID.Value = drv("OCID")
                OCID.Attributes("onclick") = "SelectOCID(this.checked,this.value);"

                Dim OCIDArray As Array = Split(OCIDValue.Value, ",")
                For i As Integer = 0 To OCIDArray.Length - 1
                    If drv("OCID").ToString = OCIDArray(i) Then
                        OCID.Checked = True
                    End If
                Next
        End Select
    End Sub

    Private Sub Radio1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio1.SelectedIndexChanged
        If Radio1.SelectedIndex = 0 Then
            ClassTR.Visible = False
        Else
            ClassTR.Visible = True
        End If
        DataGridTable.Visible = False
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SD_TD2"
                End If
                Dim drv As DataRowView = e.Item.DataItem
                Dim SeqNo As HtmlInputCheckBox = e.Item.FindControl("SeqNo")
                Dim PlanID As HtmlInputHidden = e.Item.FindControl("PlanID")
                Dim ComIDNO As HtmlInputHidden = e.Item.FindControl("ComIDNO")

                SeqNo.Value = drv("SeqNo")
                PlanID.Value = drv("PlanID")
                ComIDNO.Value = drv("ComIDNO")
                SeqNo.Attributes("onclick") = "SelectSeqNo(this.checked,this.value,'" & PlanID.Value & "','" & ComIDNO.Value & "');"

                Dim SeqNoArray As Array = Split(SeqNoValue.Value, ",")
                For i As Integer = 0 To SeqNoArray.Length - 1
                    If drv("SeqNo").ToString = SeqNoArray(i) Then
                        SeqNo.Checked = True
                    End If
                Next
        End Select
    End Sub

End Class
