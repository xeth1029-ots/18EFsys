'Imports System.Data.SqlClient
'Imports System.Data
'Imports Turbo
Partial Class SD_04_011
    Inherits AuthBasePage

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

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
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = dtPlan
        '分頁設定 End

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), v_yearlist)

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        'btnQuery.Enabled = False
        'If blnCanSech Then btnQuery.Enabled = True

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False

            Dim iYears1 As Integer = Val(sm.UserInfo.Years) - 1
            Dim iYears2 As Integer = Val(sm.UserInfo.Years) + 1
            yearlist = TIMS.GetSyear(yearlist, iYears1, iYears2, True)

            '2005/4/1--Melody年度帶預設值
            Common.SetListItem(yearlist, sm.UserInfo.Years)
            '(加強操作便利性)
            RIDValue.Value = sm.UserInfo.RID
            center.Text = sm.UserInfo.OrgName 'orgname
        End If

        '帶入查詢參數
        If Not IsPostBack Then
            If Session("search") IsNot Nothing Then
                Dim MyValue As String = ""

                MyValue = TIMS.GetMyValue(Session("search"), "yearlist")
                Common.SetListItem(yearlist, MyValue)

                center.Text = TIMS.GetMyValue(Session("search"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("search"), "RIDValue")
                start_date.Text = TIMS.GetMyValue(Session("search"), "start_date")
                end_date.Text = TIMS.GetMyValue(Session("search"), "end_date")

                MyValue = TIMS.GetMyValue(Session("search"), "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If

                'btnQuery_Click(sender, e)
                Call sSearch1()

                Session("search") = Nothing
            End If
        End If
    End Sub

    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dtPlan) '顯示列數不正確

        dtPlan.CurrentPageIndex = 0
        Call LoadData()
    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call sSearch1()
    End Sub

    Private Sub dtPlan_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dtPlan.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1
        End If
    End Sub

    Sub LoadData()
        Dim sql As String = ""
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select oo.orgid,oo.orgname,cc.ocid,ip.Years" & vbCrLf
        sql &= " from dbo.ORG_ORGINFO oo" & vbCrLf
        sql &= " join dbo.PLAN_PLANINFO pp on oo.comidno =pp.comidno" & vbCrLf
        sql &= " join dbo.CLASS_CLASSINFO cc on cc.planid =pp.planid and pp.comidno =cc.comidno and cc.seqno =pp.seqno" & vbCrLf
        sql &= " join dbo.ID_PLAN ip on ip.planid=cc.planid" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        'sql &= " and rownum <=10" & vbCrLf
        sql &= " and ip.TPlanID = '" & sm.UserInfo.TPlanID & "'"
        Select Case RIDValue.Value '(開頭若為A~G使用LIKE 其他使用=)
            Case "A"
                sql &= " and pp.RID like 'A%' "
            Case "B"
                sql &= " and pp.RID like 'B%' "
            Case "C"
                sql &= " and pp.RID like 'C%' "
            Case "D"
                sql &= " and pp.RID like 'D%' "
            Case "E"
                sql &= " and pp.RID like 'E%' "
            Case "F"
                sql &= " and pp.RID like 'F%' "
            Case "G"
                sql &= " and pp.RID like 'G%' "
            Case Else
                sql &= " and pp.RID='" & RIDValue.Value & "'"
        End Select

        If yearlist.SelectedValue <> "" Then
            sql &= " and ip.Years='" & yearlist.SelectedValue & "'"
        End If
        sql &= " )" & vbCrLf

        '排課若是假日代碼是'9999999',計算機構的排課時數排除假日
        sql &= " ,WC2 AS (" & vbCrLf
        sql &= " select cc.orgname" & vbCrLf
        sql &= " ,sum(" & vbCrLf
        sql &= " CASE WHEN (cs.class1 is not NULL and cs.class1 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class2 is not NULL and cs.class2 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class3 is not NULL and cs.class3 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class4 is not NULL and cs.class4 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class5 is not NULL and cs.class5 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class6 is not NULL and cs.class6 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class7 is not NULL and cs.class7 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class8 is not NULL and cs.class8 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class9 is not NULL and cs.class9 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class10 is not NULL and cs.class10 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class11 is not NULL and cs.class11 <> '9999999') then 1 else 0 end +" & vbCrLf
        sql &= " CASE WHEN (cs.class12 is not NULL and cs.class12 <> '9999999') then 1 else 0 end ) c12" & vbCrLf
        sql &= " FROM dbo.CLASS_SCHEDULE cs" & vbCrLf
        sql &= " JOIN WC1 cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " WHERE cs.FORMAL ='Y'" & vbCrLf
        If yearlist.SelectedValue <> "" Then
            sql &= " and cc.Years='" & yearlist.SelectedValue & "'"
        End If
        If start_date.Text <> "" Then
            sql &= " and cs.schoolDate >= " & TIMS.To_date(Me.start_date.Text) & vbCrLf
        End If
        If end_date.Text <> "" Then
            sql &= " and cs.schoolDate<= " & TIMS.To_date(Me.end_date.Text) & vbCrLf
        End If
        sql &= " GROUP BY cc.orgname" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " select cc.orgname" & vbCrLf
        sql &= " ,cc.c12 CsHours" & vbCrLf
        sql &= " from WC2 cc" & vbCrLf
        sql &= " order by cc.orgname" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True

            'PageControler1.SqlString = StrSql
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "orgname"
            PageControler1.Sort = "orgname"
            PageControler1.ControlerLoad()
        End If

    End Sub

    Sub GetSearchStr()
        Session("search") = "yearlist=" & yearlist.SelectedValue
        Session("search") += "&center=" & center.Text
        Session("search") += "&RIDValue=" & RIDValue.Value
        Session("search") += "&start_date=" & start_date.Text
        Session("search") += "&end_date=" & end_date.Text
        Session("search") += "&PageIndex=" & dtPlan.CurrentPageIndex + 1
    End Sub
End Class

