Partial Class CP_04_001_01
    Inherits AuthBasePage

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
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            If Not Session("itemstr") Is Nothing Then Me.ViewState("itemstr") = Session("itemstr")
            If Not Session("itemplan") Is Nothing Then Me.ViewState("itemplan") = Session("itemplan")
            If Not Session("itemcity") Is Nothing Then Me.ViewState("itemcity") = Session("itemcity")
            Session("itemstr") = Nothing
            Session("itemplan") = Nothing
            Session("itemcity") = Nothing
            create()
        End If

        '回上一頁
        Me.Button2.Attributes.Add("onclick", "location.href='CP_04_001.aspx';return false;")
    End Sub

    Sub create()
        Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr As String
        Dim yearlist As String = Request("yearlist")
        Dim itemstr As String = Me.ViewState("itemstr")
        Dim itemplan As String = Me.ViewState("itemplan")
        Dim itemcity As String = Me.ViewState("itemcity")
        'Dim SelectPlanTimes As Integer = Request("SelectPlanTimes")
        'Dim TPlanStr, DistStr, YearsStr, CityStr As String
        'Dim i As Integer

        '以轄區、訓練計畫做排序
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " SELECT a.RID" & vbCrLf
        sqlstr += " ,a.OrgID" & vbCrLf
        sqlstr += " ,b.OrgName" & vbCrLf
        sqlstr += " ,b.ComIDNO" & vbCrLf
        sqlstr += " ,c.Name DistName" & vbCrLf
        sqlstr += " ,c.DistID" & vbCrLf
        sqlstr += " ,e.PlanName" & vbCrLf
        sqlstr += " ,(iz.ZipName+ replace(replace(h.Address,iz.CTName,''),iz.ZName,'')) Address" & vbCrLf
        sqlstr += " ,h.ContactName" & vbCrLf
        sqlstr += " ,h.Phone" & vbCrLf
        sqlstr += " ,h.ContactEmail" & vbCrLf
        sqlstr += " FROM (SELECT RID,RSID,OrgID,DISTID,PLANID FROM Auth_Relship" & vbCrLf
        sqlstr += " WHERE (PlanID IN (SELECT PlanID FROM ID_Plan WHERE 1=1" & vbCrLf
        '選擇訓練計畫
        If itemplan <> "" Then
            sqlstr &= " and TPlanID IN (" & itemplan & ")" & vbCrLf
        End If
        '選擇轄區
        If itemstr <> "" Then
            sqlstr &= " and DistID IN (" & itemstr & ")" & vbCrLf
        End If
        '選擇年度
        If yearlist <> "" Then
            sqlstr &= " and Years='" & Trim(yearlist) & "'" & vbCrLf
        End If
        sqlstr += " ))" & vbCrLf
        sqlstr += " OR (PlanID=0" & vbCrLf
        '選擇轄區
        If itemstr <> "" Then
            sqlstr &= " and DistID IN (" & itemstr & ")" & vbCrLf
        End If
        sqlstr += " )" & vbCrLf
        sqlstr += " ) a" & vbCrLf
        sqlstr += " JOIN Org_OrgInfo b ON a.OrgID = b.OrgID" & vbCrLf
        sqlstr += " JOIN Org_OrgPlanInfo h ON a.RSID = h.RSID" & vbCrLf
        sqlstr += " JOIN view_zipname iz ON iz.ZipCode = h.ZipCode" & vbCrLf
        sqlstr += " JOIN ID_District c ON a.DistID = c.DistID" & vbCrLf
        sqlstr += " LEFT JOIN ID_Plan d ON a.PlanID = d.PlanID" & vbCrLf
        sqlstr += " LEFT JOIN Key_Plan e ON d.TPlanID = e.TPlanID" & vbCrLf
        sqlstr &= " WHERE 1=1" & vbCrLf
        '選擇縣市
        If itemcity <> "" Then
            sqlstr &= " and iz.CTID IN (" & itemcity & ")" & vbCrLf
        End If
        sqlstr += " Order By a.DistID,e.PlanName,a.RID"
        '以轄區、訓練計畫做排序
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        Me.NoData.Text = "<font color=red>查無資料</font>"
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            Me.NoData.Text = ""
            Me.DataGrid1.Visible = True
            Me.PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "RID"
            PageControler1.Sort = "DistID,PlanName"
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            '序號
            e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
        End If
    End Sub
End Class
