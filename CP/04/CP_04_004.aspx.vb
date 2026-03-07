Partial Class CP_04_004
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        If Not Page.IsPostBack Then
            Dim dt As DataTable
            Dim dr As DataRow
            Dim sqlstr As String

            Me.PageControler1.Visible = False

            yearlist = TIMS.GetSyear(yearlist)
            sqlstr = "select PlanID,Years from  ID_Plan where PlanID= '" & sm.UserInfo.PlanID & "'"
            dr = DbAccess.GetOneRow(sqlstr, objconn)
            Common.SetListItem(yearlist, dr("Years"))
            'Me.yearlist.Items.Insert(0, New ListItem("===請選擇===", ""))

            sqlstr = "SELECT Name FROM ID_District"
            dt = DbAccess.GetDataTable(sqlstr, objconn)
            Me.DistrictList.DataSource = dt
            Me.DistrictList.DataTextField = "Name"
            Me.DistrictList.DataValueField = "Name"
            Me.DistrictList.DataBind()
            Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

            Me.AllKindEngage.Checked = True

        End If

        PageControler1.PageDataGrid = DataGrid1
    End Sub

#Region "NO USE"
    'Private Sub AllDistrictList_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    '將AllDistrictList的選項都選取
    '    Dim objitem As ListItem
    '    If Me.AllDistrictList.Checked = True Then
    '        For Each objitem In DistrictList.Items
    '            objitem.Selected = True
    '        Next
    '    End If
    'End Sub

#End Region

    '選擇全部轄區
    Private Sub DistrictList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DistrictList.SelectedIndexChanged
        Dim i As Short
        If Me.DistrictList.Items(0).Selected = True Then
            For i = 0 To Me.DistrictList.Items.Count - 1
                Me.DistrictList.Items(i).Selected = True
            Next
        End If
    End Sub


    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Dim dt As DataTable
        'Dim dr As DataRow
        Dim sqlstr As String = ""

        '「*=」指要把第一個資料表裡的所有資料列都含括到查詢結果內 
        'https://msdn.microsoft.com/zh-tw/library/ee240720(v=sql.120).aspx
        sqlstr = ""
        sqlstr &= " select b.name"
        sqlstr &= " ,COUNT(a.TechID) Teacher_Count "
        sqlstr += " from Auth_Relship e"
        sqlstr &= " join Teach_TeacherInfo a on e.RID=a.RID"
        sqlstr &= " join ID_District b on e.DistID = b.DistID"
        sqlstr &= " join ID_Plan c on e.PlanID=c.PlanID"
        'sqlstr += " where  e.RID *=  a.RID  AND e.DistID = b.DistID AND   "

        'select b.name, COUNT(a.TechID) AS Teacher_Count 
        'from Auth_Relship e left join Teach_TeacherInfo a on  e.RID =  a.RID JOIN ID_District b on e.DistID=b.DistID 
        'inner join ID_Plan c on e.PlanID=c.PlanID
        'GROUP BY   b.Name,b.DistID 
        'order by b.DistID

        '選擇年度
        If Me.yearlist.SelectedIndex <> 0 Then
            sqlstr += " and d.Years='" & Me.yearlist.SelectedValue & "'"
        End If

        '選擇轄區
        'Dim objitem As ListItem
        Dim itemstr As String = ""
        For Each objitem As ListItem In Me.DistrictList.Items
            If objitem.Selected Then
                If itemstr <> "" Then itemstr &= ","
                itemstr &= "'" & objitem.Value & "'"
            End If
        Next
        If itemstr <> "" Then
            sqlstr += " and b.Name IN (" & itemstr & ")"
        End If
        sqlstr += " GROUP BY  b.Name, b.DistID "
        'sqlstr += " order by b.DistID "

        'sqlstr += " order by b.Name, c.PlanName, a.OrgName"
        'dim sql As String

        dt = DbAccess.GetDataTable(sqlstr, objconn)
        Me.NoData.Text = "<font color=red>查無資料</font>"
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            Me.NoData.Text = ""
            Me.DataGrid1.Visible = True
            Me.PageControler1.Visible = True
            'PageControler1.SqlPrimaryKeyDataCreate(sqlstr, "RID", "DistID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "RID"
            PageControler1.Sort = "DistID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub bt_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reset.Click
        'Reset
        Me.yearlist.SelectedIndex = 0
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False

        Dim i As Short
        For i = 0 To Me.DistrictList.Items.Count - 1
            Me.DistrictList.Items(i).Selected = False
        Next

        Me.AllKindEngage.Checked = True
        Me.KindEngage1.Checked = False
        Me.KindEngage2.Checked = False

    End Sub
End Class
