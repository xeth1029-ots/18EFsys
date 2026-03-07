Partial Class TC_01_005_Classid
    Inherits AuthBasePage

    Dim dt As DataTable
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
        '分頁設定 Start
        PageControler1.PageDataGrid = DG_Class
        '分頁設定 End
        Dim rqID_ClassYears As String = TIMS.ClearSQM(Request("ID_ClassYears"))
        Dim rqTPlanID As String = TIMS.ClearSQM(Request("tplanid"))
        Dim rqProcessType As String = TIMS.ClearSQM(Request("ProcessType"))
        'Common.RespWrite(Me, ProcessType)

        If Not Me.IsPostBack Then
            'Dim sql_TPlan As String = "select TPlanID,PlanName from Key_Plan "
            '改成依登入帳號,登入的年度顯示所賦予的計畫
            Call Utl_SetTPlanList()

            ddlYears = TIMS.GetSyear(ddlYears, 2009, Year(Now) + 1, False)
            If rqID_ClassYears <> "" Then
                Common.SetListItem(ddlYears, rqID_ClassYears)
            Else
                'Common.SetListItem(ddlYears, 2009)
                Common.SetListItem(ddlYears, sm.UserInfo.Years)
            End If
            TIMS.Tooltip(bt_search, "查詢該年度班別代碼(2009年以前為舊資料)", True)

            Select Case rqProcessType
                Case "Insert"
                    'Dim sql_PlanID As String = "select TPlanID from ID_Plan where PlanID='" & PlanID & "'"
                    'TPlanID = Convert.ToString(DbAccess.ExecuteScalar(sql_PlanID, objconn))
                    Common.SetListItem(TPlan_List, rqTPlanID)
                    'TPlan_List.SelectedValue = TPlanID
                Case "Update"
                    Common.SetListItem(TPlan_List, rqTPlanID)
                    'TPlan_List.SelectedValue = PlanID '(TPlanID)
            End Select

            If rqTPlanID <> "" Then
                'bt_search_Click(sender, e)
                Call Search1()
            End If
        End If

    End Sub

    '改成依登入帳號,登入的年度顯示所賦予的計畫
    Sub Utl_SetTPlanList()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        'Dim sql_TPlan As String = "select TPlanID,PlanName from Key_Plan "
        Dim sql_TPlan As String = ""
        sql_TPlan = "" & vbCrLf
        sql_TPlan += " SELECT DISTINCT kp.Tplanid,kp.planname" & vbCrLf
        sql_TPlan += " FROM Key_Plan kp" & vbCrLf
        sql_TPlan += " JOIN id_plan ip on kp.tplanid = ip.tplanid" & vbCrLf
        sql_TPlan += " JOIN auth_accrwplan an on ip.planid = an.planid" & vbCrLf
        sql_TPlan += " where an.account ='" & sm.UserInfo.UserID & "' " & vbCrLf
        sql_TPlan += " and ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
        dt = DbAccess.GetDataTable(sql_TPlan, objconn)
        TPlan_List.DataSource = dt
        TPlan_List.DataTextField = "PlanName"
        TPlan_List.DataValueField = "TPlanID"
        TPlan_List.DataBind()
        TPlan_List.Items.Insert(0, New ListItem("===全部===", ""))
    End Sub

    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select a.CLSID" & vbCrLf
        sql &= " ,a.ClassID" & vbCrLf
        sql &= " ,a.CLassName" & vbCrLf
        sql &= " ,dbo.NVL(CONVERT(varchar, a.Years),'2009(舊資料)') Years " & vbCrLf
        sql &= " from ID_Class a " & vbCrLf
        sql &= " join Key_Plan b on b.TPlanID=a.TPlanID  " & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.DistID='" & sm.UserInfo.DistID & "' " & vbCrLf

        '2010後年度應選擇(2009年(含)以前為null)
        If ddlYears.SelectedValue <> "" Then
            If ddlYears.SelectedValue < 2010 Then
                sql &= " and a.Years is null " & vbCrLf
            Else
                '2010課程資料採年度設定 by AMU (修正 年度 BUG)
                sql &= " and a.Years='" & ddlYears.SelectedValue & "'" & vbCrLf
            End If
        Else
            If sm.UserInfo.Years < 2010 Then
                sql &= " and a.Years is null " & vbCrLf
            Else
                '2010課程資料採年度設定 by AMU (修正 年度 BUG)
                sql &= " and a.Years='" & sm.UserInfo.Years & "'" & vbCrLf
            End If
        End If

        If TPlan_List.SelectedIndex <> 0 Then
            sql &= "and a.TPlanID='" & TPlan_List.SelectedValue & "'"
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        msg.Text = "查無資料!!"
        DG_Class.Visible = False
        PageControler1.Visible = False
        submit.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DG_Class.Visible = True
            PageControler1.Visible = True
            submit.Visible = True

            'PageControler1.SqlPrimaryKeyDataCreate(sql_class, "CLSID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "CLSID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call Search1()
    End Sub

    Private Sub DG_Class_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Class.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim myradio As HtmlInputRadioButton = e.Item.FindControl("myradio")
                'set_class(" & Container.DataItem("CLSID") & ",""" & Container.DataItem("ClassID") & """,""" & Container.DataItem("CLassName") & """)

                Dim strClick1 As String = "set_class('" & drv("CLSID") & "','" & drv("ClassID") & "','" & drv("CLassName") & "');"
                myradio.Attributes("onclick") = strClick1
        End Select
    End Sub


End Class
