Partial Class SYS_01_004_view
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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Hidacc.Value = ""
            Call get_years()
            Common.SetListItem(ddl_years, "") '預設帶出所有年度計畫
            'ddl_years.SelectedIndex = 0 '預設帶出所有年度計畫

            If Session("SYS_01_004_acc") <> "" Then
                If Hidacc.Value <> Session("SYS_01_004_acc") Then
                    Hidacc.Value = Session("SYS_01_004_acc")
                End If
            End If
            Call loaddata(Hidacc.Value, ddl_years.SelectedValue)
            Session("SYS_01_004_acc") = Nothing
            Session("SYS_01_004_years") = Nothing
        End If
    End Sub

    '取得下拉選單(年度)
    Private Sub get_years()
        Dim dt As New DataTable
        'Dim da As SqlDataAdapter = TIMS.GetOneDA()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select distinct a.years " & vbCrLf
        sql += " from id_plan a" & vbCrLf
        sql += " join key_plan b on a.tplanid=b.tplanid" & vbCrLf
        sql += " where a.years is not null" & vbCrLf
        sql += " order by a.years" & vbCrLf
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        'TIMS.Fill(Sql, da, dt)
        If dt.Rows.Count > 0 Then
            ddl_years.DataTextField = "years"
            ddl_years.DataValueField = "years"
            ddl_years.DataSource = dt
            ddl_years.DataBind()
        End If
        ddl_years.Items.Insert(0, New ListItem("不區分", ""))
    End Sub

    Sub loaddata(ByVal acct As String, ByVal years As String)
        If acct = "" Then Exit Sub
        If years = "" Then Exit Sub

        Dim dt As New DataTable
        'Dim da As SqlDataAdapter = TIMS.GetOneDA()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT ac.name as accName, a.RID,b.TPlanID, b.PlanID,a.CreateByAcc" & vbCrLf
        sql += " ,CASE WHEN o2.OrgName is null then b.Years+e.Name+f.PlanName+b.Seq" & vbCrLf
        sql += " ELSE b.Years+e.Name+f.PlanName+b.Seq +'　('+o2.OrgName+')' END AS PlanName" & vbCrLf
        sql += " ,d.OrgID, d.OrgName, ac.isused, c.DistID, c.OrgLevel" & vbCrLf
        sql += " ,a2.RID CRID,o2.OrgID COrgID, o2.OrgName COrgName" & vbCrLf
        sql += " FROM Auth_AccRWPlan a" & vbCrLf
        sql += " JOIN ID_Plan b on a.PlanID=b.PlanID and a.account= @acct" & vbCrLf
        sql += " JOIN (select rr.DistID,rr.OrgLevel,rr.OrgID,rr.RID,CASE" & vbCrLf
        sql += "    WHEN Len(rr.Relship)-4 >0" & vbCrLf
        sql += "    THEN replace(replace( dbo.SUBSTR(rr.Relship,5,Len(rr.Relship)-4),rr.RID,''),'/','')" & vbCrLf
        sql += "    END AS CRID from Auth_Relship rr) c on a.RID=c.RID" & vbCrLf
        sql += " JOIN Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
        sql += " JOIN ID_District e on b.DistID=e.DistID" & vbCrLf
        sql += " JOIN Key_Plan f on b.TPlanID=f.TPlanID" & vbCrLf
        sql += " JOIN auth_account ac on a.account=ac.account" & vbCrLf
        sql += " left join Auth_Relship a2 on a2.RID=c.CRID" & vbCrLf
        sql += " left join Org_OrgInfo o2 on a2.OrgID=o2.OrgID" & vbCrLf
        If years <> "" Then
            sql += " where  b.Years= @years "
        End If
        sql += " order by  b.Years "

        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("acct", SqlDbType.VarChar).Value = acct
            If years <> "" Then
                .Parameters.Add("years", SqlDbType.VarChar).Value = years
            End If
            dt.Load(.ExecuteReader())
        End With

        DataGrid1.Visible = False
        msg.Text = years & "年度 查無賦予計畫資料!"
        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            msg.Text = ""

            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub ddl_years_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_years.SelectedIndexChanged
        If Hidacc.Value <> "" Then
            Call loaddata(Hidacc.Value, ddl_years.SelectedValue)
        End If
    End Sub

End Class
