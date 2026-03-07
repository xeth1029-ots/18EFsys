Partial Class OB_01_010
    Inherits AuthBasePage

    'Protected WithEvents PageControler1 As PageControler

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        TIMS.Get_TitleLab(Request("ID"), lblTitle1, lblTitle2)

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Me.ViewState("DataGrid") Is Nothing Then
            Me.ViewState("DataGrid") = "DataGrid1"
            '分頁設定 Start
            PageControler1.PageDataGrid = DataGrid1
            '分頁設定 End
        Else
            Select Case Me.ViewState("DataGrid")
                Case "DataGrid1"
                    PageControler1.PageDataGrid = DataGrid1
                Case Else
                    PageControler1.PageDataGrid = DataGrid2
            End Select
        End If


        If Not IsPostBack Then
            'btnSave.Visible = False
            'ddlyears = TIMS.GetSyear(ddlyears, Year(Now) - 1, Year(Now) + 3)
            ddlyears = TIMS.Get_Years(ddlyears)

            DataGridTable.Visible = False
            panelSch.Visible = True
        End If

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        DataGrid1.Visible = True
        DataGrid2.Visible = False
        Me.ViewState("DataGrid") = "DataGrid1"
        PageControler1.PageDataGrid = DataGrid1

        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確
        If TxtPageSize.Text <> DataGrid2.PageSize Then DataGrid2.PageSize = TxtPageSize.Text

        schOrgName.Text = Trim(schOrgName.Text)
        schComIDNO.Text = Trim(schComIDNO.Text)
        'schAeCount.Text = Trim(schAeCount.Text)

        Query()
    End Sub

    Sub Query(Optional ByVal sql As String = "")

        Dim dt As DataTable
        If sql <> "" Then
        Else
            sql = "" & vbCrLf
            sql += " select " & vbCrLf
            sql += " 	g.years" & vbCrLf
            sql += " , g.DistID " & vbCrLf
            sql += " , g.PlanID " & vbCrLf
            sql += " , g.ComIDNO" & vbCrLf
            sql += " , g.OrgName" & vbCrLf
            sql += " , kp.PlanName" & vbCrLf
            sql += " , idt.Name as DistName" & vbCrLf
            sql += " , sum (aeCnt) as aeCount" & vbCrLf
            sql += " from " & vbCrLf
            sql += "  key_plan kp " & vbCrLf
            sql += " join (" & vbCrLf
            sql += " select " & vbCrLf
            sql += " 	ip.years" & vbCrLf
            sql += " , ip.DistID " & vbCrLf
            sql += " , ip.PlanID " & vbCrLf
            sql += " , ip.TPlanID " & vbCrLf
            sql += " , oo.ComIDNO" & vbCrLf
            sql += " , oo.OrgName" & vbCrLf
            sql += " , case " & vbCrLf
            sql += " 		when ae.OCID is null then 0 else 1 end aeCnt" & vbCrLf
            sql += " from Org_OrgInfo oo " & vbCrLf
            sql += " join Class_ClassInfo cc on cc.ComIDNO =oo.ComIDNO " & vbCrLf
            sql += " join id_Plan ip on ip.PlanID =cc.PlanID" & vbCrLf
            sql += " join Auth_REndClass ae on ae.OCID =cc.OCID  and cc.PlanID=ae.PlanID " & vbCrLf
            sql += " where 1=1 " & vbCrLf
            If schOrgName.Text <> "" Then
                sql += " AND oo.OrgName NOT LIKE '%" & schOrgName.Text & "%'" & vbCrLf
            End If
            If schComIDNO.Text <> "" Then
                sql += " AND cc.ComIDNO= '" & schComIDNO.Text & "'" & vbCrLf
            End If
            If ddlyears.SelectedValue <> "" Then
                sql += " AND ip.years='" & ddlyears.SelectedValue & "'" & vbCrLf
            End If

            sql += " ) g on kp.TPlanID =g.TPlanID" & vbCrLf
            sql += " join ID_District idt on idt.DistID =g.DistID" & vbCrLf
            sql += " group by " & vbCrLf
            sql += " 	g.years" & vbCrLf
            sql += " , g.DistID " & vbCrLf
            sql += " , g.PlanID " & vbCrLf
            sql += " , g.ComIDNO" & vbCrLf
            sql += " , g.OrgName" & vbCrLf
            sql += " , kp.PlanName" & vbCrLf
            sql += " , idt.Name" & vbCrLf

            'If schAeCount.Text <> "" Then
            '    sql += " having sum (aeCnt)>=1" & vbCrLf
            'End If
        End If

        dt = DbAccess.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            RecordCount.Text = dt.Rows.Count

            DataGridTable.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        Else
            DataGridTable.Visible = False
            msg.Text = "查無資料!!"
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim aeCount As LinkButton = e.Item.FindControl("aeCount")
                If drv("aeCount") > 0 Then
                    aeCount.CommandArgument = "o=10"
                    aeCount.CommandArgument += "&Years=" & drv("years")
                    aeCount.CommandArgument += "&DistID=" & drv("DistID")
                    aeCount.CommandArgument += "&PlanID=" & drv("PlanID")
                    aeCount.CommandArgument += "&ComIDNO=" & drv("ComIDNO")
                Else
                    aeCount.CommandName = ""
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "aeCount"
                LODataGrid2(e.CommandArgument)
        End Select
    End Sub

    Sub LODataGrid2(ByVal cmdArg As String)
        DataGrid1.Visible = False
        DataGrid2.Visible = True
        Me.ViewState("DataGrid") = "DataGrid2"
        PageControler1.PageDataGrid = DataGrid2

        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確
        If TxtPageSize.Text <> DataGrid2.PageSize Then DataGrid2.PageSize = TxtPageSize.Text

        Me.ViewState("Years") = TIMS.GetMyValue(cmdArg, "Years")
        Me.ViewState("DistID") = TIMS.GetMyValue(cmdArg, "DistID")
        Me.ViewState("PlanID") = TIMS.GetMyValue(cmdArg, "PlanID")
        Me.ViewState("ComIDNO") = TIMS.GetMyValue(cmdArg, "ComIDNO")

        Dim Sql As String = ""

        Sql = "" & vbCrLf
        Sql += " select " & vbCrLf
        Sql += "  ae.Years" & vbCrLf
        Sql += " , oo.OrgName " & vbCrLf
        Sql += " , kp.PlanName" & vbCrLf
        Sql += " , idt.Name as DistName" & vbCrLf
        Sql += " , cc.ClassCName " & vbCrLf
        Sql += " , ae.Account " & vbCrLf
        Sql += " , ae.FunID" & vbCrLf
        Sql += " , CONVERT(varchar, ae.CreateDate, 111)	as CreateDate" & vbCrLf
        Sql += " , CONVERT(varchar, ae.EndDate, 111)	as EndDate" & vbCrLf
        Sql += " from Auth_RendClass ae" & vbCrLf
        Sql += " join Class_ClassInfo cc on cc.OCID =ae.OCID and cc.PlanID=ae.PlanID " & vbCrLf
        Sql += " join id_Plan ip on ip.PlanID =cc.PlanID" & vbCrLf
        Sql += " JOIN Org_OrgInfo oo on oo.ComIDNO =cc.ComIDNO" & vbCrLf
        Sql += " JOIN Key_Plan kp on kp.TPlanID =ip.TPlanID" & vbCrLf
        Sql += " join ID_District idt on idt.DistID =ip.DistID" & vbCrLf
        Sql += " where 1=1 and ROWNUM < 2001 " & vbCrLf
        If Me.ViewState("Years") <> "" Then
            Sql += " AND ae.Years='" & Me.ViewState("Years") & "' " & vbCrLf
        End If
        If Me.ViewState("DistID") <> "" Then
            Sql += " AND ae.DistID='" & Me.ViewState("DistID") & "' " & vbCrLf
        End If
        If Me.ViewState("PlanID") <> "" Then
            Sql += " AND ae.PlanID='" & Me.ViewState("PlanID") & "' " & vbCrLf
        End If
        If Me.ViewState("ComIDNO") <> "" Then
            Sql += " AND cc.ComIDNO='" & Me.ViewState("ComIDNO") & "' " & vbCrLf
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(Sql)
        If dt.Rows.Count > 0 Then
            RecordCount.Text = dt.Rows.Count

            DataGridTable.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        Else
            '應不可能查無資料
        End If
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Const Cst_FunID As Integer = 7
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                If drv("FunID").ToString <> "" Then
                    'e.Item.Cells(Cst_FunID).Text = drv("FunID").ToString
                    e.Item.Cells(Cst_FunID).Text = GetFunName(drv("FunID").ToString)
                Else
                    e.Item.Cells(Cst_FunID).Text = "未填寫"
                End If
        End Select
    End Sub

    Function GetFunName(ByVal FunIDs As String) As String
        Dim FunNames As String = ""
        Dim sql As String
        sql = " SELECT * FROM ID_Function where FunID IN (" & FunIDs & ")"
        Dim dt As DataTable = DbAccess.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            FunNames = ""
            For i As Integer = 0 To dt.Rows.Count - 1
                If FunNames = "" Then
                    FunNames = dt.Rows(i)("Name").ToString
                Else
                    FunNames += "<br>" & dt.Rows(i)("Name").ToString
                End If
            Next
        End If
        Return FunNames
    End Function

End Class
