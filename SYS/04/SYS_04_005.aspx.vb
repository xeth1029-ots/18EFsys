Partial Class SYS_04_005
    Inherits AuthBasePage

    Dim dtPlanCostCate As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線

        If Not IsPostBack Then
            create()
        End If

    End Sub

    Sub create()
        Dim sql As String = ""
        Dim SearchStr As String = ""
        Dim dt As DataTable

        sql = "SELECT * FROM PLAN_COSTCATE ORDER BY TPLANID"
        dtPlanCostCate = DbAccess.GetDataTable(sql, objconn)

        sql = "SELECT TPlanID,PlanName FROM Key_Plan WHERE 1=1" & SearchStr
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Style.Item("display") = "none"
        If dt.Rows.Count > 0 Then
            DataGridTable.Style.Item("display") = ""

            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "TPlanID"
            DataGrid1.DataBind()
        End If
    End Sub


    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            'Dim dr As DataRow
            Dim drv As DataRowView = e.Item.DataItem
            Dim CateNo As RadioButtonList = e.Item.FindControl("CateNo")


            If dtPlanCostCate.Select("TPlanID='" & drv("TPlanID") & "'").Length <> 0 Then
                Common.SetListItem(CateNo, dtPlanCostCate.Select("TPlanID='" & drv("TPlanID") & "'")(0)("CateNo"))
            End If
        End If
    End Sub

    Sub SaveData1()
        'Dim da As SqlDataAdapter = Nothing
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sql As String
        'Dim Filter As String
        'dt = DbAccess.GetDataTable(sql, da, objconn)
        For Each item As DataGridItem In DataGrid1.Items
            Dim CateNo As RadioButtonList = item.FindControl("CateNo")
            If Not CateNo.SelectedItem Is Nothing Then
                Dim vTPlanID As String = TIMS.ClearSQM(DataGrid1.DataKeys(item.ItemIndex))
                Dim vCateNo As String = TIMS.ClearSQM(CateNo.SelectedValue)
                Dim s_Sql As String = "SELECT * FROM PLAN_COSTCATE WHERE TPLANID=@TPLANID"
                Dim pParms As New Hashtable
                pParms.Add("TPLANID", vTPlanID)
                Dim dt1 As DataTable = DbAccess.GetDataTable(s_Sql, objconn, pParms)
                If dt1.Rows.Count = 0 Then
                    Dim iPCCID As Integer = DbAccess.GetNewId(objconn, "PLAN_COSTCATE_PCCID_SEQ,PLAN_COSTCATE,PCCID")
                    Dim i_Sql As String
                    i_Sql = ""
                    i_Sql &= " INSERT INTO PLAN_COSTCATE(PCCID,TPLANID,CATENO,MODIFYACCT,MODIFYDATE)"
                    i_Sql &= " VALUES( @PCCID,@TPLANID,@CATENO,@MODIFYACCT,GETDATE())"
                    Dim iParms As New Hashtable
                    iParms.Add("PCCID", iPCCID)
                    iParms.Add("TPLANID", vTPlanID)
                    iParms.Add("CATENO", vCateNo)
                    iParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    DbAccess.ExecuteNonQuery(i_Sql, objconn, iParms)

                Else
                    Dim s2_Sql As String = "SELECT * FROM PLAN_COSTCATE WHERE TPLANID=@TPLANID AND CATENO=@CATENO"
                    Dim p2Parms As New Hashtable
                    p2Parms.Add("TPLANID", vTPlanID)
                    p2Parms.Add("CATENO", vCateNo)
                    Dim dt2 As DataTable = DbAccess.GetDataTable(s2_Sql, objconn, p2Parms)
                    If dt2.Rows.Count = 0 Then
                        Dim u_Sql As String = ""
                        u_Sql = ""
                        u_Sql &= " UPDATE PLAN_COSTCATE"
                        u_Sql &= " SET CATENO=@CATENO,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()"
                        u_Sql &= " WHERE TPlanID=@TPlanID"
                        Dim uParms As New Hashtable
                        uParms.Add("TPLANID", vTPlanID)
                        uParms.Add("CATENO", vCateNo)
                        uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                        DbAccess.ExecuteNonQuery(u_Sql, objconn, uParms)
                    End If

                End If
            End If
        Next
        'DbAccess.UpdateDataTable(dt, da)
        Common.MessageBox(Me, "儲存成功!")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call SaveData1()
    End Sub
End Class
