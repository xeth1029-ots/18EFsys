Partial Class SYS_04_010
    Inherits AuthBasePage

    'select * from  sys_budgetclose --依轄區計畫 distid tplanid
    'Dim conn As SqlConnection = DbAccess.GetConnection
    'Dim sql As String = ""

#Region "Sub"
    '載入 DropDownList 內容
    Private Sub ddlList(ByVal strType As String, ByVal ddlObj As DropDownList)
        With ddlObj
            .Items.Clear()
            Select Case strType
                Case "tplan"
                    TIMS.Get_YearTPlan(ddlTPlan, sm.UserInfo.Years, TIMS.cst_NO, objconn)
                Case "month", "day"
                    Dim intCnt As Integer = IIf(strType = "month", 12, 31)
                    For i As Integer = 1 To intCnt
                        .Items.Add(New ListItem(i, i))
                    Next
                    .Items.Insert(0, New ListItem("請選擇", ""))
            End Select
        End With
    End Sub

    '代入資料
    Private Sub loadData(ByVal strDistID As String)
        Dim v_ddlTPlan As String = TIMS.GetListValue(ddlTPlan)
        strDistID = If(strDistID <> "", strDistID, sm.UserInfo.DistID)
        Dim s_pms As New Hashtable From {{"distid", strDistID}, {"tplanid", v_ddlTPlan}}
        Dim sql As String = ""
        sql &= " select sbcid,close1,close2,close3,close4,close5,close6 "
        sql &= " from sys_budgetclose "
        sql &= " where distid= @distid and tplanid= @tplanid"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, s_pms)
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = Nothing
            dr = dt.Rows(0)
            hidSBCID.Value = Convert.ToString(dr("sbcid"))
            txtClose1.Text = Convert.ToString(dr("close1"))
            If Convert.ToString(dr("close2")) <> "" Then
                ddlClose2M.SelectedValue = Split(Convert.ToString(dr("close2")), "/")(0)
                ddlClose2D.SelectedValue = Split(Convert.ToString(dr("close2")), "/")(1)
            End If
            txtClose3.Text = Convert.ToString(dr("close3"))
            txtClose4.Text = Convert.ToString(dr("close4"))
            txtClose5.Text = Convert.ToString(dr("close5"))
            If Convert.ToString(dr("close6")) <> "" Then
                ddlClose6M.SelectedValue = Split(Convert.ToString(dr("close6")), "/")(0)
                ddlClose6D.SelectedValue = Split(Convert.ToString(dr("close6")), "/")(1)
            End If
        Else
            hidSBCID.Value = ""
            txtClose1.Text = ""
            ddlClose2M.SelectedValue = ""
            ddlClose2D.SelectedValue = ""
            txtClose3.Text = ""
            txtClose4.Text = ""
            txtClose5.Text = ""
            ddlClose6M.SelectedValue = ""
            ddlClose6D.SelectedValue = ""
        End If

    End Sub

#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            CCreate1()
        End If
    End Sub

    Sub CCreate1()
        tbList.Visible = False
        ddlList("tplan", ddlTPlan)
        ddlList("month", ddlClose2M)
        ddlList("day", ddlClose2D)
        ddlList("month", ddlClose6M)
        ddlList("day", ddlClose6D)

        '(暫)限定超級管理者可修改系統預設, 以後在看
        trSys.Visible = If(sm.UserInfo.RoleID = "0", True, False)

        btnSave.Attributes.Add("onclick", "return chkSave();")
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        hidSBCID.Value = TIMS.ClearSQM(hidSBCID.Value)
        If hidSBCID.Value = "" Then
            '新增
            Dim iSBCID As Integer = DbAccess.GetNewId(objconn, "SYS_BUDGETCLOSE_SBCID_SEQ,SYS_BUDGETCLOSE,SBCID")
            hidSBCID.Value = iSBCID 'DbAccess.GetId(conn, "SYS_BUDGETCLOSE_SBCID_SEQ")
            Dim pms_i As New Hashtable From {
                {"SBCID", iSBCID},
                {"distid", sm.UserInfo.DistID},
                {"tplanid", ddlTPlan.SelectedValue},
                {"close1", txtClose1.Text},
                {"close2", ddlClose2M.SelectedValue & "/" & ddlClose2D.SelectedValue},
                {"close3", txtClose3.Text},
                {"close4", txtClose4.Text},
                {"close5", txtClose5.Text},
                {"close6", ddlClose6M.SelectedValue & "/" & ddlClose6D.SelectedValue},
                {"modifyacct", sm.UserInfo.UserID}
            }
            Dim sql_i As String = ""
            sql_i = " insert into sys_budgetclose(SBCID,distid,tplanid,close1,close2,close3,close4,close5,close6,modifyacct,modifydate) "
            sql_i &= " values(@SBCID,@distid,@tplanid,@close1,@close2,@close3,@close4,@close5,@close6,@modifyacct,getdate()) "
            DbAccess.ExecuteNonQuery(sql_i, objconn, pms_i)

        Else
            '修改
            Dim pms_u As New Hashtable From {
                {"close1", txtClose1.Text},
                {"close2", ddlClose2M.SelectedValue & "/" & ddlClose2D.SelectedValue},
                {"close3", txtClose3.Text},
                {"close4", txtClose4.Text},
                {"close5", txtClose5.Text},
                {"close6", ddlClose6M.SelectedValue & "/" & ddlClose6D.SelectedValue},
                {"modifyacct", sm.UserInfo.UserID},
                {"sbcid", Val(hidSBCID.Value)}
            }
            Dim sql_u As String = ""
            sql_u &= " update sys_budgetclose set close1= @close1,close2= @close2,close3= @close3,close4= @close4"
            sql_u &= " ,close5= @close5,close6= @close6,modifyacct= @modifyacct,modifydate=getdate() "
            sql_u &= " where sbcid= @sbcid"
            DbAccess.ExecuteNonQuery(sql_u, objconn, pms_u)
        End If

        Common.MessageBox(Me, "儲存成功")
    End Sub

    Private Sub ddlTPlan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlTPlan.SelectedIndexChanged
        If ddlTPlan.SelectedValue <> "" Then
            tbList.Visible = True
            Call loadData("")
        Else
            tbList.Visible = False
        End If
    End Sub

    Private Sub chkSys_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSys.CheckedChanged
        ddlTPlan.SelectedValue = ""
        If chkSys.Checked = True Then
            ddlTPlan.Enabled = False
            tbList.Visible = True
            loadData("000")
        Else
            ddlTPlan.Enabled = True
            tbList.Visible = False
        End If
    End Sub
End Class
