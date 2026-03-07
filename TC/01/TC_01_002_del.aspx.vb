Partial Class TC_01_002_del
    Inherits AuthBasePage

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        'DbAccess.GetConnection(objconn)

        If Not IsPostBack Then
            'If Not Session("OrgSearchStr") Is Nothing Then Me.ViewState("OrgSearchStr") = Session("OrgSearchStr")
            'Session("OrgSearchStr") = Nothing
        End If

        Call sUtl_Delete1()

        'If Not Me.ViewState("OrgSearchStr") Is Nothing Then
        '    If Session("OrgSearchStr") Is Nothing Then Session("OrgSearchStr") = Me.ViewState("OrgSearchStr")
        'End If
    End Sub

    Sub sUtl_Delete1()
        Dim Re_orgid As String = Request("orgid")
        Dim Re_rid As String = Request("rid")
        Dim Re_planid As String = Request("planid")
        Dim rqID As String = Request("ID")
        Re_orgid = TIMS.ClearSQM(Re_orgid)
        Re_rid = TIMS.ClearSQM(Re_rid)
        Re_planid = TIMS.ClearSQM(Re_planid)
        rqID = TIMS.ClearSQM(rqID)

        'Dim sqlstr_del, check_rsid, rsid_str As String
        Dim check_rsid As String = ""
        Dim rsid_str As String = ""
        '判斷RSID為何
        check_rsid = " SELECT rsid FROM Auth_Relship WHERE rid = '" & Re_rid & "' AND orgid = '" & Re_orgid & "' "
        rsid_str = Convert.ToString(DbAccess.ExecuteScalar(check_rsid, objConn))
        If rsid_str = "" Then Exit Sub

        'Dim str, planstr, TPlanname, org_str, orgname, org_comidno, tplanid, tplan_str As String
        Dim planstr As String = ""
        Dim tplanid As String = ""
        Dim tplan_str As String = ""
        Dim TPlanname As String = ""
        Dim org_str As String = ""
        Dim orgname As String = ""
        Dim org_comidno As String = ""
        Dim dr As DataRow = Nothing
        planstr = " SELECT TPlanID FROM ID_Plan WHERE PlanID = '" & Re_planid & "' "
        dr = DbAccess.GetOneRow(planstr, objConn)
        tplanid = dr("TPlanID")
        tplan_str = " SELECT PlanName FROM key_plan WHERE TPlanID = '" & tplanid & "' "
        dr = DbAccess.GetOneRow(tplan_str, objConn)
        TPlanname = dr("PlanName")
        org_str = " SELECT * FROM Org_OrgInfo WHERE orgid = '" & Re_orgid & "' "
        dr = DbAccess.GetOneRow(org_str, objConn)
        orgname = dr("orgname")
        org_comidno = dr("ComIDNO")
        '刪除[訓練計畫名稱]-[機構名稱]-[統編]
        Dim str As String = "刪除[" & TPlanname & "]-[" & orgname & "]-[" & org_comidno & "]"
        Dim objTrans As SqlTransaction = Nothing
        Dim parms As Hashtable = New Hashtable()
        rqID = "57" 'for 訓練機構管理用

        Try
            If objConn.State = ConnectionState.Closed Then objConn.Open()
            objTrans = DbAccess.BeginTrans(objConn)
            Dim check_org As String = " SELECT * FROM Auth_Relship WHERE orgid = " & Re_orgid

            Dim sqlstr_del As String = ""
            '2018 delete auth_accrwplan 改為參數式
            sqlstr_del = " DELETE Auth_AccRWPlan WHERE rid = @rid AND PlanID = @PlanID "
            parms.Add("rid", Re_rid)
            parms.Add("PlanID", Re_planid)
            DbAccess.ExecuteNonQuery(sqlstr_del, objTrans, parms)

            If DbAccess.GetCount(check_org, objTrans) > 1 Then '有共用機構,刪除Auth_Relship,Org_OrgPlanInfo
                TIMS.InsertDelLog(sm.UserInfo.UserID, rqID, sm.UserInfo.DistID, str, Re_orgid, Re_rid, Re_planid, org_comidno)
                '2018 delete auth_relship 改為參數式
                sqlstr_del = "DELETE Auth_Relship where rid=@rid and PlanID=@PlanID and orgid=@orgid "
                parms.Clear()
                parms.Add("rid", Re_rid)
                parms.Add("PlanID", Re_planid)
                parms.Add("orgid", Re_orgid)
                DbAccess.ExecuteNonQuery(sqlstr_del, objTrans, parms)
                '2018 delete org_orgplaninfo 改為參數式
                sqlstr_del = "DELETE Org_OrgPlanInfo where rsid=@rsid "
                parms.Clear()
                parms.Add("rsid", rsid_str)
                DbAccess.ExecuteNonQuery(sqlstr_del, objTrans, parms)
            Else '沒有共用,全部刪除
                TIMS.InsertDelLog(sm.UserInfo.UserID, rqID, sm.UserInfo.DistID, str, Re_orgid, Re_rid, Re_planid, org_comidno)
                '2018 delete auth_relship 改為參數式
                sqlstr_del = "DELETE Auth_Relship  where orgid=@orgid"
                parms.Clear()
                parms.Add("orgid", Re_orgid)
                DbAccess.ExecuteNonQuery(sqlstr_del, objTrans, parms)
                '2018 delete org_orginfo 改為參數式
                sqlstr_del = "DELETE Org_orginfo  where orgid=@orgid "
                parms.Clear()
                parms.Add("orgid", Re_orgid)
                DbAccess.ExecuteNonQuery(sqlstr_del, objTrans, parms)
                '2018 delete org_orgplaninfo 改為參數式
                sqlstr_del = "DELETE Org_OrgPlanInfo where rsid=@rsid"
                parms.Clear()
                parms.Add("rsid", rsid_str)
                DbAccess.ExecuteNonQuery(sqlstr_del, objTrans, parms)
            End If
            DbAccess.CommitTrans(objTrans)
            Common.RespWrite(Me, "<script>alert('刪除成功!!');</script>")
            Common.RespWrite(Me, "<script>location.href='../01/TC_01_002.aspx?ProcessType=del&ID=" & rqID & "'</script>")
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            Common.MessageBox(Page, "訓練機構刪除失敗!!")
            Throw ex
        End Try
    End Sub
End Class