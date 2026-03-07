'Imports Microsoft.Web.UI.WebControls
Partial Class MainOrg
    Inherits AuthBasePage

    Dim flag_Login1 As Boolean = False '縣市政府承辦人登入時 為 True 其餘為 False
    Dim rqDistID As String = "" 'TIMS.ClearSQM(Request("DistID"))
    Dim rqTPlanID As String = "" 'TIMS.ClearSQM(Request("TPlanID"))
    Dim rqYEARSTYPE As String = "" 'rqYEARSTYPE = TIMS.ClearSQM(Request("YEARSTYPE"))
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        flag_Login1 = TIMS.Chk_LoginUserType1(Me)  '縣市政府承辦人登入時 為 True 其餘為 False
        rqDistID = TIMS.ClearSQM(Request("DistID"))
        rqTPlanID = TIMS.ClearSQM(Request("TPlanID"))
        rqYEARSTYPE = TIMS.ClearSQM(Request("YEARSTYPE"))
        If rqDistID = "" Then
            Common.MessageBox(Me, "輸入轄區參數有誤!")
            Exit Sub
        End If
        If rqTPlanID = "" Then
            Common.MessageBox(Me, "輸入計畫參數有誤!")
            Exit Sub
        End If
        If Not IsPostBack Then Call CreateItem()
        PlanID.Attributes("onchange") = "if(this.selectedIndex==0){return false;}"
    End Sub

    Sub CreateItem()
        If rqDistID = "" Then
            Common.MessageBox(Me, "輸入轄區參數有誤!")
            Exit Sub
        End If
        If rqTPlanID = "" Then
            Common.MessageBox(Me, "輸入計畫參數有誤!")
            Exit Sub
        End If

        Dim dt As DataTable = Nothing
        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " SELECT CONCAT(a.YEARS,c.Name,b.PlanName,a.Seq) PlanName ,a.PlanID" & vbCrLf
        sSql &= " FROM ID_Plan a" & vbCrLf
        sSql &= " JOIN Key_Plan b ON a.TPlanID = b.TPlanID" & vbCrLf
        sSql &= " JOIN ID_District c ON a.DistID = c.DistID" & vbCrLf
        sSql &= " WHERE a.DistID = '" & rqDistID & "' AND a.TPlanID = '" & rqTPlanID & "'" & vbCrLf
        If rqYEARSTYPE = "3" Then
            Dim vYEARSTYPE3 As String = sm.UserInfo.Years - 3
            sSql &= " AND a.YEARS >= '" & vYEARSTYPE3 & "'" & vbCrLf
        End If
        '縣市政府承辦人登入時 為 True 其餘為 False
        If flag_Login1 Then sSql &= " AND a.PlanID = '" & sm.UserInfo.PlanID & "'" & vbCrLf  '鎖定登入年度計畫。
        sSql &= " ORDER BY 1 DESC" & vbCrLf
        dt = DbAccess.GetDataTable(sSql, objconn)
        With PlanID
            .DataSource = dt
            .DataTextField = "PlanName"
            .DataValueField = "PlanID"
            .DataBind()
            .Items.Insert(0, New ListItem("===請選擇計畫===", ""))
        End With

        Select Case dt.Rows.Count
            Case 0
                PlanTable.Visible = False
                Page.RegisterStartupScript("000", "<script>alert('查無此轄區的計畫');window.close();</script>")
            Case 1
                PlanTable.Visible = False
                PlanID.Items(1).Selected = True
                PlanID_SelectedIndexChanged(PlanID, Nothing)
        End Select
    End Sub

    Private Sub PlanID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlanID.SelectedIndexChanged
        If PlanID.SelectedValue = "" Then Exit Sub '未選擇計畫參數
        If rqDistID = "" Then
            Common.MessageBox(Me, "輸入轄區參數有誤!")
            Exit Sub
        End If
        If rqTPlanID = "" Then
            Common.MessageBox(Me, "輸入計畫參數有誤!")
            Exit Sub
        End If

        PlanIDValue.Value = PlanID.SelectedValue
        TreeView1.Nodes.Clear()

        Dim sql As String = ""
        sql &= " SELECT a.RID ,b.OrgName ,a.Relship ,a.DistID ,a.OrgLevel" & vbCrLf
        sql &= " FROM Auth_Relship a" & vbCrLf
        sql &= " JOIN Org_OrgInfo b ON a.OrgID = b.OrgID" & vbCrLf
        sql &= " WHERE 1=1 AND (1!=1" & vbCrLf
        '縣市政府承辦人登入時 為 True 其餘為 False
        If flag_Login1 Then
            '鎖定登入年度計畫所屬機構。
            sql &= " OR (a.PlanID = '" & PlanID.SelectedValue & "' AND a.Relship LIKE '" & sm.UserInfo.RelShip & "%')" & vbCrLf
        Else
            sql &= " OR a.PlanID = '" & PlanID.SelectedValue & "'" & vbCrLf
        End If
        sql &= " OR (a.PlanID = 0 AND (a.DistID = '000' OR a.DistID = '" & rqDistID & "'))" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ORDER BY a.Relship ,a.OrgLevel ,1" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        'If dt Is Nothing Then Exit Sub
        Select Case rqDistID
            Case "000"
                AddNode(dt, "A", "A")
            Case "001"
                AddNode(dt, "B", "B")
            Case "002"
                AddNode(dt, "C", "C")
            Case "003"
                AddNode(dt, "D", "D")
            Case "004"
                AddNode(dt, "E", "E")
            Case "005"
                AddNode(dt, "F", "F")
            Case "006"
                AddNode(dt, "G", "G")
        End Select

        Call TreeView1.ExpandAll()  '以程式設計方式展開節點
    End Sub

    Sub AddNode(ByVal dt As DataTable, ByVal RID As String, ByVal Relship As String, Optional ByVal ParentNode As TreeNode = Nothing)
        Dim MyNode As TreeNode
        Dim dr As DataRow

        If ParentNode Is Nothing Then
            dr = dt.Select("RID='" & RID & "'")(0)
            MyNode = New TreeNode
            MyNode.Text = Convert.ToString(dr("OrgName"))
            MyNode.NavigateUrl = String.Concat("javascript:MainOrgGetValue('", dr("OrgName"), "','", dr("RID"), "');")
            AddNode(dt, RID, dr("Relship"), MyNode)
            TreeView1.Nodes.Add(MyNode)
        Else
            '依機構名排序。
            For Each dr In dt.Select("Relship like '" & Relship & "%' and RID<>'" & RID & "'", "OrgName")
                If Replace(dr("Relship"), Relship, "").IndexOf("/") = Replace(dr("Relship"), Relship, "").LastIndexOf("/") Then
                    MyNode = New TreeNode
                    MyNode.Text = Convert.ToString(dr("OrgName"))
                    MyNode.NavigateUrl = String.Concat("javascript:MainOrgGetValue('", dr("OrgName"), "','", dr("RID"), "');")
                    'ParentNode.Nodes.Add(MyNode)
                    ParentNode.ChildNodes.Add(MyNode)
                    AddNode(dt, dr("RID"), dr("Relship"), MyNode)
                End If
            Next
        End If
    End Sub
End Class