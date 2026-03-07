'Imports Microsoft.Web.UI.WebControls
Partial Class LevOrg2
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        If Not IsPostBack Then
            CCerate1()
        End If
    End Sub

    Sub CCerate1()
        TreeView1.Nodes.Clear()

        Dim ReqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim ReqDistID As String = TIMS.ClearSQM(Request("DistID"))

        If ReqPlanID = "" OrElse ReqDistID = "" Then Return
        If Not TIMS.IsNumberStr(ReqPlanID) Then Return
        If Not TIMS.IsNumberStr(ReqDistID) Then Return

        Dim sql As String = $"
SELECT a.RID ,b.OrgName ,a.Relship ,a.DistID
FROM (SELECT RID,Relship,DistID,OrgID,PlanID FROM AUTH_RELSHIP WHERE PlanID={ReqPlanID} OR (PlanID=0 AND (DistID='000' OR DistID='{ReqDistID}'))) a
JOIN Org_OrgInfo b ON a.OrgID=b.OrgID
"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        Dim ff As String = "DistID='" & ReqDistID & "'"
        If dt.Select(ff).Length = 1 Then
            Dim dr As DataRow = dt.Select(ff)(0)
            AddNode(dt, dr("RID"), dr("RID"))
        End If


#Region "(No Use)"

        'sql="SELECT RID,OrgID,DistID FROM AUTH_RELSHIP WHERE DistID=@DistID and PlanID =0 order by DistID"
        'Call TIMS.OpenDbConn(objconn)
        'Dim dtDist As New DataTable
        'Dim oCmd As New SqlCommand(sql, objconn)
        'With oCmd
        '    .Parameters.Clear()
        '    .Parameters.Add("DistID", SqlDbType.VarChar).Value=Convert.ToString(Request("DistID"))
        '    dtDist.Load(.ExecuteReader())
        'End With
        'If dtDist.Rows.Count > 0 Then
        '    Dim dr As DataRow=dtDist.Rows(0)
        '    AddNode(dt, dr("RID"), dr("RID"))
        'End If

        'Select Case Convert.ToString(Request("DistID"))
        '    Case "000"
        '        AddNode(dt, "A", "A")
        '    Case "001"
        '        AddNode(dt, "B", "B")
        '    Case "002"
        '        AddNode(dt, "C", "C")
        '    Case "003"
        '        AddNode(dt, "D", "D")
        '    Case "004"
        '        AddNode(dt, "E", "E")
        '    Case "005"
        '        AddNode(dt, "F", "F")
        '    Case "006"
        '        AddNode(dt, "G", "G")
        'End Select

#End Region
    End Sub

    Sub AddNode(ByVal dt As DataTable, ByVal RID As String, ByVal Relship As String, Optional ByVal ParentNode As TreeNode = Nothing)
        Dim MyNode As TreeNode
        Dim dr As DataRow
        If ParentNode Is Nothing Then
            dr = dt.Select("RID='" & RID & "'")(0)
            MyNode = New TreeNode
            MyNode.Text = dr("OrgName")
            MyNode.NavigateUrl = "javascript:GetValue('" & dr("OrgName") & "','" & dr("RID") & "')"
            AddNode(dt, RID, dr("Relship"), MyNode)
            TreeView1.Nodes.Add(MyNode)
        Else
            For Each dr In dt.Select("Relship like '" & Relship & "%' and RID<>'" & RID & "'", "OrgName")
                If Replace(dr("Relship"), Relship, "").IndexOf("/") = Replace(dr("Relship"), Relship, "").LastIndexOf("/") Then
                    MyNode = New TreeNode
                    MyNode.Text = dr("OrgName")
                    MyNode.NavigateUrl = "javascript:GetValue('" & dr("OrgName") & "','" & dr("RID") & "')"
                    'ParentNode.Nodes.Add(MyNode)
                    ParentNode.ChildNodes.Add(MyNode)
                    AddNode(dt, dr("RID"), dr("Relship"), MyNode)
                End If
            Next
        End If
    End Sub
End Class