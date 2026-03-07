Partial Class SYS_06_007
    Inherits AuthBasePage

    Const cst_text_color As String = "blue"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        'TIMS.TestDbConn(Me, objconn)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End
        'PageControler1.PageDataGrid = DataGrid1

        If Not Page.IsPostBack Then
            'Call Show_KeyPlan()
            Call TIMS.sUtl_ShowTPlanID("1999", planlist, objconn)

            dopNextNode.Visible = False
            panEdit.Visible = False
        Else
            If dopParentNode.SelectedValue <> "" Then
                createTable("0", dopParentNode.SelectedValue, "")
            End If
            If dopNextNode.SelectedValue <> "" Then
                createTable("1", dopParentNode.SelectedValue, dopNextNode.SelectedValue)
            End If
        End If
    End Sub

    Private Sub createTable(ByVal NodeKey As String, ByVal Kind As String, ByVal ParentID As String)
        Dim dt As DataTable
        Dim vsDistID As String = "" & Convert.ToString(sm.UserInfo.DistID)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT  " & vbCrLf
        sql += " case dbo.NVL(x.Sort,0) WHEN 0 THEN p.Sort else x.Sort end newSort " & vbCrLf
        sql += " ,p.FunID" & vbCrLf ' /*PK*/
        sql += " ,p.Name" & vbCrLf
        sql += " ,p.SPage" & vbCrLf
        sql += " ,p.Kind" & vbCrLf
        sql += " ,p.Levels" & vbCrLf
        sql += " ,p.Parent" & vbCrLf
        sql += " ,p.Valid" & vbCrLf
        sql += " ,p.Sort" & vbCrLf
        sql += " FROM ID_Function p" & vbCrLf
        sql += " left join Plan_Func x on p.FunID=x.FunID AND x.TPlanID='" & planlist.SelectedValue & "'  and x.DistID='" & Convert.ToString(sm.UserInfo.DistID) & "' " & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND exists (" & vbCrLf
        sql += "     select 'x' from auth_groupfun c where c.funid=p.funid" & vbCrLf
        sql += "     and exists (" & vbCrLf
        sql += "         select 'x' from auth_planfunction c2 where c2.tplanid=" & planlist.SelectedValue & "  and c2.funid=c.funid" & vbCrLf
        sql += "     ))" & vbCrLf
        sql += " AND p.Kind='" & Kind & "' " & vbCrLf
        sql += " AND p.Valid='Y'" & vbCrLf
        'sql += " AND p.Levels='" & NodeKey & "'" & vbCrLf
        'If ParentID <> "" Then
        '    sql += " AND p.Parent = '" & ParentID & "'" & vbCrLf
        'End If
        'sql += " order by p.Kind, p.Levels, p.newSort, p.Parent"
        sql += " order by p.Kind, p.Levels, 1 , p.Parent"
        dt = DbAccess.GetDataTable(sql, objconn)

        tbDataTable.Rows.Clear()
        If dt.Rows.Count > 0 Then
            Dim MyCell As TableCell
            Dim MyRow As New TableRow
            Dim MyHiddenBox As HtmlInputHidden 'HiddenField
            Dim MyUpBtn As HtmlInputButton

            Dim sSort As String = " Kind, Levels, newSort, Parent"
            Dim sFilter As String = ""
            sFilter = ""
            sFilter &= " Levels='" & NodeKey & "' " 'Level
            If ParentID <> "" Then
                sFilter &= " AND [Parent] = '" & ParentID & "'" & vbCrLf
            End If
            '產生表格 
            Dim i As Integer = 0
            For Each dr As DataRow In dt.Select(sFilter, sSort)

                MyRow = New TableRow
                MyCell = New TableCell
                MyCell.Width = Unit.Pixel("30") '"30"
                MyCell.Text = (i + 1).ToString
                MyRow.Cells.Add(MyCell)
                MyCell = New TableCell
                MyHiddenBox = New HtmlInputHidden 'HiddenField
                MyHiddenBox.Value = dr("FunID")
                MyCell.Controls.Add(MyHiddenBox)
                MyRow.Cells.Add(MyCell)
                MyCell = New TableCell
                MyCell.Width = Unit.Pixel("250") '"250"
                MyCell.Text = dr("Name")
                MyRow.Cells.Add(MyCell)
                MyCell = New TableCell
                MyUpBtn = New HtmlInputButton
                MyUpBtn.Attributes.Add("value", "上移")
                MyUpBtn.Attributes.Add("type", "button")
                MyUpBtn.Attributes.Add("onclick", "moveUp(this);")
                MyCell.Controls.Add(MyUpBtn)
                MyRow.Cells.Add(MyCell)
                MyCell = New TableCell
                MyUpBtn = New HtmlInputButton
                MyUpBtn.Attributes.Add("value", "下移")
                MyUpBtn.Attributes.Add("type", "button")
                MyUpBtn.Attributes.Add("onclick", "moveDown(this);")
                MyCell.Controls.Add(MyUpBtn)
                MyRow.Cells.Add(MyCell)

                tbDataTable.Rows.Add(MyRow)
                i += 1
            Next

            panEdit.Visible = True
            If NodeKey = "0" Then
                'param("Levels") = ""
                'param("Parent") = ""
                'dt = db.QueryForDataTableAll(statementsID, param)
                sFilter = " Levels='' and [Parent]='' "
                dt.DefaultView.RowFilter = sFilter
                dt.DefaultView.Sort = sSort

                dt = TIMS.dv2dt(dt.DefaultView)

                AddTreeView(dt, Kind)
                'Treeview1.ExpandLevel 
                'Call Treeview1.ExpandAll()
            End If

        Else
            panEdit.Visible = False

        End If

    End Sub

    Protected Sub dopParentNode_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles dopParentNode.SelectedIndexChanged

        If dopParentNode.SelectedValue <> "" Then
            createTable("0", dopParentNode.SelectedValue, "")
            createL1Node("0", dopParentNode.SelectedValue)
        End If

    End Sub

    Protected Sub dopNextNode_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles dopNextNode.SelectedIndexChanged

        If dopNextNode.SelectedValue <> "" Then
            createTable("1", dopParentNode.SelectedValue, dopNextNode.SelectedValue)
        End If
    End Sub

    '第二層的List
    Sub createL1Node(ByVal NodeKey As String, ByVal Kind As String)
        Dim dt As DataTable
        Dim vsDistID As String = "" & Convert.ToString(sm.UserInfo.DistID)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT  " & vbCrLf
        sql += " case ISNULL(x.Sort,0) WHEN 0 THEN p.Sort else x.Sort end newSort " & vbCrLf
        sql += " ,p.FunID" & vbCrLf ' /*PK*/
        sql += " ,p.Name" & vbCrLf
        sql += " ,p.SPage" & vbCrLf
        sql += " ,p.Kind" & vbCrLf
        sql += " ,p.Levels" & vbCrLf
        sql += " ,p.Parent" & vbCrLf
        sql += " ,p.Valid" & vbCrLf
        sql += " ,p.Sort" & vbCrLf
        sql += " FROM ID_Function p" & vbCrLf
        sql += " left join Plan_Func x on p.FunID=x.FunID AND x.TPlanID='" & planlist.SelectedValue & "'  and x.DistID='" & Convert.ToString(sm.UserInfo.DistID) & "' " & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND exists (" & vbCrLf
        sql += "     select 'x' from auth_groupfun c where c.funid=p.funid" & vbCrLf
        sql += "     and exists (" & vbCrLf
        sql += "         select 'x' from auth_planfunction c2 where c2.tplanid=" & planlist.SelectedValue & "  and c2.funid=c.funid" & vbCrLf
        sql += "     ))" & vbCrLf
        sql += " AND p.Kind='" & Kind & "' " & vbCrLf
        sql += " AND p.Valid='Y'" & vbCrLf
        sql += " AND p.Levels='" & NodeKey & "'" & vbCrLf
        sql += " AND isnull(p.SPage,'0')='0'" & vbCrLf
        sql += " order by p.Kind, p.Levels, 1 , p.Parent"
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            With dopNextNode
                .Visible = True
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "FunID"
                .DataBind()
                .Items.Insert(0, New ListItem("===請選擇===", ""))
            End With
        Else
            dopNextNode.Visible = False
        End If
    End Sub

    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        Dim vsDistID As String = "" & Convert.ToString(sm.UserInfo.DistID)
        Dim vsUserID As String = "" & Convert.ToString(sm.UserInfo.UserID)
        Dim vsTPlanID As String = "" & planlist.SelectedValue
        Dim strSourceFunID As String = ""
        Dim hidField As New HtmlInputHidden  'HiddenField
        Dim arr As Array

        strSourceFunID = txtFunID.Value
        arr = Split(strSourceFunID, ",")

        Common.MessageBox(Me, "資料儲存成功!!")
        Me.Page_Load(Me, e)
    End Sub

    Protected Sub planlist_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles planlist.SelectedIndexChanged
        dopParentNode.SelectedValue = ""
        dopNextNode.Visible = False
        panEdit.Visible = False
    End Sub

    Private Sub AddTreeView(ByVal dt As DataTable, ByVal Kind As String)
        Treeview1.Nodes.Clear()
        Select Case Kind
            Case "TC"
                AddTitleNode(dt, "TC")
            Case "SD"
                AddTitleNode(dt, "SD")
            Case "CP"
                AddTitleNode(dt, "CP")
            Case "TR"
                AddTitleNode(dt, "TR")
            Case "CM"
                AddTitleNode(dt, "CM")
            Case "FM"
                AddTitleNode(dt, "FM")
            Case "OB"
                AddTitleNode(dt, "OB")
            Case "SE"
                AddTitleNode(dt, "SE")
            Case "EXAM"
                AddTitleNode(dt, "EXAM")
            Case "SV"
                AddTitleNode(dt, "SV")
            Case "SYS"
                AddTitleNode(dt, "SYS")
            Case "FAQ"
                AddTitleNode(dt, "FAQ")
            Case "OO"
                AddTitleNode(dt, "OO")
        End Select
    End Sub

    Private Sub AddTitleNode(ByVal dt As DataTable, ByVal Kind As String)
        If dt.Select("Kind='" & Kind & "' and Levels=0").Length <> 0 Or Kind = "FAQ" Then
            ' Microsoft.Web.UI.WebControls.TreeView
            'Dim MyNode As New Microsoft.Web.UI.WebControls.TreeNode ' = New TreeNode
            Dim MyNode As New TreeNode
            Dim vsStudExamUrl As String = ""

            Dim rst As String
            rst = TIMS.Get_MainMenuName(UCase(Kind))
            If rst <> "" Then
                MyNode.Text = "<font color='" & cst_text_color & "'>" & rst & "</font>"
            End If

            'Treeview1.Nodes.Add(MyNode)
            TreeView1.Nodes.Add(MyNode)

            Select Case Kind
                Case "SE"
                Case Else
                    AddNode(dt, Kind, , MyNode)
            End Select
        End If
    End Sub

    Private Sub AddNode(ByVal dt As DataTable, ByVal Kind As String, Optional ByVal drChild As DataRow = Nothing, _
        Optional ByVal ParentNode As TreeNode = Nothing)
        'Microsoft.Web.UI.WebControls.TreeNode
        'TreeNode 
        Dim vsUserID As String = "" & Convert.ToString(sm.UserInfo.UserID)
        Dim vsDistID As String = "" & Convert.ToString(sm.UserInfo.DistID)
        Dim vsLID As String = "" & Convert.ToString(sm.UserInfo.LID)
        If vsLID <> "" Then vsLID = CInt(vsLID)

        Dim dr As DataRow
        'Dim MyNode As Microsoft.Web.UI.WebControls.TreeNode 'TreeNode
        Dim MyNode As TreeNode

        'Const cst_FunLength_set As Integer = 7 '10
        If drChild Is Nothing Then
            For Each dr In dt.Select("Kind='" & Kind & "' and Levels=0", "newSort")
                Dim FunName As String = ""
                Dim FunLength As Integer = dr("Name").ToString.Length

                'MyNode = New Microsoft.Web.UI.WebControls.TreeNode 'TreeNode
                MyNode = New TreeNode

                FunName = dr("Name")
                If IsDBNull(dr("SPage")) Then
                    MyNode.Text = "<font color='" & cst_text_color & "'>" & FunName & "</font>"
                Else
                    'MyNode.Text = "<img src='images/i2/point.gif' border='0' align='absmiddle'><font color='" & cst_text_color & "'>" & "&nbsp;&nbsp;" & FunName & "</font>"
                    MyNode.Text = "<img src='../../images/point.gif' border='0' align='absmiddle'><font color='" & cst_text_color & "'>" & FunName & "</font>"
                End If

                Dim vsIDNO As String = ""
                Dim vsuser_ID As String = ""
                Dim vspassword As String = ""

                If dt.Select("Kind='" & Kind & "' and [Parent]='" & dr("FunID") & "'").Length <> 0 Then
                    AddNode(dt, Kind, dr, MyNode)
                End If

                'ParentNode.Nodes.Add(MyNode)
                ParentNode.ChildNodes.Add(MyNode)
            Next
        Else
            For Each dr In dt.Select("Kind='" & Kind & "' and [Parent]='" & drChild("FunID") & "'", "newSort")
                Dim FunName As String = ""
                Dim FunLength As Integer = dr("Name").ToString.Length

                'MyNode = New Microsoft.Web.UI.WebControls.TreeNode 'TreeNode
                MyNode = New TreeNode

                If FunLength > 6 Then
                    Dim i As Integer = 0

                    FunName = dr("Name").ToString
                    MyNode.Text = "<font color='" & cst_text_color & "'>。" & FunName & "</font>"
                Else
                    MyNode.Text = "<font color='" & cst_text_color & "'>。" & dr("Name") & "</font>"
                End If


                'ParentNode.Nodes.Add(MyNode)
                ParentNode.ChildNodes.Add(MyNode)

                If dt.Select("Kind='" & Kind & "' and [Parent]='" & dr("FunID") & "'").Length <> 0 Then
                    AddNode(dt, Kind, dr, MyNode)
                End If
            Next
        End If
    End Sub

End Class
