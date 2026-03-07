Partial Class LevOrg1
    Inherits AuthBasePage

    'SELECT * FROM AUTH_ACCTORG WHERE Acct1='shamesala' or Acct2='shamesala' or Acct3='shamesala' or Acct4='shamesala'
    Dim Gdt As DataTable = Nothing
    '判斷機構是否只有一個班級
    Dim drOCID As DataRow = Nothing
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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'txtSearch1.Attributes("onkeypress")="Search_click();return false;"

        If Not Page.IsPostBack Then
            txtSearch1.Attributes("onkeypress") = "Search_click();"
            btnSearch.Attributes("onclick") = "Search_click2();return false;"
            btnSearch.Attributes("onkeypress") = "Search_click2();return false;"

            Dim r_btnName As String = TIMS.ClearSQM(Request("btnName"))
            If r_btnName <> "" Then hidbtnName.Value = r_btnName '按鈕

            Call Search1() 'sender, e
        End If
    End Sub

    Sub AddTreeNodes(ByVal dr As DataRow, ByVal objTable As DataTable, ByVal objTreeView As TreeView, ByVal ParentNode As TreeNode)
        Dim NewNode As New TreeNode
        Dim drChild As DataRow
        Dim strFilter As String

        NewNode.Text = $"{dr("OrgName")}"
        '用Nothing 與空白來做顯示方式(OrgName2,OrgName)判斷
        'NewNode.Text=If(Hid_OrgName2.Value="", dr("OrgName").ToString(), dr("OrgName2").ToString())
        'NewNode.NavigateUrl="javascript:returnValue('" & dr("RID") & "','" & dr("OrgName") & "','" & Request("btnName") & "');"
        'NewNode.NavigateUrl="javascript:returnValue('" & dr("RID") & "','" & dr("OrgName") & "','" & Request("btnName") & "','" & dr("OrgID") & "');"
        Dim RqbtnName As String = TIMS.ClearSQM(Request("btnName")) 'Request("btnName")
        Dim vsNavigateUrl As String = $"javascript:returnValue('{dr("RID")}','{dr("OrgName")}','{RqbtnName}','{dr("OrgID")}','{dr("isBlack")}');"
        NewNode.NavigateUrl = vsNavigateUrl

        If ParentNode Is Nothing Then
            objTreeView.Nodes.Add(NewNode)
        Else
            'ParentNode.Nodes.Add(NewNode)
            ParentNode.ChildNodes.Add(NewNode)
        End If

        '加入子節點
        Dim strRid As String = $"{dr("RID")}/"
        strFilter = $"Relship like '%{strRid}%'"        '先找出符合父節點 xxx\ 開頭的關係
        For Each drChild In objTable.Select(strFilter, "OrgName")
            Dim strRelship As String = drChild("Relship")
            Dim pos As Integer = strRelship.IndexOf(strRid)

            '若出現格式為「%父節點/子節點/」的Relship值，視為子節點
            If pos <> -1 And (pos + strRid.Length) < strRelship.Length Then
                If strRelship.IndexOf("/", pos + strRid.Length) = strRelship.Length - 1 Then
                    If Gdt IsNot Nothing Then
                        If Gdt.Select("RID='" & drChild("RID") & "'").Length <> 0 Or Split(strRelship, "/").Length > 4 Then AddTreeNodes(drChild, objTable, objTreeView, NewNode)
                    Else
                        AddTreeNodes(drChild, objTable, objTreeView, NewNode)
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim CookieTable_CanSave As Boolean = False

        '確認RID 業務權限 與PLANID 計畫 為登入權限 sm.UserInfo.PlanID 
        CookieTable_CanSave = TIMS.CheckRIDsPLAN(Me, OrgRID.Value, objconn)

        If CookieTable_CanSave Then
            dt = TIMS.GetCookieTable(Me, da, objconn)
            For i As Integer = 1 To 5
                If dt.Select(String.Format("ItemName='LevOrgRID{0}'", i)).Length = 0 Then
                    Dim InsertFlag As Boolean = True
                    For j As Integer = 1 To 5
                        If dt.Select(String.Format("ItemName='LevOrgRID{0}' and ItemValue='{1}'", j, OrgRID.Value)).Length <> 0 Then
                            InsertFlag = False
                            Exit For
                        End If
                    Next
                    If InsertFlag = True Then
                        TIMS.InsertCookieTable(Me, dt, da, "LevOrgName" & i, OrgName.Value, False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "LevOrgRID" & i, OrgRID.Value, True, objconn)
                    End If
                    Exit For
                Else
                    If i = 5 Then
                        For j As Integer = 1 To 4
                            Dim NewDr As DataRow = dt.Select("ItemName='LevOrgRID" & j + 1 & "'")(0)
                            Dim OldDr As DataRow = dt.Select("ItemName='LevOrgRID" & j & "'")(0)
                            OldDr("ItemValue") = NewDr("ItemValue")
                            NewDr = dt.Select("ItemName='LevOrgName" & j + 1 & "'")(0)
                            OldDr = dt.Select("ItemName='LevOrgName" & j & "'")(0)
                            OldDr("ItemValue") = NewDr("ItemValue")
                        Next
                        Dim InsertFlag As Boolean = True
                        For j As Integer = 1 To 5
                            If dt.Select("ItemName='LevOrgRID" & j & "' and ItemValue='" & OrgRID.Value & "'").Length <> 0 Then
                                InsertFlag = False
                                Exit For
                            End If
                        Next
                        If InsertFlag = True Then
                            TIMS.InsertCookieTable(Me, dt, da, "LevOrgName" & i, OrgName.Value, False, objconn)
                            TIMS.InsertCookieTable(Me, dt, da, "LevOrgRID" & i, OrgRID.Value, True, objconn)
                        End If
                    End If
                End If
            Next
        End If

        '判斷機構是否只有一個班級
        drOCID = Nothing
        drOCID = TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1, objconn)
        If drOCID IsNot Nothing Then
            Common.RespWrite(Me, "<script language=javascript>")
            Common.RespWrite(Me, "function GetValue(){")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMID1')!=null) window.opener.document.form1.TMID1.value='" & TMID1.Text & "';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMIDValue1')!=null) window.opener.document.form1.TMIDValue1.value='" & TMIDValue1.Value & "';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCID1')!=null) window.opener.document.form1.OCID1.value='" & OCID1.Text & "';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCIDValue1')!=null) window.opener.document.form1.OCIDValue1.value='" & OCIDValue1.Value & "';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('search')!=null) window.opener.document.getElementById('search').click();")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TDate')!=null) window.opener.document.form1.TDate.value='" & drOCID("FTDate") & "';")
            Common.RespWrite(Me, "    window.close();")
            Common.RespWrite(Me, "}")
            Common.RespWrite(Me, "GetValue();")
            Common.RespWrite(Me, "</script>")
            Common.RespWrite(Me, "<script>")
            If hidbtnName.Value <> "" Then Common.RespWrite(Me, "    if (window.opener.document.getElementById('" & hidbtnName.Value & "')) window.opener.document.getElementById('" & hidbtnName.Value & "').click();")
            Common.RespWrite(Me, "window.close();</script>")
        Else
            Common.RespWrite(Me, "<script language=javascript>")
            Common.RespWrite(Me, "function GetValue3(){")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMID1')!=null) window.opener.document.form1.TMID1.value='';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMIDValue1')!=null) window.opener.document.form1.TMIDValue1.value='';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCID1')!=null) window.opener.document.form1.OCID1.value='';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCIDValue1')!=null) window.opener.document.form1.OCIDValue1.value='';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TDate')!=null) window.opener.document.form1.TDate.value ='';")
            Common.RespWrite(Me, "    window.close();")
            Common.RespWrite(Me, "}")
            Common.RespWrite(Me, "GetValue3();")
            Common.RespWrite(Me, "</script>")
            Common.RespWrite(Me, "<script>window.close();</script>")
        End If
    End Sub

    '隱藏式 Client 按鈕
    Private Sub hbtnSearch_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hbtnSearch.ServerClick
        '執行搜尋
        Call Search1() 'sender, e
        Call TreeView1.ExpandAll()  '以程式設計方式展開節點
    End Sub

    'SQL
    Sub Search1()
        'ByVal sender As System.Object, ByVal e As System.EventArgs
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        '檢查Session是否存在 End

        Dim strMaxOrgLevel As String = ""
        strMaxOrgLevel = TIMS.Get_MaxOrgLevel(sm.UserInfo.PlanID, objconn)
        txtSearch1.Text = TIMS.ClearSQM(txtSearch1.Text)
        ViewState("Search1") = txtSearch1.Text

        Dim objtable As New DataTable
        'treeview begin
        Me.TreeView1.Nodes.Clear()
        Dim strFilter As String = ""
        Dim dr As DataRow = Nothing
        '用Nothing 與空白來做顯示方式(OrgName2,OrgName)判斷
        'Me.ViewState.Remove("OrgName2")
        'Hid_OrgName2.Value="" 'Nothing

        Dim parms As New Hashtable
        Dim sql As String = ""
        Dim rq_selected_year As String = TIMS.ClearSQM(Request("selected_year"))
        If rq_selected_year <> "" Then
            '用Nothing 與空白來做顯示方式(OrgName2,OrgName)判斷
            'Hid_OrgName2.Value=""
            sql = ""
            sql &= " SELECT ISNULL(a.OrgName+'('+ip.Seq+')',a.OrgName) OrgName2" & vbCrLf
            sql &= "  ,a.OrgID, a.OrgKind, a.OrgName, a.ComIDNO, a.ComCIDNO" & vbCrLf
            sql &= "  ,a.IsConUnit, a.TradeID, a.EmpNum, a.OrgUrl, a.OrgKind2, a.LastYearExeRate" & vbCrLf
            sql &= "  ,a.IsConTTQS, a.BankName, a.ExBankName, a.AccNo, a.AccName" & vbCrLf
            sql &= "  ,b.RSID, b.PlanID, b.RID, b.Relship, b.OrgLevel, b.DistID" & vbCrLf
            sql &= "  ,CASE WHEN ob.ComIDNO is not null THEN 'Y' ELSE 'N' END AS isBlack" & vbCrLf
            sql &= " FROM Org_OrgInfo a" & vbCrLf
            sql &= " JOIN (SELECT * FROM Auth_Relship WHERE PlanID=0 OR (PlanID IN (SELECT DISTINCT vp.Planid" & vbCrLf
            sql &= "  FROM VIEW_LOGINPLAN vp" & vbCrLf
            sql &= "  JOIN VIEW_RIDNAME vr on vp.PlanID=vr.PlanID" & vbCrLf
            sql &= "  WHERE vp.YEARS=@VPYEARS" & vbCrLf
            sql &= "  AND vp.TPLANID=@VPTPLANID" & vbCrLf
            sql &= "  AND vr.RID like @VRRID+'%'" & vbCrLf
            sql &= "  AND vp.PlanID <> 0))) b ON a.OrgID=b.OrgID" & vbCrLf
            sql &= " JOIN ORG_ORGPLANINFO c ON b.RSID=c.RSID" & vbCrLf
            sql &= " LEFT JOIN ID_Plan ip ON b.PlanID=ip.PlanID" & vbCrLf
            sql &= " LEFT JOIN (SELECT distinct dob.ComIDNO FROM Org_BlackList dob WHERE dob.Avail='Y' AND (DATEDIFF(day,dob.OBSDATE,GETDATE())>=0 and DATEDIFF(day, getdate(),DATEADD(month, 12*dob.obYears,dob.OBSDATE))>=0)) ob ON ob.ComIDNO=a.ComIDNO" & vbCrLf
            parms.Add("VPYEARS", rq_selected_year)
            parms.Add("VPTPLANID", sm.UserInfo.TPlanID)
            parms.Add("VRRID", sm.UserInfo.RID)

            If Me.ViewState("Search1") <> "" Then
                Select Case strMaxOrgLevel
                    Case "1", "2"
                        sql &= " WHERE (1!=1 OR b.OrgLevel IN (0,1)" & vbCrLf
                        sql &= "  OR (a.OrgName LIKE '%'+@Search1+'%' AND b.OrgLevel <= " & strMaxOrgLevel & "))" & vbCrLf
                        parms.Add("Search1", Me.ViewState("Search1"))
                    Case "3"
                        sql &= " WHERE (1!=1 OR b.OrgLevel IN (0,1,2)" & vbCrLf
                        sql &= "  OR (a.OrgName LIKE '%'+@Search1+'%' AND b.OrgLevel <= " & strMaxOrgLevel & "))" & vbCrLf
                        parms.Add("Search1", Me.ViewState("Search1"))
                    Case Else
                        sql &= " WHERE (1!=1 OR b.OrgLevel IN (0,1,2,3)" & vbCrLf
                        sql &= "  OR (a.OrgName LIKE '%'+@Search1+'%' AND b.OrgLevel <= " & strMaxOrgLevel & "))" & vbCrLf
                        parms.Add("Search1", Me.ViewState("Search1"))
                End Select
            End If
            'objstr=sql
            'sql=Nothing
        Else
            sql = ""
            sql &= " SELECT a.OrgID, a.OrgKind, a.OrgName, a.ComIDNO, a.ComCIDNO" & vbCrLf
            sql &= "  ,a.IsConUnit, a.TradeID, a.EmpNum, a.OrgUrl, a.OrgKind2, a.LastYearExeRate" & vbCrLf
            sql &= "  ,a.IsConTTQS, a.BankName, a.ExBankName, a.AccNo, a.AccName" & vbCrLf
            sql &= "  ,b.RSID, b.PlanID, b.RID, b.Relship, b.OrgLevel, b.DistID" & vbCrLf
            sql &= "  ,CASE WHEN ob.ComIDNO IS NOT NULL THEN 'Y' ELSE 'N' END AS isBlack" & vbCrLf
            sql &= " FROM Org_OrgInfo a" & vbCrLf
            sql &= " JOIN (SELECT * FROM Auth_Relship WHERE PlanID=0 OR (PlanID=" & sm.UserInfo.PlanID & " AND PlanID <> 0)) b ON a.OrgID=b.OrgID" & vbCrLf
            sql &= " JOIN Org_OrgPlanInfo c ON b.RSID=c.RSID" & vbCrLf
            sql &= " LEFT JOIN (SELECT distinct dob.ComIDNO FROM Org_BlackList dob WHERE dob.Avail='Y' AND (DATEDIFF(day,dob.obSdate,GETDATE())>=0 AND DATEDIFF(day, GETDATE(),DATEADD(month, 12*dob.obYears,dob.obSdate))>=0)) ob ON ob.ComIDNO=a.ComIDNO" & vbCrLf

            If Me.ViewState("Search1") <> "" Then
                Select Case strMaxOrgLevel
                    Case "1", "2"
                        sql &= " WHERE (1!=1 OR b.OrgLevel IN (0,1)" & vbCrLf
                        sql &= "  OR (a.OrgName LIKE '%'+@Search1+'%' AND b.OrgLevel <= " & strMaxOrgLevel & "))" & vbCrLf
                        parms.Add("Search1", Me.ViewState("Search1"))
                    Case "3"
                        sql &= " WHERE (1!=1 OR b.OrgLevel IN (0,1,2)" & vbCrLf
                        sql &= "  OR (a.OrgName LIKE '%'+@Search1+'%' AND b.OrgLevel <= " & strMaxOrgLevel & "))" & vbCrLf
                        parms.Add("Search1", Me.ViewState("Search1"))
                    Case Else
                        sql &= " WHERE (1!=1 OR b.OrgLevel IN (0,1,2,3)" & vbCrLf
                        sql &= "  OR (a.OrgName LIKE '%'+@Search1+'%' and b.OrgLevel <= " & strMaxOrgLevel & "))" & vbCrLf
                        parms.Add("Search1", Me.ViewState("Search1"))
                End Select
            End If
        End If

        'If sm.UserInfo.LID=1 AndAlso sm.UserInfo.RoleID <> 1 AndAlso sm.UserInfo.RoleID <> 0 AndAlso sm.UserInfo.RoleID <> 99 Then
        If sm.UserInfo.LID = 1 Then
            '計畫賦予設定  二級 與 承辦人 可做委訓單位歸屬
            Dim sqlxa As String = ""
            Select Case sm.UserInfo.RoleID
                Case "2"
                    sqlxa = " SELECT DISTINCT RID FROM AUTH_ACCTORG WHERE Acct4='" & sm.UserInfo.UserID & "' "
                Case "3"
                    sqlxa = " SELECT DISTINCT RID FROM AUTH_ACCTORG WHERE Acct3='" & sm.UserInfo.UserID & "' "
                Case "4"
                    sqlxa = " SELECT DISTINCT RID FROM AUTH_ACCTORG WHERE Acct2='" & sm.UserInfo.UserID & "' "
                Case "5"
                    sqlxa = " SELECT DISTINCT RID FROM AUTH_ACCTORG WHERE Acct1='" & sm.UserInfo.UserID & "' "
            End Select
            If sqlxa <> "" Then
                TIMS.WriteLog_1(Me, "LevOrg1 ##Gdt", sqlxa, Nothing)
                Gdt = DbAccess.GetDataTable(sqlxa, objconn)
            End If
        End If
        TIMS.WriteLog_1(Me, "LevOrg1", sql, parms)
        objtable = DbAccess.GetDataTable(sql, objconn, parms)

        'select *  FROM Auth_AcctOrg WHERE Acct1='gmoa'
        'TIMS.Fill(objstr, objAdapter, objtable)
        '  Common.RespWrite(Me, objstr)
        'Dim objconn As SqlConnection
        'objconn=DbAccess.GetConnection()
        'objAdapter=New SqlDataAdapter(objstr, objconn)
        'objAdapter.Fill(objtable)

        strFilter = "RID='" & sm.UserInfo.RID & "'"
        ' Common.RespWrite(Me, strFilter)
        For Each dr In objtable.Select(strFilter)
            AddTreeNodes(dr, objtable, Me.TreeView1, Nothing)
        Next
        'treeview end
    End Sub
End Class