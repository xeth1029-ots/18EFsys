Partial Class LevOrg
    Inherits AuthBasePage

    'Dim fg_test_ISBLACK As Boolean=True '測試黑名單功能
    Dim dtOrgBlack As DataTable  '取出系統現有黑名單
    Dim objConn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        TIMS.CloseDbConn(objConn)
    End Sub

    '署、分署(局、中心)使用此 機構搜尋功能
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        '(直接在AuthBasePage處理,不用個別檢查 Session) TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objConn, "") 'Response.End()
        End If
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        If Not TIMS.OpenDbConn(objConn) Then Return
        '檢查Session是否存在 End

        dtOrgBlack = TIMS.Get_OrgBlackList(Me, objConn)  '取出系統現有黑名單

        If Not Page.IsPostBack Then
            CCreate1()
        End If
    End Sub

    Sub CCreate1()
        DistID.Attributes("onchange") = "GetRID();"
        txtSearch1.Attributes("onkeypress") = "Search_click();"
        btnSearch.Attributes("onclick") = "Search_click2();return false;"
        btnSearch.Attributes("onkeypress") = "Search_click2();return false;"
        Dim rqName As String = TIMS.ClearSQM(Request("name")) '隱藏欄位
        Dim rqbtnName As String = TIMS.ClearSQM(Request("btnName")) '按鈕

        Hidnorus.Value = ""
        TableSearch.Visible = False '搜尋關鍵字

        '自己的暫存自已清理
        TIMS.DeleteCookieTableOld3(sm, objConn)

        hid_ORGBLACKLIST.Value = ""
        If TIMS.Check_OrgBlackList(Me, "", objConn) Then hid_ORGBLACKLIST.Value = "Y"  '確認登入帳號之機構是否在黑名單中
        hidname.Value = "" '取得名稱參數
        If rqName <> "" Then hidname.Value = rqName
        'btnCheckClass 採特殊動作
        If rqbtnName <> "" Then hidbtnName.Value = rqbtnName
        '依據名稱參數條件，執行搜尋
        If hidname.Value = "Stu_Maintain" Then
            CStu_Maintain() '學員資料維護專用
        Else
            CSearch1()  '班級使用
        End If
    End Sub

    '顯示下拉轄區
    Sub SUtl_SetDistIDName()
        Dim myValue As String = ""
        DistID.Items.Add(New ListItem("==請選擇==", ""))
        If sm.UserInfo.RID = "A" Then myValue = TIMS.Get_DistName2("000") : DistID.Items.Add(New ListItem(myValue, "000"))
        Dim sql As String = " SELECT DISTID, NAME FROM ID_DISTRICT WHERE DISTID !='000' ORDER BY DISTID"
        Dim dtDist As DataTable = DbAccess.GetDataTable(sql, objConn)
        For i As Integer = 0 To dtDist.Rows.Count - 1
            Dim dr As DataRow = dtDist.Rows(i)
            myValue = TIMS.Get_DistName2(dr("DISTID"))
            If myValue = "" Then myValue = Convert.ToString(dr("NAME")) '空白使用原資料。
            DistID.Items.Add(New ListItem(myValue, dr("DISTID")))
        Next
    End Sub

    '顯示下拉轄區(單1)
    Sub SUtl_SetDistIDName(ByVal ssDistid As String)
        Dim myValue As String = ""
        DistID.Items.Add(New ListItem("==請選擇==", ""))
        myValue = TIMS.Get_DistName2(ssDistid) : DistID.Items.Add(New ListItem(myValue, ssDistid))  '只可選自己轄區，不可跨區
    End Sub

    '一般資料搜尋 (班級使用)
    Sub CSearch1()
        '檢查Session是否存在 Start
        '(直接在AuthBasePage處理,不用個別檢查 Session) TIMS.CheckSession(Me, 9)
        '檢查Session是否存在 End

        Dim rqGetOther As String = TIMS.ClearSQM(Request("GetOther"))
        Dim rqCOPY02 As String = TIMS.ClearSQM(Request("COPY02"))

        'schType :搜尋方式 '1:一般搜尋 '0:署(局)搜尋 '36:青年職涯搜尋 '2013:2013年後自辦搜尋
        Dim schType As Integer = 1 '1:一般搜尋
        DistID.Items.Clear()
        If sm.UserInfo.RID = "A" OrElse (sm.UserInfo.DistID <> "000" AndAlso sm.UserInfo.RoleID = "1" AndAlso rqGetOther = "1") Then
            '如果是署(局), 就顯示選擇年度轄區的下拉式選單
            schType = 0 '0:署(局)搜尋 
        ElseIf sm.UserInfo.DistID = "002" AndAlso sm.UserInfo.TPlanID = "36" Then
            schType = 36 '36:青年職涯搜尋
        End If
        '非署(局) 自辦 2013年後 (自辦應為中心)
        If rqCOPY02 = "02" Then
            If sm.UserInfo.RID <> "A" AndAlso sm.UserInfo.TPlanID = "02" AndAlso sm.UserInfo.Years >= 2013 Then schType = 2013 '2013:2013年後自辦搜尋
        End If

        Select Case schType
            Case 0 '0:署(局)搜尋 
                '如果是署(局), 就顯示選擇年度轄區的下拉式選單
                supershow.Visible = True
                '加入年度 by nick
                Downyear = TIMS.GetSyear(Downyear)
                '2005/4/1--Melody年度帶預設值
                Common.SetListItem(Downyear, sm.UserInfo.Years)
                SUtl_SetDistIDName()
                '如果非署(局)，則自動選擇轄區，
                If sm.UserInfo.DistID <> "000" Then
                    Common.SetListItem(DistID, sm.UserInfo.DistID)
                    DistID_Changed()
                    RID.Value = sm.UserInfo.RID
                    Common.SetListItem(planlist, sm.UserInfo.PlanID)
                    Planlist_Changed()
                Else
                    If planlist.Items.FindByValue("") Is Nothing Then
                        '第一次無此值，可塞入 請選擇轄區
                        planlist.Items.Remove(planlist.Items.FindByValue(""))
                        planlist.Items.Insert(0, New ListItem("==請選擇轄區==", ""))
                    Else
                        '有此值，重新搜尋工作
                        Planlist_Changed()
                    End If
                End If
            Case 36 '36:青年職涯搜尋
                supershow.Visible = True
                '加入年度 by nick
                Downyear = TIMS.GetSyear(Downyear)
                '2005/4/1--Melody年度帶預設值
                Common.SetListItem(Downyear, sm.UserInfo.Years)
                SUtl_SetDistIDName()
                planlist.Items.Remove(planlist.Items.FindByValue(""))
                planlist.Items.Insert(0, New ListItem("==請選擇轄區==", ""))
            Case 2013 '2013:2013年後自辦搜尋
                supershow.Visible = True
                '加入年度 by nick
                Downyear = TIMS.GetSyear(Downyear)
                '2005/4/1--Melody年度帶預設值
                Common.SetListItem(Downyear, sm.UserInfo.Years)
                SUtl_SetDistIDName(sm.UserInfo.DistID)
                '如果非署(局)，則自動選擇轄區，
                Common.SetListItem(DistID, sm.UserInfo.DistID)
                DistID_Changed() '依轄區年度 => 選擇計畫下拉
                RID.Value = sm.UserInfo.RID '自辦應為分署(中心)
                Common.SetListItem(planlist, sm.UserInfo.PlanID) '依年度計畫
                Planlist_Changed()
            Case Else '1:一般搜尋
                '選單關閉
                supershow.Visible = False
                Planlist_Changed()
        End Select
    End Sub

    '學員資料維護專用
    Sub CStu_Maintain()
        '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9) '檢查Session是否存在 End

        Dim rqGetOther As String = TIMS.ClearSQM(Request("GetOther"))
        DistID.Items.Clear()
        '如果是署(局), 就顯示選擇年度轄區的下拉式選單
        If sm.UserInfo.RID = "A" Or (sm.UserInfo.DistID <> "000" And sm.UserInfo.RoleID = "1" And rqGetOther = "1") Then
            supershow.Visible = True
            '加入年度 by nick
            Downyear = TIMS.GetSyear(Downyear)
            '2005/4/1--Melody年度帶預設值
            Common.SetListItem(Downyear, sm.UserInfo.Years)
            SUtl_SetDistIDName()
            Common.SetListItem(DistID, sm.UserInfo.DistID)
            DistID_Changed()
            RID.Value = sm.UserInfo.RID
            Common.SetListItem(planlist, sm.UserInfo.PlanID)
            Planlist_Changed()
            If planlist.Items.FindByValue("") Is Nothing Then
                '第一次無此值，可塞入 請選擇轄區
                planlist.Items.Remove(planlist.Items.FindByValue(""))
                planlist.Items.Insert(0, New ListItem("==請選擇轄區==", ""))
            Else
                '有此值，重新搜尋工作
                Planlist_Changed()
            End If
        ElseIf sm.UserInfo.DistID = "002" AndAlso sm.UserInfo.TPlanID = "36" Then
            supershow.Visible = True
            '加入年度 by nick
            Downyear = TIMS.GetSyear(Downyear)
            '2005/4/1--Melody年度帶預設值
            Common.SetListItem(Downyear, sm.UserInfo.Years)
            SUtl_SetDistIDName()
            planlist.Items.Remove(planlist.Items.FindByValue(""))
            planlist.Items.Insert(0, New ListItem("==請選擇轄區==", ""))
            Common.SetListItem(DistID, sm.UserInfo.DistID)
            DistID_Changed()
            RID.Value = sm.UserInfo.RID
            Common.SetListItem(planlist, sm.UserInfo.PlanID)
            Planlist_Changed()
        Else
            '否則,選單關閉
            supershow.Visible = False
            Planlist_Changed()
        End If
    End Sub

    '加入節點
    Sub AddTreeNodes(ByVal dr As DataRow, ByVal objTable As DataTable, ByVal objTreeView As TreeView, ByVal ParentNode As TreeNode)
        Dim NewNode As New TreeNode
        Dim drChild As DataRow
        Dim strFilter As String = ""
        Dim strNavigateUrl As String = ""
        Dim rqbtnName As String = TIMS.ClearSQM(Request("btnName"))
        Dim rqSubmit As String = TIMS.ClearSQM(Request("submit"))

        'Relship
        Dim rValue As String = ""
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("RID")))
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("ORGNAME")))
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("COMIDNO")))
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("CONTACTEMAIL")))
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("ORGID")))
        TIMS.AddSQMValue(rValue, rqSubmit)
        TIMS.AddSQMValue(rValue, String.Concat("(", dr("ZIPCODE"), ")", TIMS.Get_ZipName(dr("ZIPCODE"), objConn)))
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("ZIPCODE")))
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("ADDRESS")))
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("ORGLEVEL")))
        TIMS.AddSQMValue(rValue, Convert.ToString(dr("ISBLACK")))
        TIMS.AddSQMValue(rValue, DistID.SelectedValue) 'DISTID (add for甄試通知單設定)
        TIMS.AddSQMValue(rValue, planlist.SelectedValue) 'PLANID (add for甄試通知單設定)

        If hidbtnName.Value = "btnCheckClass" Then
            TIMS.AddSQMValue(rValue, "") 'btnbtn
        Else
            TIMS.AddSQMValue(rValue, rqbtnName) 'btnbtn
        End If

        strNavigateUrl = "javascript:returnValue(" & rValue & ");"
        NewNode.NavigateUrl = strNavigateUrl

        '用Nothing 與空白來做顯示方式(OrgName2, OrgName)判斷
        NewNode.Text = If(ViewState("ORGNAME2") Is Nothing, dr("ORGNAME"), dr("ORGNAME2"))
        'If fg_test_ISBLACK AndAlso Convert.ToString(dr("ISBLACK"))="Y" Then TIMS.Tooltip(NewNode, "(機構已列入黑名單)")

        If ParentNode Is Nothing Then
            objTreeView.Nodes.Add(NewNode)
        Else
            ParentNode.ChildNodes.Add(NewNode)
        End If

        '加入子節點
        Dim strRid As String = dr("RID") & "/"
        strFilter = "RELSHIP LIKE '%" & strRid & "%'"        '先找出符合父節點 xxx\ 開頭的關係
        For Each drChild In objTable.Select(strFilter, "ORGNAME")
            Dim strRelship As String = drChild("RELSHIP")
            Dim pos As Integer = strRelship.IndexOf(strRid)
            '若出現格式為「%父節點/子節點/」的Relship值，視為子節點
            If pos <> -1 And (pos + strRid.Length) < strRelship.Length Then
                If strRelship.IndexOf("/", pos + strRid.Length) = strRelship.Length - 1 Then
                    If sm.UserInfo.DistID = "000" Then
                        If drChild("DISTID") = DistID.SelectedValue Then AddTreeNodes(drChild, objTable, objTreeView, NewNode)
                    Else
                        AddTreeNodes(drChild, objTable, objTreeView, NewNode)
                    End If
                End If
            End If
        Next
    End Sub

    Sub DistID_Changed()
        planlist.Items.Clear()
        TreeView1.Nodes.Clear()

        Dim vDownyear As String = TIMS.GetListValue(Downyear)
        Dim vDistID As String = TIMS.GetListValue(DistID)
        'If vDownyear="" OrElse vDistID="" Then Return
        If DistID.SelectedIndex <> 0 AndAlso vDistID <> "" Then
            'Dim dr As DataRow 'parms.Clear() 'edit，by:20181001
            Dim parms As New Hashtable From {{"YEARS", vDownyear}, {"DISTID", vDistID}, {"CLSYEAR", vDownyear}}
            Dim sSql As String = "
SELECT CONCAT(a.YEARS,c.NAME,b.PLANNAME,a.SEQ) PLANNAME ,a.PLANID,b.TPLANID
FROM ID_PLAN a
JOIN ID_DISTRICT c ON a.DistID=c.DistID AND a.YEARS=@YEARS AND c.DISTID=@DISTID
JOIN KEY_PLAN b ON a.TPLANID=b.TPLANID AND (b.CLSYEAR IS NULL OR b.CLSYEAR > @CLSYEAR)
ORDER BY b.TPLANID,a.PLANID
"
            Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, parms)

            Dim vsPLANID As String = ""
            If $"{sm.UserInfo.TPlanID}" <> "" Then
                Dim ff3 As String = $"TPLANID='{sm.UserInfo.TPlanID}'"
                If dt.Select(ff3).Length > 0 Then vsPLANID = $"{dt.Select(ff3)(0)("PLANID")}"
            End If

            If TIMS.dtNODATA(dt) Then
                planlist.Items.Insert(0, New ListItem("==查無資料==", ""))
            Else
                planlist.DataSource = dt
                planlist.DataTextField = "PlanName"
                planlist.DataValueField = "PlanID"
                planlist.DataBind()
                planlist.Items.Insert(0, New ListItem("==請選擇==", ""))

                If dt.Rows.Count = 1 Then
                    If planlist.SelectedItem IsNot Nothing Then planlist.SelectedItem.Selected = False
                    planlist.Items(1).Selected = True
                    Planlist_Changed()
                ElseIf vsPLANID <> "" Then
                    Common.SetListItem(planlist, vsPLANID)
                    Planlist_Changed()
                End If
            End If
        Else
            planlist.Items.Insert(0, New ListItem("==請選擇轄區==", ""))
        End If
    End Sub

    '轄區選擇後 SQL
    Private Sub DistID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DistID.SelectedIndexChanged
        DistID_Changed()
    End Sub

    Sub Planlist_Changed()
        '檢查Session是否存在 Start
        ' (直接在AuthBasePage處理,不用個別檢查Session) TIMS.CheckSession(Me, 9)
        '檢查Session是否存在 End
        objConn = DbAccess.GetConnection()
        If Not TIMS.OpenDbConn(objConn) Then Return
        'ViewState.Remove("SEARCH1")
        txtSearch1.Text = TIMS.Get_Substr1(TIMS.ClearSQM(txtSearch1.Text), 1000)
        ViewState("SEARCH1") = txtSearch1.Text 'If(txtSearch1.Text <> "", txtSearch1.Text, "")

        '用Nothing 與空白來做顯示方式(OrgName2, OrgName)判斷
        ViewState("ORGNAME2") = Nothing

        Dim rqGetOther As String = TIMS.ClearSQM(Request("GetOther"))
        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqSelected_year As String = TIMS.ClearSQM(Request("selected_year"))
        'If (rqSelected_year <> "") Then rqSelected_year=rqSelected_year
        Dim planlist_SelectVal As String = TIMS.ClearSQM(planlist.SelectedValue)

        '計畫未選擇
        If planlist.SelectedIndex <> 0 Then
            'If planlist_SelectVal="" OrElse Not TIMS.IsNumberStr(planlist_SelectVal) Then Return
            Dim strMaxOrgLevel As String = TIMS.ClearSQM(TIMS.Get_MaxOrgLevel(planlist_SelectVal, objConn))

            'treeview begin
            TreeView1.Nodes.Clear()
            Dim strFilter As String = ""
            Dim parms As New Hashtable
            Dim objstr As String = ""
            Dim dr As DataRow = Nothing

            TableSearch.Visible = True '可使用搜尋關鍵字
            '如果是署(局), 則從supershow來選
            If sm.UserInfo.RID = "A" OrElse (sm.UserInfo.RoleID = "1" AndAlso rqGetOther = "1") Then
                If planlist_SelectVal = "" Then planlist_SelectVal = -1
                Dim sql As String = ""
                sql &= " SELECT a.ORGNAME, a.COMIDNO, b.*" & vbCrLf
                sql &= " ,c.CONTACTEMAIL, c.ZIPCODE, c.ADDRESS,'N' ISBLACK" & vbCrLf
                sql &= " FROM ORG_ORGINFO a" & vbCrLf
                sql &= $" JOIN (SELECT RSID, PLANID, RID, ORGID, RELSHIP, ORGLEVEL, DISTID FROM AUTH_RELSHIP WHERE PLANID=0 OR (PLANID={planlist_SelectVal} AND PlanID <> 0) ) b ON a.ORGID=b.ORGID" & vbCrLf
                sql &= " JOIN ORG_ORGPLANINFO c ON b.RSID=c.RSID" & vbCrLf

                If ViewState("SEARCH1") <> "" Then
                    Dim v_SEARCH1 As String = TIMS.ClearSQM(ViewState("SEARCH1"))
                    Select Case strMaxOrgLevel
                        Case "1", "2"
                            sql &= " WHERE (1!=1 OR b.ORGLEVEL IN (0,1) OR (a.ORGNAME LIKE '%' + @Search1 + '%' AND b.ORGLEVEL <= @strMaxOrgLevel))" & vbCrLf
                            parms.Add("Search1", v_SEARCH1)
                            parms.Add("strMaxOrgLevel", strMaxOrgLevel)
                            'objAdapter.SelectCommand.Parameters.Add("@Search1", SqlDbType.NVarChar).Value=ViewState("SEARCH1")
                            'objAdapter.SelectCommand.Parameters.Add("@strMaxOrgLevel", SqlDbType.VarChar).Value=strMaxOrgLevel
                        Case "3"
                            sql &= " WHERE (1!=1 OR b.ORGLEVEL IN (0,1,2) OR (a.ORGNAME LIKE '%' + @Search1 + '%' AND b.OrgLevel <= @strMaxOrgLevel))" & vbCrLf
                            parms.Add("Search1", v_SEARCH1)
                            parms.Add("strMaxOrgLevel", strMaxOrgLevel)
                            'objAdapter.SelectCommand.Parameters.Add("@Search1", SqlDbType.NVarChar).Value=ViewState("SEARCH1")
                            'objAdapter.SelectCommand.Parameters.Add("@strMaxOrgLevel", SqlDbType.VarChar).Value=strMaxOrgLevel
                        Case Else
                            sql &= " WHERE (1!=1 OR b.ORGLEVEL IN (0,1,2,3) OR (a.ORGNAME LIKE '%' + @Search1 + '%' AND b.ORGLEVEL <= 0))" & vbCrLf
                            parms.Add("Search1", v_SEARCH1)
                            'objAdapter.SelectCommand.Parameters.Add("@Search1", SqlDbType.NVarChar).Value=ViewState("SEARCH1")
                    End Select
                End If
                sql &= vbCrLf
                objstr = sql
                sql = Nothing
            Else
                Dim sql As String = ""
                '有選擇計畫
                If rqPlanID <> "" Then
                    If Not TIMS.IsNumberStr(rqPlanID) Then Return
                    strMaxOrgLevel = TIMS.ClearSQM(TIMS.Get_MaxOrgLevel(rqPlanID, objConn))
                    sql = ""
                    sql &= " SELECT a.ORGNAME, a.COMIDNO, b.*, c.CONTACTEMAIL, c.ZIPCODE, c.ADDRESS,'N' ISBLACK" & vbCrLf
                    sql &= " FROM ORG_ORGINFO a" & vbCrLf
                    'edit by nick 依傳回年度選擇 '有選擇計畫
                    If rqSelected_year <> "" Then
                        If Not TIMS.IsNumberStr(rqSelected_year) Then Return
                        '有選擇計畫，有選擇年度，以年度、主計畫為主 子計畫排除
                        sql &= " JOIN (SELECT RSID, PLANID, RID, ORGID, RELSHIP, ORGLEVEL, DISTID" & vbCrLf
                        sql &= " FROM AUTH_RELSHIP" & vbCrLf
                        sql &= " WHERE PLANID=0 OR (PLANID IN (SELECT DISTINCT vp.PLANID" & vbCrLf
                        sql &= "    FROM VIEW_LOGINPLAN vp" & vbCrLf
                        sql &= "    JOIN VIEW_RIDNAME vr ON vp.PLANID=vr.PLANID" & vbCrLf
                        sql &= $" WHERE vp.YEARS='{rqSelected_year}' AND vp.TPLANID='{sm.UserInfo.TPlanID}' AND vr.RID LIKE '{sm.UserInfo.RID}%' AND vp.PLANID <> 0)" & vbCrLf
                    Else
                        '有選擇計畫
                        sql &= " 	JOIN (SELECT RSID, PLANID, RID, ORGID, RELSHIP, ORGLEVEL, DISTID" & vbCrLf
                        sql &= " 	FROM AUTH_RELSHIP" & vbCrLf
                        sql &= $" 	WHERE PLANID=0 OR (PLANID={rqPlanID} AND PLANID <> 0)" & vbCrLf
                    End If
                    sql &= " ) b ON a.ORGID=b.ORGID" & vbCrLf
                    sql &= " JOIN ORG_ORGPLANINFO c ON b.RSID=c.RSID" & vbCrLf
                Else
                    If Not TIMS.IsNumberStr(sm.UserInfo.PlanID) Then Return
                    strMaxOrgLevel = TIMS.ClearSQM(TIMS.Get_MaxOrgLevel(sm.UserInfo.PlanID, objConn))
                    'edit by nick 依傳回年度選擇
                    If rqSelected_year <> "" Then '有選擇年度
                        If Not TIMS.IsNumberStr(rqSelected_year) Then Return
                        '用Nothing 與空白來做顯示方式(OrgName2,OrgName)判斷
                        ViewState("ORGNAME2") = ""
                        sql = ""
                        sql &= " SELECT dbo.NVL(a.ORGNAME+'('+ip.SEQ+')',a.ORGNAME) ORGNAME2,a.ORGNAME ,a.COMIDNO" & vbCrLf
                        sql &= " ,b.* ,c.CONTACTEMAIL ,c.ZIPCODE ,c.ADDRESS,'N' ISBLACK" & vbCrLf
                        sql &= " FROM ORG_ORGINFO a" & vbCrLf
                        sql &= " JOIN (SELECT RSID, PlanID, RID, OrgID, Relship, OrgLevel, DistID FROM Auth_Relship" & vbCrLf
                        sql &= " WHERE PLANID=0 OR (PLANID IN (SELECT DISTINCT vp.PLANID" & vbCrLf
                        sql &= " FROM VIEW_LOGINPLAN vp" & vbCrLf
                        sql &= " JOIN VIEW_RIDNAME vr on vp.PLANID=vr.PLANID" & vbCrLf
                        sql &= $" WHERE vp.YEARS='{rqSelected_year}' AND vp.TPLANID='{sm.UserInfo.TPlanID}' AND vr.RID LIKE '{sm.UserInfo.RID}%' AND vp.PLANID <> 0))) b ON a.ORGID=b.ORGID" & vbCrLf
                        sql &= " JOIN ORG_ORGPLANINFO c ON b.RSID=c.RSID" & vbCrLf
                        sql &= " LEFT JOIN ID_PLAN ip ON b.PlanID=ip.PLANID" & vbCrLf
                    Else '未選擇年度
                        sql = ""
                        sql &= " SELECT a.ORGNAME, a.COMIDNO, b.*, c.CONTACTEMAIL, c.ZIPCODE, c.ADDRESS,'N' ISBLACK" & vbCrLf
                        sql &= " FROM ORG_ORGINFO a" & vbCrLf
                        sql &= " JOIN (SELECT RSID, PLANID, RID, ORGID, RELSHIP, ORGLEVEL, DISTID" & vbCrLf
                        sql &= " FROM AUTH_RELSHIP" & vbCrLf
                        sql &= $" WHERE PLANID=0 OR (PLANID={sm.UserInfo.PlanID} AND PLANID <> 0) ) b ON a.ORGID=b.ORGID" & vbCrLf
                        sql &= " JOIN ORG_ORGPLANINFO c ON b.RSID=c.RSID" & vbCrLf
                    End If
                End If

                If ViewState("SEARCH1") <> "" Then
                    Dim v_SEARCH1 As String = TIMS.ClearSQM(ViewState("SEARCH1"))
                    Select Case strMaxOrgLevel
                        Case "1", "2"
                            sql &= " WHERE ( 1!=1 OR b.ORGLEVEL IN (0,1) OR (a.ORGNAME LIKE '%' + @Search1 + '%' AND b.ORGLEVEL <= @strMaxOrgLevel))" & vbCrLf
                            parms.Add("Search1", v_SEARCH1)
                            parms.Add("strMaxOrgLevel", strMaxOrgLevel)
                        Case "3"
                            sql &= " WHERE ( 1!=1 OR b.ORGLEVEL IN (0,1,2) OR (a.ORGNAME LIKE '%' + @Search1 + '%' AND b.ORGLEVEL <= @strMaxOrgLevel))" & vbCrLf
                            parms.Add("Search1", v_SEARCH1)
                            parms.Add("strMaxOrgLevel", strMaxOrgLevel)
                        Case Else
                            sql &= " WHERE ( 1!=1 OR b.ORGLEVEL IN (0,1,2,3) OR (a.ORGNAME LIKE '%' + @Search1 + '%' AND b.ORGLEVEL <= 0))" & vbCrLf
                            parms.Add("Search1", v_SEARCH1)
                    End Select
                End If
                sql &= vbCrLf
                objstr = sql
                sql = Nothing
                '特殊規則：計畫青年職涯 轄區為泰山，以署(局)的角度搜尋資料
                'If sm.UserInfo.DistID="002" And sm.UserInfo.TPlanID="36" Then
                '    TableSearch.Visible=False '可使用搜尋關鍵字(關閉)
                '    objstr=""
                '    objstr += " SELECT a.ORGNAME, a.COMIDNO, b.*, c.CONTACTEMAIL ,c.ZIPCODE, c.ADDRESS" & vbCrLf
                '    objstr += " ,'N' ISBLACK" & vbCrLf
                '    objstr += " FROM ORG_ORGINFO a "
                '    objstr += " JOIN (SELECT * FROM AUTH_RELSHIP WHERE PLANID=0 OR (PLANID=" & planSelectid & " AND PLANID <> 0)) b ON a.ORGID=b.ORGID "
                '    objstr += " JOIN ORG_ORGPLANINFO c ON b.RSID=c.RSID "
                'End If
            End If
            Dim dtObj1 As DataTable = DbAccess.GetDataTable(objstr, objConn, parms)
            'TIMS.Fill(objstr, objAdapter, objtable)
            'If objtable Is Nothing Then Return
            'If fg_test_ISBLACK Then
            '    For Each odr As DataRow In dtObj1.Rows
            '        odr("ISBLACK")="Y"
            '    Next
            'End If

            '該登入帳號在黑名單中。
            If hid_ORGBLACKLIST.Value.Equals("Y") Then
                If TIMS.dtHaveDATA(dtObj1) Then
                    '檢測黑名單機構
                    For Each odr As DataRow In dtObj1.Rows
                        Dim v_COMIDNO As String = TIMS.ClearSQM(odr("COMIDNO"))
                        Dim v_DistID As String = TIMS.ClearSQM(odr("DistID"))
                        If dtOrgBlack.Select("ISBLACK='Y' AND COMIDNO='" & v_COMIDNO & "' AND OBTERMS <> '38'").Length > 0 Then
                            odr("ISBLACK") = "Y"
                        Else
                            If dtOrgBlack.Select("ISBLACK='Y' AND COMIDNO='" & v_COMIDNO & "' AND OBTERMS='38' AND DISTID='" & v_DistID & "'").Length > 0 Then
                                odr("ISBLACK") = "Y"
                            End If
                        End If
                    Next
                    'objtable.AcceptChanges()
                End If
            End If

            RID.Value = TIMS.ClearSQM(RID.Value)
            If (sm.UserInfo.DistID <> "000" And sm.UserInfo.RoleID = "1" And rqGetOther = "1") Then
                strFilter = "RID='" & RID.Value & "'"  '署(局)
            Else
                strFilter = "RID='" & sm.UserInfo.RID & "'"  '分署(中心)
                '特殊規則：計畫青年職涯 轄區為泰山，以署(局)的角度搜尋資料
                If sm.UserInfo.DistID = "002" AndAlso sm.UserInfo.TPlanID = "36" Then strFilter = "RID='" & RID.Value & "'"
            End If

            For Each dr In dtObj1.Select(strFilter)
                AddTreeNodes(dr, dtObj1, TreeView1, Nothing)
            Next
            'treeview end
            TreeView1.ExpandAll()  '以程式來展開節點
        End If
    End Sub

    'TreeView1 開始 SQL 語法組合位置
    Private Sub Planlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles planlist.SelectedIndexChanged
        Planlist_Changed()
    End Sub

    '回傳值
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim fg_CookieTable_CanSave As Boolean = False

        Select Case sm.UserInfo.LID
            Case "0"
                fg_CookieTable_CanSave = True
            Case Else
                '確認RID 業務權限 與PLANID 計畫 為登入權限 sm.UserInfo.PlanID 
                fg_CookieTable_CanSave = TIMS.CheckRIDsPLAN(Me, OrgRID.Value, objConn)
        End Select

        'tb: COOKIE_DATA 取得快取資料
        Dim rqGetOther As String = TIMS.ClearSQM(Request("GetOther"))
        If fg_CookieTable_CanSave Then
            If rqGetOther <> "1" Then
                dt = TIMS.GetCookieTable(Me, da, objConn)
                For i As Integer = 1 To 5
                    If dt.Select("ITEMNAME='LevOrgRID" & i & "'").Length = 0 Then
                        Dim fg_InsertFlag As Boolean = True
                        For j As Integer = 1 To 5
                            If dt.Select("ITEMNAME='LevOrgRID" & j & "' AND ITEMVALUE='" & OrgRID.Value & "'").Length <> 0 Then
                                fg_InsertFlag = False
                                Exit For
                            End If
                        Next
                        If fg_InsertFlag = True Then
                            TIMS.InsertCookieTable(Me, dt, da, "LevOrgName" & i, OrgName.Value, False, objConn)
                            TIMS.InsertCookieTable(Me, dt, da, "LevOrgRID" & i, OrgRID.Value, True, objConn)
                        End If
                        Exit For
                    Else
                        If i = 5 Then
                            For j As Integer = 1 To 4
                                Dim NewDr As DataRow
                                Dim OldDr As DataRow

                                NewDr = dt.Select("ITEMNAME='LevOrgRID" & j + 1 & "'")(0)
                                OldDr = dt.Select("ITEMNAME='LevOrgRID" & j & "'")(0)
                                OldDr("ITEMVALUE") = NewDr("ITEMVALUE")

                                NewDr = dt.Select("ITEMNAME='LevOrgName" & j + 1 & "'")(0)
                                OldDr = dt.Select("ITEMNAME='LevOrgName" & j & "'")(0)
                                OldDr("ITEMVALUE") = NewDr("ITEMVALUE")
                            Next
                            Dim fg_InsertFlag As Boolean = True
                            For j As Integer = 1 To 5
                                If dt.Select("ITEMNAME='LevOrgRID" & j & "' AND ITEMVALUE='" & OrgRID.Value & "'").Length <> 0 Then
                                    fg_InsertFlag = False
                                    Exit For
                                End If
                            Next
                            If fg_InsertFlag = True Then
                                If OrgName.Value <> "" AndAlso OrgRID.Value <> "" Then
                                    TIMS.InsertCookieTable(Me, dt, da, "LevOrgName" & i, OrgName.Value, False, objConn)
                                    TIMS.InsertCookieTable(Me, dt, da, "LevOrgRID" & i, OrgRID.Value, True, objConn)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If

        Dim dr As DataRow = Nothing
        '(20180703 因在點選完訓練機構後,會莫名的發生錯誤,所以暫時先將下方的程式碼先註解掉。再依後續情況,重新調整程式碼)
        'dr=TIMS.GET_OnlyOne_OCID(OrgRID.Value) '判斷機構是否只有一個班級

        If dr IsNot Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                Common.RespWrite(Me, "<script language=javascript>")
                Common.RespWrite(Me, "function GetValue(){")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMID1')!=null) window.opener.document.form1.TMID1.value='" & dr("trainname") & "';")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMIDValue1')!=null) window.opener.document.form1.TMIDValue1.value='" & dr("trainid") & "';")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCID1')!=null) window.opener.document.form1.OCID1.value='" & dr("classname") & "';")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCIDValue1')!=null) window.opener.document.form1.OCIDValue1.value='" & dr("ocid") & "';")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('TDate')!=null) window.opener.document.form1.TDate.value='" & dr("FTDate") & "';")
                Common.RespWrite(Me, "    window.close();")
                Common.RespWrite(Me, "}")
                Common.RespWrite(Me, "GetValue();")
                hidbtnName.Value = TIMS.ClearSQM(hidbtnName.Value)
                If hidbtnName.Value <> "" Then Common.RespWrite(Me, "    if (window.opener.document.getElementById('" & hidbtnName.Value & "')) window.opener.document.getElementById('" & hidbtnName.Value & "').click();")
                Common.RespWrite(Me, "</script>")
            Else  '不只一個班級
                Common.RespWrite(Me, "<script language=javascript>")
                Common.RespWrite(Me, "function Value2(){")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMID1')!=null) window.opener.document.form1.TMID1.value='';")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMIDValue1')!=null) window.opener.document.form1.TMIDValue1.value='';")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCID1')!=null) window.opener.document.form1.OCID1.value='';")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCIDValue1')!=null) window.opener.document.form1.OCIDValue1.value='';")
                Common.RespWrite(Me, "    if(window.opener.document.getElementById('TDate')!=null) window.opener.document.form1.TDate.value ='';")
                Common.RespWrite(Me, "    window.close();")
                Common.RespWrite(Me, "}")
                Common.RespWrite(Me, "Value2();")
                Common.RespWrite(Me, "</script>")
            End If
        Else
            Common.RespWrite(Me, "<script language=javascript>")
            Common.RespWrite(Me, "function Value3(){")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMID1')!=null) window.opener.document.form1.TMID1.value='';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TMIDValue1')!=null) window.opener.document.form1.TMIDValue1.value='';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCID1')!=null) window.opener.document.form1.OCID1.value='';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('OCIDValue1')!=null) window.opener.document.form1.OCIDValue1.value='';")
            Common.RespWrite(Me, "    if(window.opener.document.getElementById('TDate')!=null) window.opener.document.form1.TDate.value ='';")
            Common.RespWrite(Me, "    window.close();")
            Common.RespWrite(Me, "}")
            Common.RespWrite(Me, "Value3();")
            Common.RespWrite(Me, "</script>")
        End If
        Common.RespWrite(Me, "<script>window.close();</script>")
    End Sub

    Sub Downyear_Changed()
        planlist.Items.Clear()
        TreeView1.Nodes.Clear()

        Dim vDownyear As String = TIMS.GetListValue(Downyear)
        Dim vDistID As String = TIMS.GetListValue(DistID)
        If vDownyear = "" OrElse vDistID = "" Then Return
        If Downyear.SelectedIndex <> 0 AndAlso vDownyear <> "" Then
            Dim parms As New Hashtable From {{"YEARS", vDownyear}, {"DISTID", vDistID}}
            Dim sql As String = ""
            sql &= " SELECT CONCAT(a.YEARS,c.NAME,b.PLANNAME,a.SEQ) PLANNAME ,a.PLANID,a.TPLANID" & vbCrLf
            sql &= " FROM ID_PLAN a" & vbCrLf
            sql &= " JOIN KEY_PLAN b ON a.TPLANID=b.TPLANID" & vbCrLf
            sql &= " JOIN ID_DISTRICT c ON a.DISTID=c.DISTID" & vbCrLf
            sql &= " WHERE a.YEARS=@YEARS AND c.DISTID=@DISTID" & vbCrLf
            sql &= " ORDER BY 1" & vbCrLf
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objConn, parms)

            Dim vsPLANID As String = ""
            If $"{sm.UserInfo.TPlanID}" <> "" Then
                Dim ff3 As String = $"TPLANID='{sm.UserInfo.TPlanID}'"
                If dt.Select(ff3).Length > 0 Then vsPLANID = $"{dt.Select(ff3)(0)("PLANID")}"
            End If

            If TIMS.dtNODATA(dt) Then
                planlist.Items.Insert(0, New ListItem("==查無資料==", ""))
            Else
                planlist.DataSource = dt
                planlist.DataTextField = "PLANNAME"
                planlist.DataValueField = "PLANID"
                planlist.DataBind()
                planlist.Items.Insert(0, New ListItem("==請選擇==", ""))
                If dt.Rows.Count = 1 Then
                    If Not planlist.SelectedItem Is Nothing Then planlist.SelectedItem.Selected = False
                    planlist.Items(1).Selected = True
                    Planlist_Changed()
                ElseIf vsPLANID <> "" Then
                    Common.SetListItem(planlist, vsPLANID)
                    Planlist_Changed()
                End If
            End If
        Else
            planlist.Items.Insert(0, New ListItem("==請選擇年度==", ""))
        End If
    End Sub

    '加入年度 by nick 060301
    Private Sub Downyear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Downyear.SelectedIndexChanged
        Downyear_Changed()
    End Sub

    '隱藏式 Client 按鈕(btnSearch_Click)
    Private Sub HbtnSearch_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hbtnSearch.ServerClick
        Planlist_Changed()  '重新搜尋工作
        TreeView1.ExpandAll()  '以程式設計方式展開節點
    End Sub
End Class
