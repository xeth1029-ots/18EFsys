Partial Class SYS_06_003
    Inherits AuthBasePage

    Const Cst_EmptySelValue As String = TIMS.cst_ddl_PleaseChoose3
    Const cst_sql As String = "sql" 'sql語法使用
    Const cst_save As String = "save" 'save時使用
    Const cst_addnew As String = "addnew" '新增
    Const cst_update As String = "update" '修改
    Const cst_save_正式 As String = "Y"
    Const cst_save_草稿 As String = "N"
    Const cst_objtype_txt As String = "訓練單位,報名者,在訓學員,結訓學員,曾受訓的所有學員" '由1開始 用逗號分隔
    '1:訓練單位/2:報名者/3:在訓學員/4:結訓學員/5:曾受訓的所有學員
    Const cst_vs_sqlObjectType As String = "sqlObjectType"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        sm = SessionModel.Instance()
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = Me.DataGrid1

        If Not Me.IsPostBack Then
            Call Create1()

            DataGridTable1.Visible = False '預設搜尋資料不顯示

            panelSearch.Visible = True '搜尋功能啟動
            PanelEdit1.Visible = False '修改功能關閉
        End If
    End Sub

    Sub Create1()
        cblobjecttype_s.Items.Clear() '查詢用
        cblobjecttype.Items.Clear() '儲存用
        ddlplanlist_s.Items.Clear()
        ddlplanlist.Items.Clear()

        Dim objtypeA() As String = cst_objtype_txt.Split(",")
        For i As Integer = 0 To objtypeA.Length - 1
            cblobjecttype_s.Items.Add(New ListItem(Convert.ToString(objtypeA(i)), CStr(i + 1)))
            cblobjecttype.Items.Add(New ListItem(Convert.ToString(objtypeA(i)), CStr(i + 1)))
        Next

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select planid " & vbCrLf
        sql &= " ,years+distname+planname+seq planname " & vbCrLf
        sql &= " FROM VIEW_PLAN WITH(NOLOCK)" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and distid ='" & sm.UserInfo.DistID & "' " & vbCrLf
        sql &= " and years ='" & sm.UserInfo.Years & "'" & vbCrLf
        DbAccess.MakeListItem(ddlplanlist_s, sql, objconn)
        ddlplanlist_s.Items.Insert(0, New ListItem(Cst_EmptySelValue, ""))

        DbAccess.MakeListItem(ddlplanlist, sql, objconn)
        ddlplanlist.Items.Insert(0, New ListItem(Cst_EmptySelValue, ""))

    End Sub

    '設定新增初值
    Sub sUtl_ClearPanelEdit1(ByVal sType As String)
        Hid_MSID.Value = ""

        txtSubject.Text = ""
        SendDate.Text = ""
        rblPlanID.SelectedIndex = -1
        ddlplanlist.SelectedIndex = -1
        For i As Integer = 0 To cblobjecttype.Items.Count - 1
            cblobjecttype.Items(i).Selected = False
        Next
        txtcontents.Text = ""
        labIsApprPaper.Text = "" 'vIsApprPaper
        txtAcctEmail.Text = "" 'Convert.ToString(dr("AcctEmail"))

        'hidmsid.Value = ""
        'txtSubject.Text = ""
        'SendDate.Text = ""
        If sType = cst_addnew Then
            '搜尋->新增
            If rblPlanID_s.SelectedValue <> "" Then
                Common.SetListItem(rblPlanID, rblPlanID_s.SelectedValue)
            End If
            '搜尋->新增
            If ddlplanlist_s.SelectedValue <> "" Then
                Common.SetListItem(ddlplanlist, ddlplanlist_s.SelectedValue)
            End If
            '搜尋->新增
            Dim sValue1 As String = TIMS.GetCblValue(cblobjecttype_s)
            Call TIMS.SetCblValue(cblobjecttype, sValue1)
            'For i As Integer = 0 To cblobjecttype_s.Items.Count - 1
            '    If cblobjecttype_s.Items(i).Selected Then
            '        cblobjecttype.Items(i).Selected = True
            '    End If
            'Next
        End If

    End Sub

    Sub sUtl_ShowPanelEdit1(ByVal sSearchW As String)
        panelSearch.Visible = False '搜尋功能關閉
        PanelEdit1.Visible = True '修改功能啟動

        '修改
        'sSearchW = e.CommandArgument
        Hid_MSID.Value = TIMS.GetMyValue(sSearchW, "MSID")

        Dim dr As DataRow = schDataRow1(Hid_MSID.Value) '搜尋顯示
        If dr Is Nothing Then Exit Sub

        txtSubject.Text = Convert.ToString(dr("Subject"))
        SendDate.Text = Convert.ToString(dr("SendDate"))
        SendDate.Text = TIMS.Cdate3(SendDate.Text)

        If Convert.ToString(dr("PlanID")) <> "" Then
            Select Case Convert.ToString(dr("PlanID"))
                Case "0" '0全計畫
                    Common.SetListItem(rblPlanID, "0")
                    ddlplanlist.SelectedIndex = -1

                Case Else '非0數字 但也非1 
                    Common.SetListItem(rblPlanID, "1")
                    Common.SetListItem(ddlplanlist, Convert.ToString(dr("PlanID")))

            End Select

            If Convert.ToString(dr("PlanID")) = "1" Then
                For i As Integer = 0 To cblobjecttype.Items.Count - 1
                    cblobjecttype.Items(i).Selected = True
                Next
            End If
        End If

        If Convert.ToString(dr("objectType")) <> "" Then
            For i As Integer = 0 To cblobjecttype.Items.Count - 1
                If TIMS.Chk_SplitValeu1(Convert.ToString(dr("objectType")), cblobjecttype.Items(i).Value) Then
                    cblobjecttype.Items(i).Selected = True
                End If
            Next
        End If
        txtcontents.Text = dr("contents").ToString

        btnSave1.Visible = True
        btnSave2.Visible = True
        Dim vIsApprPaper As String = "草稿"
        If Convert.ToString(dr("IsApprPaper")) = "Y" Then vIsApprPaper = "正式"
        labIsApprPaper.Text = vIsApprPaper
        If Convert.ToString(dr("IsApprPaper")) = "Y" Then
            btnSave2.Visible = False
        End If

        txtAcctEmail.Text = Convert.ToString(dr("AcctEmail"))
    End Sub

    '取得 objectType 發送對象 取得SQL/取得OBJTYPE
    Function sUtl_GetobjectType(ByVal obj As CheckBoxList, ByVal sType As String) As String
        'sType @cst_sql@sql語法使用 / cst_save@save時使用
        Dim rst As String = ""
        rst = ""
        If Not obj Is Nothing Then
            For i As Integer = 0 To obj.Items.Count - 1
                If obj.Items(i).Selected Then
                    Select Case sType
                        Case cst_sql
                            'a.objectType
                            rst += " or ','+a.objectType+',' like '%," & obj.Items(i).Value & ",%'" & vbCrLf
                        Case cst_save
                            If rst <> "" Then rst += ","
                            rst += obj.Items(i).Value
                    End Select
                End If
            Next
        End If

        '後續處理
        Select Case sType
            Case cst_sql
                'sql後續處理
                If rst <> "" Then
                    Dim tmp As String = ""
                    tmp = ""
                    tmp += " and (1!=1 " & vbCrLf
                    tmp += rst & ")" & vbCrLf
                    rst = tmp
                End If
        End Select

        Return rst
    End Function

    '顯示文字
    Function sUtl_ShowobjectType(ByVal sValues As String, ByVal obj As CheckBoxList) As String
        Dim rst As String = ""
        If Not obj Is Nothing Then
            For i As Integer = 0 To obj.Items.Count - 1
                If TIMS.Chk_SplitValeu1(sValues, obj.Items(i).Value) Then
                    If rst <> "" Then rst += ","
                    rst += obj.Items(i).Value & "." & obj.Items(i).Text
                End If
            Next
        End If
        Return rst
    End Function

    'SQL
    Sub sSearch1(ByRef sParms As Hashtable)
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉

        Dim vPlanID As String = TIMS.GetMyValue2(sParms, "PlanID")
        Dim vSqlObjectType As String = TIMS.GetMyValue2(sParms, cst_vs_sqlObjectType)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.msid" & vbCrLf '/*PK*/
        sql &= " ,a.Subject" & vbCrLf
        sql &= " ,CONVERT(varchar, a.SendDate, 111) SendDate" & vbCrLf
        sql &= " ,a.PlanID " & vbCrLf
        sql &= " ,case when a.PlanID=0 then N'全系統' else ip.years+ip.distname+ip.planname+ip.seq end PlanName " & vbCrLf
        sql &= " ,a.objectType" & vbCrLf
        sql &= " ,a.SendState" & vbCrLf
        sql &= " ,a.contents" & vbCrLf
        sql &= " ,a.IsApprPaper" & vbCrLf
        sql &= " ,a.ACCTEMAIL" & vbCrLf
        sql &= " FROM SYS_MAILSEND a " & vbCrLf
        sql &= " LEFT JOIN VIEW_PLAN ip on ip.planid =a.planid " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.Years =@YEARS " & vbCrLf
        sql &= " AND a.DistID =@DistID " & vbCrLf
        If vPlanID <> "" Then sql &= " AND a.PlanID =@PlanID " & vbCrLf
        If vSqlObjectType <> "" Then sql &= vSqlObjectType

        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("YEARS", sm.UserInfo.Years)
        parms.Add("DistID", sm.UserInfo.DistID)
        If vPlanID <> "" Then parms.Add("PlanID", vPlanID)

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        DataGridTable1.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    'SQL return datarow
    Function schDataRow1(ByVal vMSID As String) As DataRow
        Dim rst As DataRow = Nothing

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.msid" & vbCrLf '/*PK*/
        sql &= " ,a.Subject" & vbCrLf
        sql &= " ,CONVERT(varchar, a.SendDate, 111) SendDate" & vbCrLf
        sql &= " ,a.PlanID" & vbCrLf
        sql &= " ,a.objectType" & vbCrLf
        sql &= " ,a.SendState" & vbCrLf
        sql &= " ,a.contents" & vbCrLf
        'Sql += " ,ModifyAcct" & vbCrLf
        'Sql += " ,ModifyDate" & vbCrLf
        sql &= " ,a.IsApprPaper " & vbCrLf
        sql &= " ,a.ACCTEMAIL" & vbCrLf
        sql &= " FROM Sys_MailSend a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.Years =@YEARS " & vbCrLf
        sql &= " AND a.DistID =@DistID " & vbCrLf
        sql &= " and a.MSID =@MSID" & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("YEARS", sm.UserInfo.Years)
        parms.Add("DistID", sm.UserInfo.DistID)
        parms.Add("MSID", Val(vMSID))

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then rst = dt.Rows(0)

        Return rst
    End Function

    '設定 ViewState Value
    Function sUtl_GetViewStateValue() As Hashtable
        Dim sParms As New Hashtable '回傳參數
        sParms.Clear()

        '發送範圍 (全系統 / 指定計畫)
        Select Case Me.rblPlanID_s.SelectedValue
            Case "0"
                '發送範圍 (全系統)
                'Me.ViewState("PlanID") = "0"
                sParms.Add("PlanID", "0")
            Case Else '"1"
                '發送範圍 (指定計畫)
                'Me.ViewState("PlanID") = IIf(Me.ddlplanlist_s.SelectedValue <> "", Me.ddlplanlist_s.SelectedValue, "")
                sParms.Add("PlanID", IIf(Me.ddlplanlist_s.SelectedValue <> "", Me.ddlplanlist_s.SelectedValue, ""))
        End Select
        '發送對象
        'Me.ViewState(vs_sqlObjectType) = sUtl_GetobjectType(cblobjecttype_s, cst_sql)
        sParms.Add(cst_vs_sqlObjectType, sUtl_GetobjectType(cblobjecttype_s, cst_sql))
        Return sParms
    End Function

    '儲存 動作
    Function sUtl_SaveData1(ByVal sParms As Hashtable) As Integer
        Dim rst As Integer = 0 '0:異常儲存 1:正常儲存

        Dim sType As String  '儲存狀態 (cst_addnew / cst_update )
        sType = cst_update '修改
        If Hid_MSID.Value = "" Then
            sType = cst_addnew '新增
        End If
        'Dim sql As String = ""
        '確認
        'sSearchW = e.CommandArgument
        Dim sPlanID As String = TIMS.GetMyValue2(sParms, "PlanID")
        Dim vObjecttype As String = TIMS.GetMyValue2(sParms, "objecttype")
        Dim vIsApprPaper As String = TIMS.GetMyValue2(sParms, "IsApprPaper")

        Dim vSubject As String = TIMS.ClearSQM(txtSubject.Text)
        Dim vSendDate As String = TIMS.Cdate3(SendDate.Text)
        Dim vcontents As String = TIMS.ClearSQM(txtcontents.Text)
        Dim vAcctEmail As String = TIMS.ClearSQM(txtAcctEmail.Text)

        'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
        Dim sql As String = ""
        Select Case sType
            Case cst_addnew '新增
                Dim iMSID As Integer = DbAccess.GetNewId(objconn, "SYS_MAILSEND_MSID_SEQ,SYS_MAILSEND,MSID")
                'sql &= "  /* IDENTITY(1, 1) : msid */ " & vbCrLf
                sql = "" & vbCrLf
                sql &= " INSERT INTO SYS_MAILSEND(MSID" & vbCrLf
                sql &= " ,Subject" & vbCrLf
                sql &= " ,SendDate" & vbCrLf
                sql &= " ,PlanID" & vbCrLf
                sql &= " ,objectType" & vbCrLf
                sql &= " ,SendState" & vbCrLf
                sql &= " ,contents" & vbCrLf
                sql &= " ,DistID" & vbCrLf
                sql &= " ,Years" & vbCrLf
                sql &= " ,ModifyAcct" & vbCrLf
                sql &= " ,ModifyDate" & vbCrLf
                sql &= " ,IsApprPaper" & vbCrLf
                sql &= " ,ACCTEMAIL" & vbCrLf
                sql &= " ) VALUES (@MSID" & vbCrLf
                sql &= " ,@Subject" & vbCrLf
                sql &= " ,@SendDate" & vbCrLf
                sql &= " ,@PlanID" & vbCrLf
                sql &= " ,@objectType" & vbCrLf
                sql &= " ,NULL" & vbCrLf
                sql &= " ,@contents" & vbCrLf
                sql &= " ,@DistID" & vbCrLf
                sql &= " ,@Years" & vbCrLf
                sql &= " ,@ModifyAcct" & vbCrLf
                sql &= " ,getdate()" & vbCrLf
                sql &= " ,@IsApprPaper" & vbCrLf
                sql &= " ,@ACCTEMAIL" & vbCrLf
                sql &= " ) " & vbCrLf

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("MSID", iMSID)
                parms.Add("Subject", vSubject)
                parms.Add("SendDate", vSendDate)
                parms.Add("PlanID", sPlanID)
                parms.Add("objectType", vObjecttype)
                parms.Add("contents", vcontents)
                parms.Add("DistID", sm.UserInfo.DistID)
                parms.Add("Years", sm.UserInfo.Years)
                parms.Add("ModifyAcct", sm.UserInfo.UserID)
                parms.Add("IsApprPaper", vIsApprPaper)
                parms.Add("ACCTEMAIL", vAcctEmail)
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
                rst = 1 '儲存成功

            Case Else 'cst_update '修改
                sql = "" & vbCrLf
                sql &= " UPDATE Sys_MailSend " & vbCrLf
                sql &= " SET Subject= @Subject" & vbCrLf
                sql &= " ,SendDate= @SendDate" & vbCrLf
                sql &= " ,PlanID= @PlanID" & vbCrLf
                sql &= " ,objectType= @objectType" & vbCrLf
                'sql += "  ,SendState= @SendState" & vbCrLf
                sql &= " ,contents= @contents" & vbCrLf
                sql &= " ,DistID= @DistID" & vbCrLf
                sql &= " ,Years= @Years" & vbCrLf
                sql &= " ,ModifyAcct= @ModifyAcct" & vbCrLf
                sql &= " ,ModifyDate=getdate()" & vbCrLf
                sql &= " ,IsApprPaper=@IsApprPaper" & vbCrLf
                sql &= " ,ACCTEMAIL=@ACCTEMAIL" & vbCrLf
                sql &= " WHERE 1=1 " & vbCrLf
                sql &= " AND MSID= @MSID" & vbCrLf

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("Subject", vSubject)
                parms.Add("SendDate", vSendDate)
                parms.Add("PlanID", sPlanID)
                parms.Add("objectType", vObjecttype)
                parms.Add("contents", vcontents)
                parms.Add("DistID", sm.UserInfo.DistID)
                parms.Add("Years", sm.UserInfo.Years)
                parms.Add("ModifyAcct", sm.UserInfo.UserID)
                parms.Add("IsApprPaper", vIsApprPaper)
                parms.Add("ACCTEMAIL", vAcctEmail)
                parms.Add("MSID", Val(Hid_MSID.Value))
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
                rst = 1 '儲存成功

        End Select

        Return rst
    End Function

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Dim vsParms As Hashtable = sUtl_GetViewStateValue()  '設定 ViewState Value
        Call sSearch1(vsParms)
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sSearchW As String = ""

        Select Case e.CommandName
            Case "Edit1"
                '修改
                sSearchW = e.CommandArgument
                Hid_MSID.Value = TIMS.GetMyValue(sSearchW, "MSID")

                Call sUtl_ClearPanelEdit1("")
                Call sUtl_ShowPanelEdit1(sSearchW)

            Case "Delete1"
                '刪除
                sSearchW = e.CommandArgument
                Hid_MSID.Value = TIMS.GetMyValue(sSearchW, "MSID")

                Try
                    '刪除sql
                    Dim sql As String = ""
                    sql = ""
                    sql &= " DELETE Sys_MailSend " & vbCrLf
                    sql &= " WHERE 1=1 " & vbCrLf
                    sql &= " AND MSID= @MSID" & vbCrLf
                    Dim parms As New Hashtable
                    parms.Clear()
                    parms.Add("MSID", Val(Hid_MSID.Value))
                    DbAccess.ExecuteNonQuery(sql, objconn, parms)

                    Call sUtl_ClearPanelEdit1("")
                    Dim vsParms As Hashtable
                    vsParms = sUtl_GetViewStateValue()  '設定 ViewState Value
                    Call sSearch1(vsParms)

                Catch ex As Exception
                    Common.MessageBox(Me, ex.ToString)
                End Try

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbobjectType As Label = e.Item.FindControl("lbobjectType")
                Dim btnEdit1 As LinkButton = e.Item.FindControl("btnEdit1")
                Dim btnDelete1 As LinkButton = e.Item.FindControl("btnDelete1")
                '序號
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex
                '顯示文字
                lbobjectType.Text = Me.sUtl_ShowobjectType(drv("objecttype").ToString, cblobjecttype_s)
                Dim sCmdArg As String = ""
                sCmdArg = ""
                sCmdArg &= "&msid=" & Convert.ToString(drv("MSID"))
                btnEdit1.CommandArgument = sCmdArg
                btnDelete1.CommandArgument = sCmdArg
                btnDelete1.Attributes("onclick") = "return confirm('此動作會刪除此筆資料，是否確定刪除?');"
        End Select
    End Sub

    '新增
    Private Sub btnAddnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddnew.Click
        panelSearch.Visible = False '搜尋功能關閉
        PanelEdit1.Visible = True '修改功能啟動

        Call sUtl_ClearPanelEdit1(cst_addnew)
        'Call sUtl_InsertPanelEdit1()
    End Sub

    '回上一頁
    Private Sub btnBack1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack1.Click
        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉
    End Sub

    Function CheckSaveData1(ByRef sErrmsg As String) As Hashtable
        Dim parms As New Hashtable

        '發送範圍 (全系統 / 指定計畫)
        Dim PlanID As String = ""
        Select Case Me.rblPlanID.SelectedValue
            Case "0"
                '發送範圍 (全系統)
                PlanID = "0"
            Case Else '"1"
                '發送範圍 (指定計畫)
                PlanID = IIf(Me.ddlplanlist.SelectedValue <> "", Me.ddlplanlist.SelectedValue, "")
        End Select
        '發送對象
        Dim v_objecttype As String = TIMS.GetCblValue(cblobjecttype) 'sUtl_GetobjectType(cblobjecttype, cst_save)

        '-------------------------------儲存前檢核-------------------------------
        txtSubject.Text = TIMS.ClearSQM(txtSubject.Text)
        SendDate.Text = TIMS.ClearSQM(SendDate.Text)
        txtcontents.Text = TIMS.ClearSQM(txtcontents.Text)
        'labIsApprPaper.Text = "" 'vIsApprPaper
        txtAcctEmail.Text = TIMS.ClearSQM(txtAcctEmail.Text)

        'Dim sErrmsg As String = ""
        sErrmsg = ""
        If txtSubject.Text = "" Then
            sErrmsg += "請輸入主題" & vbCrLf
        End If
        If SendDate.Text = "" Then
            sErrmsg += "請輸入發送日期" & vbCrLf
        End If
        If PlanID = "" Then
            sErrmsg += "請選擇發送範圍之計畫" & vbCrLf
        End If
        If v_objecttype = "" Then
            sErrmsg += "請選擇發送對象" & vbCrLf
        End If
        If txtcontents.Text = "" Then
            sErrmsg += "請輸入內容" & vbCrLf
        End If
        If txtAcctEmail.Text <> "" Then
            If Not TIMS.CheckEmail(txtAcctEmail.Text) Then
                sErrmsg += "寄送報告EMAIL，格式有誤!" & vbCrLf
            End If
        End If

        Dim IsApprPaper As String = TIMS.ClearSQM(Hid_IsApprPaper.Value)
        'Dim parms As New Hashtable
        parms.Clear()
        parms.Add("PlanID", PlanID)
        parms.Add("objecttype", v_objecttype)
        parms.Add("IsApprPaper", IsApprPaper)
        Return parms
    End Function


    Sub SaveData1()
        Dim sErrmsg As String = ""
        Dim csd_Parms As Hashtable = CheckSaveData1(sErrmsg)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If
        '-------------------------------儲存前檢核-------------------------------

        If sUtl_SaveData1(csd_Parms) = 1 Then
            '正常儲存
            Dim vsParms As Hashtable
            vsParms = sUtl_GetViewStateValue()  '設定 ViewState Value
            Call sSearch1(vsParms)
        End If

    End Sub

    '儲存
    Private Sub btnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave1.Click
        Hid_IsApprPaper.Value = cst_save_正式
        SaveData1()
    End Sub

    Protected Sub btnSave2_Click(sender As Object, e As EventArgs) Handles btnSave2.Click
        Hid_IsApprPaper.Value = cst_save_草稿
        SaveData1()
    End Sub

    Sub SHOW_Query1()
        Dim sErrmsg As String = ""
        Dim csd_Parms As Hashtable = CheckSaveData1(sErrmsg)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        Dim sPlanID As String = TIMS.GetMyValue2(csd_Parms, "PlanID")
        Dim vObjecttype As String = TIMS.GetMyValue2(csd_Parms, "objecttype")
        Dim vIsApprPaper As String = TIMS.GetMyValue2(csd_Parms, "IsApprPaper")
        'Dim org As Integer
        'Const cst_objtype_txt As String = "訓練單位,報名者,在訓學員,結訓學員,曾受訓的所有學員" '由1開始 用逗號分隔
        '1:訓練單位/2:報名者/3:在訓學員/4:結訓學員/5:曾受訓的所有學員
        'Dim sendObjType1 As String = TIMS.GetCblValue(cblobjecttype)
        'Dim sPlanID As String = TIMS.GetMyValue2(sParms, "PlanID")

        Dim parms As New Hashtable
        Dim sql As String = ""

        Dim iCNT_ALL As Integer = 0
        Dim iCNT_1 As Integer = 0
        Dim iCNT_2 As Integer = 0
        Dim iCNT_3 As Integer = 0
        Dim iCNT_4 As Integer = 0
        Dim iCNT_5 As Integer = 0

        Dim objtypeB() As String = vObjecttype.Split(",")
        For i As Integer = 0 To objtypeB.Length - 1
            'cblobjecttype_s.Items.Add(New ListItem(Convert.ToString(objtypeA(i)), CStr(i + 1)))
            'cblobjecttype.Items.Add(New ListItem(Convert.ToString(objtypeA(i)), CStr(i + 1)))
            Select Case Convert.ToString(objtypeB(i))
                Case "1"  '1:訓練單位/2:報名者/3:在訓學員/4:結訓學員/5:曾受訓的所有學員
                    sql = "" & vbCrLf
                    sql &= " SELECT COUNT(CASE WHEN CONTACTEMAIL LIKE '%@%' THEN 1 END) CNT1" & vbCrLf
                    sql &= " FROM VIEW_RIDNAME" & vbCrLf
                    sql &= " WHERE 1=1" & vbCrLf
                    sql &= " AND CONTACTEMAIL IS NOT NULL" & vbCrLf
                    sql &= " AND YEARS =@YEARS" & vbCrLf
                    sql &= " AND DISTID =@DISTID" & vbCrLf
                    If sPlanID <> "0" Then sql &= " AND PLANID =@PLANID" & vbCrLf
                    parms.Clear()
                    parms.Add("YEARS", sm.UserInfo.Years)
                    parms.Add("DISTID", sm.UserInfo.DistID)
                    If sPlanID <> "0" Then parms.Add("PLANID", sPlanID)
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
                    If dr IsNot Nothing Then
                        If Convert.ToString(dr("CNT1")) <> "" Then
                            iCNT_1 = Val(dr("CNT1"))
                            iCNT_ALL += iCNT_1
                        End If
                    End If

                Case "2"  '1:訓練單位/2:報名者/3:在訓學員/4:結訓學員/5:曾受訓的所有學員
                    sql = "" & vbCrLf
                    sql &= " SELECT COUNT(CASE WHEN EMAIL LIKE '%@%' THEN 1 END) CNT1" & vbCrLf
                    sql &= " FROM V_ENTERTYPE2" & vbCrLf
                    sql &= " WHERE 1=1" & vbCrLf
                    sql &= " AND EMAIL IS NOT NULL" & vbCrLf
                    sql &= " AND YEARS =@YEARS" & vbCrLf
                    sql &= " AND DISTID =@DISTID" & vbCrLf
                    If sPlanID <> "0" Then sql &= " AND PLANID =@PLANID" & vbCrLf
                    parms.Clear()
                    parms.Add("YEARS", sm.UserInfo.Years)
                    parms.Add("DISTID", sm.UserInfo.DistID)
                    If sPlanID <> "0" Then parms.Add("PLANID", sPlanID)
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)

                    If dr IsNot Nothing Then
                        If Convert.ToString(dr("CNT1")) <> "" Then
                            iCNT_2 = Val(dr("CNT1"))
                            iCNT_ALL += iCNT_2
                        End If
                    End If

                Case "3"  '1:訓練單位/2:報名者/3:在訓學員/4:結訓學員/5:曾受訓的所有學員
                    sql = "" & vbCrLf
                    sql &= " SELECT COUNT(CASE WHEN EMAIL LIKE '%@%' THEN 1 END) CNT1" & vbCrLf
                    sql &= " FROM V_STUDENTINFO" & vbCrLf
                    sql &= " WHERE 1=1" & vbCrLf
                    sql &= " AND STUDSTATUS NOT IN (2,3)" & vbCrLf
                    sql &= " AND STDATE<=GETDATE()" & vbCrLf
                    sql &= " AND FTDATE>=GETDATE()" & vbCrLf
                    sql &= " AND EMAIL IS NOT NULL" & vbCrLf
                    sql &= " AND YEARS =@YEARS" & vbCrLf
                    sql &= " AND DISTID =@DISTID" & vbCrLf
                    If sPlanID <> "0" Then sql &= " AND PLANID =@PLANID" & vbCrLf
                    parms.Clear()
                    parms.Add("YEARS", sm.UserInfo.Years)
                    parms.Add("DISTID", sm.UserInfo.DistID)
                    If sPlanID <> "0" Then parms.Add("PLANID", sPlanID)
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)

                    If dr IsNot Nothing Then
                        If Convert.ToString(dr("CNT1")) <> "" Then
                            iCNT_3 = Val(dr("CNT1"))
                            iCNT_ALL += iCNT_3
                        End If
                    End If

                Case "4"  '1:訓練單位/2:報名者/3:在訓學員/4:結訓學員/5:曾受訓的所有學員
                    sql = "" & vbCrLf
                    sql &= " SELECT COUNT(CASE WHEN EMAIL LIKE '%@%' THEN 1 END) CNT1" & vbCrLf
                    sql &= " FROM V_STUDENTINFO" & vbCrLf
                    sql &= " WHERE 1=1" & vbCrLf
                    sql &= " AND STUDSTATUS NOT IN (2,3)" & vbCrLf
                    sql &= " AND FTDATE<=GETDATE()" & vbCrLf
                    sql &= " AND EMAIL IS NOT NULL" & vbCrLf
                    sql &= " AND YEARS =@YEARS" & vbCrLf
                    sql &= " AND DISTID =@DISTID" & vbCrLf
                    If sPlanID <> "0" Then sql &= " AND PLANID =@PLANID" & vbCrLf
                    parms.Clear()
                    parms.Add("YEARS", sm.UserInfo.Years)
                    parms.Add("DISTID", sm.UserInfo.DistID)
                    If sPlanID <> "0" Then parms.Add("PLANID", sPlanID)
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)

                    If dr IsNot Nothing Then
                        If Convert.ToString(dr("CNT1")) <> "" Then
                            iCNT_4 = Val(dr("CNT1"))
                            iCNT_ALL += iCNT_4
                        End If
                    End If

                Case "5"  '1:訓練單位/2:報名者/3:在訓學員/4:結訓學員/5:曾受訓的所有學員
                    sql = "" & vbCrLf
                    sql &= " SELECT COUNT(CASE WHEN EMAIL LIKE '%@%' THEN 1 END) CNT1" & vbCrLf
                    sql &= " FROM V_STUDENTINFO" & vbCrLf
                    sql &= " WHERE 1=1" & vbCrLf
                    'sql &= " AND STUDSTATUS NOT IN (2,3)" & vbCrLf
                    'sql &= " AND FTDATE<=GETDATE()" & vbCrLf
                    sql &= " AND EMAIL IS NOT NULL" & vbCrLf
                    sql &= " AND YEARS =@YEARS" & vbCrLf
                    sql &= " AND DISTID =@DISTID" & vbCrLf
                    If sPlanID <> "0" Then sql &= " AND PLANID =@PLANID" & vbCrLf
                    parms.Clear()
                    parms.Add("YEARS", sm.UserInfo.Years)
                    parms.Add("DISTID", sm.UserInfo.DistID)
                    If sPlanID <> "0" Then parms.Add("PLANID", sPlanID)
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)

                    If dr IsNot Nothing Then
                        If Convert.ToString(dr("CNT1")) <> "" Then
                            iCNT_5 = Val(dr("CNT1"))
                            iCNT_ALL += iCNT_5
                        End If
                    End If

            End Select
        Next

        Dim lrMsg As String = ""
        lrMsg = "收件人數預估：" & CStr(iCNT_ALL) & "人" & vbCrLf
        If sm.UserInfo.DistID = "000" Then lrMsg &= "該使用者 所屬單位：署"
        sm.LastResultMessage = lrMsg
    End Sub

    '收件人數查詢
    Protected Sub BtnQuery1_Click(sender As Object, e As EventArgs) Handles BtnQuery1.Click
        SHOW_Query1()
    End Sub

    Protected Sub btnMailTest1_Click(sender As Object, e As EventArgs) Handles btnMailTest1.Click
        SendMailTest_ws()

        Common.MessageBox(Me, "ws郵件測試-郵件已寄送！")
        Exit Sub
    End Sub

    Sub SendMailTest_ws()
        Dim s_SESSSION_INFO As String = ""
        Try
            s_SESSSION_INFO = TIMS.GetErrorMsgSys()
        Catch ex As Exception
        End Try

        Dim s_USERAGENT_INFO As String = TIMS.GetUserAgent(Me, True)

        Dim from_emailaddress As String = TIMS.Utl_GetConfigSet("from_emailaddress") 'from_emailaddress'Cst_FromEmail 

        Dim sMailBody As String = $"##SendMailTest_ws{vbCrLf}{s_USERAGENT_INFO}{s_SESSSION_INFO}{vbCrLf}"
        '寫入新的錯誤訊息 'GetErrorMsgSys()
        sMailBody &= String.Format("Mail帳號：{0}", from_emailaddress) & vbCrLf
        sMailBody &= String.Format("MachineName：{0}", HttpContext.Current.Server.MachineName) & vbCrLf
        sMailBody &= String.Format("發出時間：{0}", Now) & vbCrLf

        Dim xSubject As String = "在職訓練資訊管理系統-ws郵件測試"

        'Call TIMS.SendMailTest(sMailBody)
        'txtTestEmail1.Text = TIMS.ClearSQM(txtTestEmail1.Text)
        'If txtTestEmail1.Text <> "" Then Call TIMS.SendMailTest(sMailBody, txtTestEmail1.Text)

        Threading.Thread.Sleep(1) '假設處理某段程序需花費n毫秒 (避免機器不同步)
        Try
            If txtTestEmail1.Text <> "" Then
                For Each v_m1 As String In txtTestEmail1.Text.Replace(",", ";").Split(";")
                    v_m1 = TIMS.ClearSQM(v_m1)
                    If v_m1 <> "" Then Call TIMS.SendMailTest(sMailBody, v_m1, xSubject)
                Next
            End If
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
        End Try

        Try
            Dim strToEmail As String = TIMS.Cst_EmailtoMe
            Call TIMS.SendMailTest(sMailBody, strToEmail, xSubject)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Throw ex
        End Try

    End Sub

    '系統自動發信 錯誤寄回
    Sub SendMailTest_sp()
        Dim s_SESSSION_INFO As String = ""
        Try
            s_SESSSION_INFO = TIMS.GetErrorMsgSys()
        Catch ex As Exception
        End Try

        Dim s_USERAGENT_INFO As String = TIMS.GetUserAgent(Me, True)
        'Dim sm As SessionModel = SessionModel.Instance()
        Const Cst_FromName As String = "系統自動發信"
        Dim from_emailaddress As String = TIMS.Utl_GetConfigSet("from_emailaddress") 'from_emailaddress'Cst_FromEmail 
        Dim xFrom As String = """" & Cst_FromName & """ <" & from_emailaddress & ">"

        '寫入新的錯誤訊息 'GetErrorMsgSys()
        Dim sMailBody As String = $"##SendMailTest_sp{vbCrLf}{s_USERAGENT_INFO}{s_SESSSION_INFO}{vbCrLf}"
        sMailBody &= String.Format("Mail帳號：{0}", from_emailaddress) & vbCrLf
        sMailBody &= String.Format("MachineName：{0}", HttpContext.Current.Server.MachineName) & vbCrLf
        sMailBody &= String.Format("發出時間：{0}", Now) & vbCrLf

        '置換換行符號
        Dim xMybody As String = TIMS.HtmlEncode1(sMailBody.Replace(TIMS.cst_js_chgRow, vbCrLf)).Replace(vbCrLf, TIMS.cst_html_br1)

        Dim xSubject As String = "在職訓練資訊管理系統-sp郵件測試"

        Dim xEncoding As System.Text.Encoding '= System.Text.Encoding.UTF8 'UniCode信件內容不以亂碼顯示
        xEncoding = System.Text.Encoding.UTF8 'UniCode信件內容不以亂碼顯示

        Threading.Thread.Sleep(1) '假設處理某段程序需花費n毫秒 (避免機器不同步)
        Try
            TIMS.LOG.Info(sMailBody)
            If txtTestEmail1.Text <> "" Then
                For Each v_m1 As String In txtTestEmail1.Text.Replace(",", ";").Split(";")
                    v_m1 = TIMS.ClearSQM(v_m1)
                    If v_m1 <> "" Then Call TIMS.SendMail(xFrom, v_m1, xSubject, xMybody, xEncoding)
                Next
            End If
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
        End Try

        Try
            Dim xEmailto As String = TIMS.Cst_EmailtoMe 'str_email2 'TIMS.Cst_EmailtoMe
            TIMS.SendMail(xFrom, xEmailto, xSubject, xMybody, xEncoding)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Throw ex
        End Try
    End Sub

    Protected Sub BtnMailTest2_Click(sender As Object, e As EventArgs) Handles btnMailTest2.Click
        SendMailTest_sp()

        Common.MessageBox(Me, "sp郵件測試-郵件已寄送！")
        Exit Sub
    End Sub

    Protected Sub btnMailInfo1_Click(sender As Object, e As EventArgs) Handles btnMailInfo1.Click
        Dim s_log1 As String = ""
        s_log1 &= String.Format("UserName: {0}", TIMS.Utl_GetConfigSet("MailuserName")) & vbCrLf
        s_log1 &= String.Format("from email: {0}", TIMS.Utl_GetConfigSet("from_emailaddress")) & vbCrLf
        s_log1 &= String.Format("SmtpServer: {0}", TIMS.Utl_GetConfigSet("MailServer")) & vbCrLf

        Common.MessageBox(Me, s_log1)
        Exit Sub
    End Sub
End Class
