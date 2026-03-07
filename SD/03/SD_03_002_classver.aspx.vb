Partial Class SD_03_002_classver
    Inherits AuthBasePage

    Dim rq_OCID_value As String = ""
    Const cst_序號 As Integer = 0
    Const cst_身分證號碼 As Integer = 2
    'Const cst_性別=3
    'Const cst_報名日期=8
    Const cst_supplyID As Integer = 5 '//預算別(補助比例)
    Const cst_AppliedResult As Integer = 8 '//審核
    Const cst_AppliedResultR As Integer = 9 '//還原審核
    'Const cst_supplyID=13 '//預算別(補助比例)
    'Const cst_AppliedResult=17 '//審核
    'Const cst_AppliedResultR=18 '//還原審核
    Const cst_AR2R_R_還原審核 As String = "R" 'R:還原審核'AppliedResult2_R

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        '檢查Session是否存在 End

        '分頁設定 Start
        'PageControler1.PageDataGrid=DataGrid1
        '分頁設定 End
        If Not IsPostBack Then
            rq_OCID_value = TIMS.ClearSQM(Request("OCID")) '接收OCID
            OCIDValue1.Value = rq_OCID_value
            If OCIDValue1.Value <> "" Then Call DISPLAY_CLASSVER(rq_OCID_value) '執行DISPLAY_CLASSVER 程式
        End If
    End Sub

    Function SUtl_GetpointYN(OCID As String) As String
        Dim rst As String = "N"
        '是否為學分班
        Dim sql As String = ""
        sql &= " SELECT ISNULL(ppi.PointYN,'N') PointYN"
        sql &= " FROM Class_ClassInfo cci"
        sql &= " JOIN Plan_PlanInfo ppi ON cci.PlanID=ppi.PlanID AND cci.SeqNo=ppi.SeqNo AND cci.ComIDNO=ppi.ComIDNO"
        sql &= $" WHERE cci.OCID={OCID}"
        Try
            rst = $"{DbAccess.ExecuteScalar(sql, objconn)}"
        Catch ex As Exception
            rst = "N"
        End Try
        Return rst
    End Function

    Function SUtl_GetStudentsdt(OCID As String) As DataTable
        OCID = TIMS.ClearSQM(OCID)
        Dim parms As New Hashtable From {{"ocid", OCID}}
        Dim sql As String = ""
        sql &= " SELECT cs.SOCID AS SOCID," & vbCrLf
        sql &= " ss.Name AS CName," & vbCrLf
        sql &= " UPPER(ss.IDNO) AS IDNO," & vbCrLf
        sql &= " CASE WHEN ss.sex='M' THEN '男' WHEN ss.sex='F' THEN '女' END AS Sex," & vbCrLf
        'sql += "      DATEDIFF(YEAR ,ss.birthday ,GETDATE()) AS YearsOld," & vbCrLf
        'sql += "      dbo.TRUNC(dbo.MONTHS_BETWEEN(GETDATE() ,ss.birthday)/12) AS YearsOld," & vbCrLf
        sql &= " FLOOR(dbo.MONTHS_BETWEEN(GETDATE(),ss.birthday)/12) AS YearsOld," & vbCrLf
        sql &= " ki.Name AS IdentityName," & vbCrLf
        sql &= " oo.orgName AS OrgName," & vbCrLf
        sql &= " cc.ClassCName AS OCIDName," & vbCrLf
        sql &= " cc.OCID AS OCID," & vbCrLf
        sql &= " st.RelEnterDate AS RelEnterDate," & vbCrLf
        sql &= " st.eSerNum AS eSerNum," & vbCrLf
        sql &= " se.Actname AS Actname," & vbCrLf
        sql &= " se.ActNo AS ActNo," & vbCrLf
        '**by Milor 20080611----start
        '訓練費用跟補助比例, 都直接以Plan_PlanInfo.DefStdCost + Plan_PlanInfo.DefGovCost，除以開班人數
        sql &= " ROUND((ISNULL(c.DefStdCost,0)+ISNULL(c.DefGovCost,0))/cc.TNum,1) AS TotalCost," & vbCrLf
        sql &= " CASE WHEN cc.TNum > 0 THEN" & vbCrLf
        sql &= "  CASE WHEN CONVERT(VARCHAR ,cs.supplyid)='1' THEN ROUND(((ISNULL(c.DefStdCost,0)+ISNULL(c.DefGovCost,0))/cc.TNum * 0.8),1)" & vbCrLf
        sql &= "  WHEN CONVERT(VARCHAR ,cs.supplyid)='2' THEN ROUND(((ISNULL(c.DefStdCost,0)+ISNULL(c.DefGovCost,0))/cc.TNum * 1) ,1)" & vbCrLf
        sql &= "  WHEN CONVERT(VARCHAR ,cs.supplyid)='9' THEN 0 END END supplyIdCost," & vbCrLf
        '**by Milor 20080611----end
        '**by Milor 20080508--96年以前不管是否學分班，都是取自Plan_CostItem.Oprice X Plan_CostItem.ItemCost的總和----start
        '                     97年開始，非學分班取自Plan_CostItem.Oprice X Plan_CostItem.Itemage的總和
        '                     學分班則直接取用Plan_PlanInfo.TotalCost
        sql &= " dbo.fn_GET_GOVCOST(ss.IDNO,CONVERT(varchar, cc.STDate, 111)) GovCost," & vbCrLf
        sql &= " dbo.fn_GET_BIEPTBL(ss.IDNO,cc.STDate) BIEPTBL," & vbCrLf
        sql &= " DATEADD(month, -6, cc.STDate) AS BFDate," & vbCrLf
        sql &= " cc.STDate AS STDate," & vbCrLf
        sql &= " cs.AppliedResult AS AppliedResult," & vbCrLf
        sql &= " cs.budgetid as budgetid," & vbCrLf
        sql &= " kb.budName as budName," & vbCrLf
        sql &= " CASE WHEN CONVERT(VARCHAR ,cs.supplyid)='1' THEN '80%'" & vbCrLf
        sql &= "      WHEN CONVERT(VARCHAR ,cs.supplyID)='2' THEN '100%'" & vbCrLf
        sql &= "      WHEN CONVERT(VARCHAR ,cs.supplyID)='9' THEN '不補助' END AS supplyID," & vbCrLf
        sql &= " cs.MEMO AS MEMO" & vbCrLf
        sql &= " FROM class_classinfo cc" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass cs ON cc.OCID=cs.OCID" & vbCrLf
        sql &= " JOIN stud_studentinfo ss ON ss.SID=cs.SID" & vbCrLf
        sql &= " LEFT JOIN Stud_ServicePlace se ON se.SOCID=cs.SOCID" & vbCrLf
        'sql += " LEFT JOIN Stud_SubsidyCost scs ON scs.SOCID=cs.SOCID" & vbCrLf
        'sql += " LEFT JOIN stud_entertype st ON cs.SETID=st.SETID AND cc.ocid=st.ocid1 AND st.ocid1='" & OCID & "'" & vbCrLf
        '因為報名資料 stud_entertype 重複，增加檢測、排除的語法
        sql &= " LEFT JOIN (SELECT se.* FROM stud_entertype se" & vbCrLf
        sql &= " WHERE EXISTS (SELECT 'x' FROM (" & vbCrLf
        sql &= "    SELECT setid ,MAX(sernum) sernum ,MAX(EnterDate) EnterDate" & vbCrLf
        sql &= "    FROM stud_entertype WHERE ocid1='" & OCID & "' AND eSerNum IS NOT NULL "
        sql &= "    GROUP BY setid) g" & vbCrLf
        sql &= "   WHERE g.setid=se.setid AND g.EnterDate=se.EnterDate AND g.sernum=se.sernum)" & vbCrLf
        sql &= "   ) st ON cs.SETID=st.SETID AND cc.ocid=st.ocid1 AND cc.ocid=st.ocid1" & vbCrLf
        sql &= " LEFT JOIN org_orgInfo oo ON oo.comidno=cc.comidno" & vbCrLf
        sql &= " LEFT JOIN key_budget kb ON kb.budid=cs.budgetid" & vbCrLf
        sql &= " LEFT JOIN Key_Identity ki ON ki.identityid=cs.MIdentityID" & vbCrLf
        sql &= " LEFT JOIN Plan_PlanInfo c ON cc.PlanID=c.PlanID AND cc.ComIDNO=c.ComIDNO AND cc.SeqNo=c.SeqNo" & vbCrLf
        sql &= " WHERE cs.StudStatus NOT IN (2,3) AND cc.ocid=@ocid" & vbCrLf
        If InStr(Me.ViewState("sort"), "IDNO") > 0 Then
            sql &= " ORDER BY ss." & ViewState("sort").ToString & vbCrLf
        ElseIf InStr(Me.ViewState("sort"), "SEX") > 0 Then
            sql &= " ORDER BY ss." & ViewState("sort").ToString & vbCrLf
        Else
            sql &= " ORDER BY ss.IDNO" & vbCrLf
        End If

        Return DbAccess.GetDataTable(sql, objconn, parms)
    End Function

    Sub DISPLAY_CLASSVER(ByVal OCID As String)
        Button11.Visible = False '審核 儲存
        Button2.Visible = False '還原審核 儲存

        If Convert.ToString(Request("act")) = "R" Then '還原審核
            Button2.Visible = True
            'Me.DataGrid1.Columns(cst_AppliedResult).Visible=False
            DataGrid1.Columns(cst_AppliedResult).Visible = True
            DataGrid1.Columns(cst_AppliedResultR).Visible = True
        Else '學員資料審核
            Button11.Visible = True
            DataGrid1.Columns(cst_AppliedResult).Visible = True
            DataGrid1.Columns(cst_AppliedResultR).Visible = False
        End If

        Dim pointYN As String = SUtl_GetpointYN(OCID) '"N"
        Dim dt As DataTable = SUtl_GetStudentsdt(OCID)
        Button11.Enabled = False '學員資料 審核
        Button2.Enabled = False  '學員資料 還原審核

        'Button4.Enabled=False '列印
        DataGrid1.Visible = False
        msg.Text = "查無資料"
        If TIMS.dtNODATA(dt) Then Return

        If Convert.ToString(Request("act")) = "R" Then
            Button2.Enabled = True '學員資料 還原審核
        Else
            Button11.Enabled = True '學員資料 審核
        End If

        'Button4.Enabled=True
        DataGrid1.Visible = True
        msg.Text = ""

        If ViewState("sort") = "" Then ViewState("sort") = "IDNO"
        OrgName1.Text = dt.Rows(0).Item(6).ToString '報名機構名稱
        OCIDName1.Text = dt.Rows(0).Item(7).ToString '報名班級名稱
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "view"
                'KeepSearch()
                Dim rqMID As String = TIMS.Get_MRqID(Me)
                Dim s_act As String = TIMS.ClearSQM(Request("act"))
                Dim s_redirect_view As String = String.Concat("../01/SD_01_004_add.aspx?ID=", rqMID, "&", e.CommandArgument, If(s_act <> "", String.Concat("&act=", s_act), ""))
                Call TIMS.Utl_Redirect(Me, objconn, s_redirect_view)
            Case "link" 'Else '姓名(學員)
                Dim rqMID As String = TIMS.Get_MRqID(Me)
                Dim s_act As String = TIMS.ClearSQM(Request("act"))
                Dim s_redirect_link As String = String.Concat("../03/SD_03_002_add.aspx?ID=", rqMID & "&todo=1&", e.CommandArgument, If(s_act <> "", String.Concat("&act=", s_act), ""))
                Call TIMS.Utl_Redirect(Me, objconn, s_redirect_link)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound

        Select Case e.Item.ItemType
            Case ListItemType.Header
                'SelectAll.Attributes("onchange")="ChangeAll();"
                Dim HBudID_all As DropDownList = e.Item.FindControl("HBudID_all")
                HBudID_all = TIMS.Get_Budget(HBudID_all, 2)
                'HBudID_all.Attributes("onchange")="ChangeAll(" & cst_supplyID & ", this.selectedIndex );"
                HBudID_all.Attributes("onchange") = "ChangeAll('HBudID','" & HBudID_all.ClientID & "');"

                Dim SelectAll As DropDownList = e.Item.FindControl("SelectAll")
                'SelectAll.Attributes("onchange")="ChangeAll(" & cst_AppliedResult & ", this.selectedIndex );"
                SelectAll.Attributes("onchange") = "ChangeAll('AppliedResult2','" & SelectAll.ClientID & "');"

                Dim SelectAllR As DropDownList = e.Item.FindControl("SelectAllR")
                'SelectAllR.Attributes("onchange")="ChangeAll(" & CStr(cst_AppliedResultR - 1) & ",this.selectedIndex);"
                'SelectAllR.Attributes("onchange")="ChangeAll(" & cst_AppliedResultR & ",this.selectedIndex);"
                SelectAllR.Attributes("onchange") = "ChangeAll('AppliedResult2_R','" & SelectAllR.ClientID & "');"

                'R
                HBudID_all.Visible = Button11.Visible
                SelectAll.Visible = Button11.Visible

                e.Item.CssClass = "head_navy"
                If ViewState("sort").ToString.ToUpper <> "" Then
                    Dim mylabel As String
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Select Case ViewState("sort")
                        Case "IDNO", "IDNO DESC"
                            mylabel = "IDNO"
                            i = cst_身分證號碼
                            mysort.ImageUrl = If(Me.ViewState("sort") = "IDNO", "../../images/SortUp.gif", "../../images/SortDown.gif")
                    End Select
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(mysort)
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim LinkButton1 As LinkButton = e.Item.FindControl("LinkButton1")
                Dim signUpMemo As TextBox = e.Item.FindControl("signUpMemo")
                Dim IDLab As Label = e.Item.FindControl("IDLab")
                Dim star As Label = e.Item.FindControl("star1")
                Dim btn_view As Button = e.Item.FindControl("Btn_VIEW")
                ''Dim mybtn As LinkButton
                'M:請選擇/Y:審核通過/N:不補助/R:退件修正
                Dim AppliedResult2 As DropDownList = e.Item.FindControl("AppliedResult2")
                AppliedResult2.Enabled = Button11.Visible

                Dim Hid_SOCID1 As HiddenField = e.Item.FindControl("Hid_SOCID1")
                Dim Hid_SOCID2 As HiddenField = e.Item.FindControl("Hid_SOCID2")
                Dim KeyValue As HtmlInputHidden = e.Item.FindControl("KeyValue")
                Dim KeyValueR As HtmlInputHidden = e.Item.FindControl("KeyValueR")
                Dim HBudID As DropDownList = e.Item.FindControl("HBudID")
                HBudID = TIMS.Get_Budget(HBudID, 2)

                'Dim H_HBudid As TextBox=e.Item.FindControl("H_HBudid")
                Dim star2 As Label = e.Item.FindControl("star2")
                LinkButton1.Text = drv("CName").ToString
                LinkButton1.CommandArgument = "OCID=" & drv("OCID") & "&SD_03_002_classver=VIEW" & "&SOCID=" & drv("SOCID").ToString
                Hid_SOCID1.Value = Convert.ToString(drv("SOCID"))
                Hid_SOCID2.Value = Convert.ToString(drv("SOCID"))
                IDLab.Text = TIMS.Get_DGSeqNo(sender, e) '序號 

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    btn_view.CommandArgument = "eSerNum=" & drv("eSerNum").ToString _
                                            & "&IDNO=" & TIMS.ChangeIDNO(drv("IDNO").ToString) _
                                            & "&BFDate=" & Common.FormatDate(drv("BFDate").ToString) _
                                            & "&STDate=" & Common.FormatDate(drv("STDate").ToString) _
                                            & "&OCID=" & rq_OCID_value _
                                            & "&SD_03_002_classver=VIEW"

                    star.Visible = If(Val(drv("BIEPTBL")) > 0, True, False)
                    btn_view.Visible = If(star.Visible, True, False)
                End If

                If drv("budgetid").ToString <> "" Then '學員報名資料 預算別狀態
                    Common.SetListItem(HBudID, drv("budgetid").ToString)
                    If Convert.ToString(drv("budgetid")) = "99" Then HBudID.Enabled = False
                End If

                signUpMemo.Text = drv("MEMO").ToString '備註

                'M:請選擇/Y:審核通過/N:不補助/R:退件修正
                AppliedResult2.Enabled = True
                If Convert.ToString(drv("budgetid")) = "99" AndAlso IsDBNull(drv("AppliedResult")) Then
                    '預設值 (不補助)
                    Common.SetListItem(AppliedResult2, "N")
                    AppliedResult2.Enabled = False
                Else
                    '學員資料複審狀態
                    Dim Val_AR2 As String = ""
                    Val_AR2 = If(IsDBNull(drv("AppliedResult")), "M", Convert.ToString(drv("AppliedResult")))
                    Common.SetListItem(AppliedResult2, Val_AR2)
                End If
                ''HBudID.SelectedValue=Trim(drv("budid").ToString)
                'H_HBudid.Text=Trim(drv("budid").ToString)
                'If drv("budid").ToString <> "" Then '複審資料 預算別狀態
                '    Common.SetListItem(HBudID, drv("budid").ToString)
                'Else
                '    If drv("budgetid").ToString <> "" Then Common.SetListItem(HBudID, drv("budgetid").ToString) '學員報名資料 預算別狀態
                'End If
                'Button11.Visible 
                '學員資料審核
                HBudID.Enabled = True
                signUpMemo.Enabled = True
                If Convert.ToString(Request("act")) = "R" Then
                    '還原審核
                    HBudID.Enabled = False
                    signUpMemo.Enabled = False
                End If

                '20080624  andy  Start--- 
                star2.Visible = False
                Dim v_AppliedResult2 As String = TIMS.GetListValue(AppliedResult2)
                If v_AppliedResult2 = "M" Then star2.Visible = True
                ' andy  end-- 
                KeyValue.Value = "SOCID='" & drv("SOCID").ToString & "'"
                KeyValueR.Value = "SOCID='" & drv("SOCID").ToString & "'"
                'e.Item.Cells(cst_序號).ToolTip="SOCID '" & drv("SOCID").ToString & "'"
                TIMS.Tooltip(e.Item.Cells(cst_序號), "SOCID '" & drv("SOCID").ToString & "'")
                'Session("SearchSOCID")=drv("SOCID").ToString
        End Select
    End Sub

    '儲存-學員資料審核-產投
    '97年產業人才投資方案改為下列3點：
    '       1.分署(中心)不再做學員資格複審，由訓練單位作完即可【原來是由分署(中心)做複審】。
    '       2.訓練單位做學員資料維護的同時，也進行學員資料審核。
    '       3.分署(中心)可查詢學員資料，並有學員資料審核還原的權限。
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        '接收OCID
        rq_OCID_value = TIMS.ClearSQM(Request("OCID")) '接收OCID
        If OCIDValue1.Value = "" OrElse OCIDValue1.Value <> rq_OCID_value Then
            Common.MessageBox(Me, "查無審核資料，請重新查詢審核資料!!")
            Exit Sub
        End If

        Dim flagStar2a As Boolean = False
        Dim flagStar2b As Boolean = False
        Dim i As Integer = 0 '是否有審核資料
        'flagStar2a = False,'flagStar2b = False,'i = 0 '是否有審核資料

        For Each Item As DataGridItem In DataGrid1.Items
            Dim star2 As Label = Item.FindControl("star2")
            'If Not star2 Is Nothing Then star2.Visible=False
            Dim AppliedResult2 As DropDownList = Item.FindControl("AppliedResult2")
            Dim HBudID As DropDownList = Item.FindControl("HBudID")

            star2.Visible = False
            Select Case AppliedResult2.SelectedValue '審核選擇
                Case "M" '請選擇
                    star2.Visible = True
                    flagStar2a = True
                    'Common.MessageBox(Me, "請確認審核答案，(請選擇)無法存取!!") 'Exit Sub 'Case Else
                    'M:請選擇 Y:審核通過 N:不補助 R:退件修正
            End Select
            Select Case AppliedResult2.SelectedIndex '若審核的下接式選單是選
                Case 0 '請選擇
                    star2.Visible = True
                    flagStar2a = True
                    'Common.MessageBox(Me, "請確認審核答案，(請選擇)無法存取!!") 'Exit Sub
            End Select
            If HBudID.SelectedValue = "" Then
                star2.Visible = True
                flagStar2b = True
                'Common.MessageBox(Me, "請確認預算別，(請選擇)無法存取!!") 'Exit Sub
            End If
            'Select Case HBudID.SelectedValue 'Case "99" '不補助 'End Select
            i += 1
        Next

        If flagStar2a = True Then
            Common.MessageBox(Me, "請確認審核答案，(請選擇)無法存取!!")
            Exit Sub
        End If
        If flagStar2b = True Then
            Common.MessageBox(Me, "請確認預算別，(請選擇)無法存取!!")
            Exit Sub
        End If
        If i = 0 Then
            Common.MessageBox(Me, "查無審核資料，請重新查詢審核資料!!")
            Exit Sub
        End If
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "查無審核資料，請重新查詢審核資料!!")
            Exit Sub
        End If

        'pass = "Y" '判斷是否是不補助及成功,預設=Y 是指全班皆為不補助及成功
        Dim pass As String = "Y"
        Dim sqlS1 As String = $" SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID={OCIDValue1.Value}"
        Dim dt As DataTable = DbAccess.GetDataTable(sqlS1, objconn)
        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "查無資料，儲存有誤!!")
            Exit Sub
        End If

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim oTrans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                Dim sqlU1 As String = ""
                sqlU1 = " UPDATE CLASS_STUDENTSOFCLASS" & vbCrLf
                sqlU1 &= " SET MEMO=@MEMO ,APPLIEDRESULT=@APPLIEDRESULT ,ISAPPRPAPER=@ISAPPRPAPER" & vbCrLf
                sqlU1 &= " ,SUPPLYID=@SUPPLYID ,BUDGETID=@BUDGETID" & vbCrLf
                sqlU1 &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
                sqlU1 &= " WHERE SOCID=@SOCID" & vbCrLf
                Dim uCmd As New SqlCommand(sqlU1, TransConn, oTrans)

                For Each Item As DataGridItem In DataGrid1.Items
                    Dim KeyValue As HtmlInputHidden = Item.FindControl("KeyValue")
                    Dim AppliedResult2 As DropDownList = Item.FindControl("AppliedResult2")
                    Dim signUpMemo As TextBox = Item.FindControl("signUpMemo")
                    Dim HBudID As DropDownList = Item.FindControl("HBudID") 'DropDownList
                    'Dim H_HBudid As TextBox=Item.FindControl("H_HBudid") 'TextBox

                    If dt.Select(KeyValue.Value).Length > 0 Then
                        Dim dr As DataRow = dt.Select(KeyValue.Value)(0)

                        'M:請選擇/Y:審核通過/N:不補助/R:退件修正
                        Dim v_AppliedResult2 As String = TIMS.GetListValue(AppliedResult2)
                        Dim v_APPLIEDRESULT As String = "Y"
                        Dim v_HBudID As String = TIMS.GetListValue(HBudID)
                        If v_HBudID = "99" Then v_APPLIEDRESULT = "N"

                        With uCmd
                            .Parameters.Clear()
                            .Parameters.Add("MEMO", SqlDbType.VarChar).Value = signUpMemo.Text
                            'AppliedResult2 /M:請選擇/Y:審核通過/N:不補助/R:退件修正/
                            Select Case v_AppliedResult2'Convert.ToString(AppliedResult2.SelectedIndex) '若審核的下接式選單是選
                                Case "M" '"0" '請選擇(異常)
                                    pass = "N" '表班級裡有未審核狀況
                                    .Parameters.Add("SUPPLYID", SqlDbType.VarChar).Value = If(IsDBNull(dr("SUPPLYID")), Convert.DBNull, Convert.ToString(dr("SUPPLYID")))
                                    .Parameters.Add("APPLIEDRESULT", SqlDbType.VarChar).Value = If(IsDBNull(dr("APPLIEDRESULT")), Convert.DBNull, Convert.ToString(dr("APPLIEDRESULT")))
                                    .Parameters.Add("ISAPPRPAPER", SqlDbType.VarChar).Value = If(IsDBNull(dr("ISAPPRPAPER")), Convert.DBNull, Convert.ToString(dr("ISAPPRPAPER")))
                                Case "Y" '"1"  '審核通過
                                    .Parameters.Add("SUPPLYID", SqlDbType.VarChar).Value = If(IsDBNull(dr("SUPPLYID")), Convert.DBNull, Convert.ToString(dr("SUPPLYID")))
                                    .Parameters.Add("APPLIEDRESULT", SqlDbType.VarChar).Value = v_APPLIEDRESULT ' "Y" '學員資料複審狀態=Y
                                    .Parameters.Add("ISAPPRPAPER", SqlDbType.VarChar).Value = "Y"
                                Case "N" '"2"  '不補助
                                    .Parameters.Add("SUPPLYID", SqlDbType.VarChar).Value = "9" '0%
                                    .Parameters.Add("APPLIEDRESULT", SqlDbType.VarChar).Value = "N" '學員資料複審狀態=N
                                    .Parameters.Add("ISAPPRPAPER", SqlDbType.VarChar).Value = "Y"
                                Case "R" '"3"  '退件修正
                                    pass = "N" '表示班級裡有退件修正狀態
                                    .Parameters.Add("SUPPLYID", SqlDbType.VarChar).Value = If(IsDBNull(dr("SUPPLYID")), Convert.DBNull, Convert.ToString(dr("SUPPLYID")))
                                    .Parameters.Add("APPLIEDRESULT", SqlDbType.VarChar).Value = "R" '學員資料複審狀態=R
                                    .Parameters.Add("ISAPPRPAPER", SqlDbType.VarChar).Value = Convert.DBNull
                            End Select

                            '===學員資料審核==SD_03_002_classver===
                            'M:請選擇/Y:審核通過/N:不補助/R:退件修正
                            'Dim v_AppliedResult2 As String=TIMS.GetListValue(AppliedResult2)
                            Select Case v_AppliedResult2 'Convert.ToString(AppliedResult2.SelectedIndex)
                                Case "N" '"2"  '不補助
                                    .Parameters.Add("BUDGETID", SqlDbType.VarChar).Value = "99" '不補助
                                    '.Parameters.Add("BUDGETID", SqlDbType.VarChar).Value=If(IsDBNull(dr("budgetid")), Convert.DBNull, Convert.ToString(dr("budgetid")))
                                Case Else '不補助除外
                                    If v_HBudID <> "" Then
                                        .Parameters.Add("BUDGETID", SqlDbType.VarChar).Value = v_HBudID 'HBudID.SelectedValue '儲存目前狀態
                                    Else
                                        .Parameters.Add("BUDGETID", SqlDbType.VarChar).Value = If(IsDBNull(dr("BUDGETID")), Convert.DBNull, Convert.ToString(dr("BUDGETID")))
                                    End If
                            End Select
                            '===學員資料審核==SD_03_002_classver===
                            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            '.Parameters.Add("MODIFYDATE", SqlDbType.DateTime).Value=Now
                            .Parameters.Add("SOCID", SqlDbType.Int).Value = Convert.ToInt64(dr("SOCID"))
                            .ExecuteNonQuery()
                            'DbAccess.ExecuteNonQuery(uCmd.CommandText, trans, uCmd.Parameters)
                        End With
                    End If
                Next

                'DbAccess.UpdateDataTable(dt, da, trans)

                Dim u_sql As String = " UPDATE CLASS_CLASSINFO SET AppliedResultR=@AppliedResultR ,ModifyAcct=@ModifyAcct ,ModifyDate=GETDATE() WHERE OCID=@OCID"
                If (pass = "Y") Then '表示全班的審核資料只有通過跟不補助的狀態 
                    Dim u_parms As New Hashtable From {
                        {"ModifyAcct", sm.UserInfo.UserID},
                        {"OCID", rq_OCID_value},
                        {"AppliedResultR", "Y"}
                    }
                    DbAccess.ExecuteNonQuery(u_sql, oTrans, u_parms)
                Else ' 委訓單位有做過全班確認動作，但分署(中心)審核後退件
                    Dim u_parms As New Hashtable From {
                        {"ModifyAcct", sm.UserInfo.UserID},
                        {"OCID", rq_OCID_value},
                        {"AppliedResultR", "R"}
                    }
                    DbAccess.ExecuteNonQuery(u_sql, oTrans, u_parms)
                End If
                DbAccess.CommitTrans(oTrans)
                Common.MessageBox(Me, "儲存成功")

            Catch ex As Exception
                DbAccess.RollbackTrans(oTrans)
                Common.MessageBox(Me, "!!儲存失敗!!")

                'Common.MessageBox(Me, ex.ToString)
                Dim strErrmsg As String = ""
                strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
                Exit Sub
                'Common.MessageBox(Me, ex.ToString)
                'Throw ex
            End Try

        End Using

    End Sub

    '回上一頁功能
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Session("MySreach")=Me.ViewState("MySreach")
        'call TIMS.Utl_Redirect(Me, objconn,"SD_03_002_ver.aspx?ID=" & Request("ID"))
        Dim url As String = TIMS.GetFunIDUrl(Request("ID"), 0, objconn)
        Call TIMS.Utl_Redirect(Me, objconn, url & "?ID=" & Request("ID"))
    End Sub

    ''列印
    'Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_03_002_classver", "OCID=" & Request("OCID"))
    'End Sub

    '儲存-選擇還原審核
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        Dim v_OCID As String = TIMS.ClearSQM(Request("OCID"))
        If v_OCID = "" Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If
        Dim className As String = TIMS.GET_OCIDInfo(v_OCID, "OCID1", objconn)
        If className = "" Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        'Dim sSql As String = ""
        Dim FlagReturn As Boolean = False '選擇還原審核
        For Each item As DataGridItem In DataGrid1.Items
            Dim KeyValueR As HtmlInputHidden = item.FindControl("KeyValueR") 'SOCID "SOCID='" & drv("SOCID").ToString & "'"
            Dim AppliedResult2_R As DropDownList = item.FindControl("AppliedResult2_R")
            Dim v_AppliedResult2_R As String = TIMS.GetListValue(AppliedResult2_R) 'R:還原審核
            If KeyValueR.Value <> "" AndAlso v_AppliedResult2_R = cst_AR2R_R_還原審核 Then
                Dim flag_ok_1 As Boolean = False
                Dim sSql As String = ""
                sSql &= " SELECT COUNT(1) CNT1 FROM CLASS_STUDENTSOFCLASS"
                sSql &= $" WHERE OCID={v_OCID} AND {KeyValueR.Value}"
                Dim drCnt1 As DataRow = DbAccess.GetOneRow(sSql, objconn)
                If drCnt1 IsNot Nothing AndAlso Convert.ToString(drCnt1("CNT1")) = "1" Then
                    flag_ok_1 = True
                End If

                If flag_ok_1 Then
                    Dim pmsXU As New Hashtable From {{"ModifyAcct", sm.UserInfo.UserID}}
                    Dim sSqlXU As String = ""
                    sSqlXU &= " UPDATE CLASS_STUDENTSOFCLASS"
                    sSqlXU &= " SET APPLIEDRESULT=NULL,ModifyAcct=@ModifyAcct,ModifyDate=GETDATE()"
                    sSqlXU &= $" WHERE OCID={v_OCID} AND {KeyValueR.Value}"
                    DbAccess.ExecuteNonQuery(sSqlXU, objconn, pmsXU) '學員審核狀態=NULL
                End If

                If Not FlagReturn Then FlagReturn = True '選擇還原審核
            End If
        Next

        If Not FlagReturn Then
            Common.MessageBox(Me, "無還原審核資料，儲存結束。")
            Exit Sub
        End If

        Dim PXU2 As New Hashtable From {{"ModifyAcct", sm.UserInfo.UserID}}
        Dim sSql2 As String = ""
        sSql2 &= " UPDATE CLASS_CLASSINFO"
        sSql2 &= " SET APPLIEDRESULTR='C',ModifyAcct=@ModifyAcct,ModifyDate=GETDATE()"
        sSql2 &= $" WHERE OCID={v_OCID}"
        '班級審核狀態=C
        DbAccess.ExecuteNonQuery(sSql2, objconn, PXU2)
        Common.MessageBox(Me, "儲存成功")

    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        ViewState("sort") = If(ViewState("sort") <> e.SortExpression, e.SortExpression, $"{e.SortExpression} DESC")
        rq_OCID_value = TIMS.ClearSQM(Request("OCID")) '接收OCID
        DISPLAY_CLASSVER(rq_OCID_value)
    End Sub

End Class
