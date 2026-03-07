Partial Class SD_08_015
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    Dim ttlPeopleMoney As Int64
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Memo.Visible = False
        msg.Text = ""

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            SearchTable.Visible = True

            ImgLSDate.Attributes("onclick") = "show_calendar('" & LSDate.ClientID & "','','','CY/MM/DD');" '離退日期(起)
            ImgLEDate.Attributes("onclick") = "show_calendar('" & LEDate.ClientID & "','','','CY/MM/DD');" '離退日期(迄)
        End If

        Button1.Attributes("onclick") = "javascript:return search();"
        Button3.Visible = False

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button7.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "Button8")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If LSDate.Text <> "" Then LSDate.Text = Trim(LSDate.Text)
        If LSDate.Text <> "" Then
            'LSDate.Text = LSDate.Text.Trim
            If Not TIMS.IsDate1(LSDate.Text) Then Errmsg += "離退日期區間 起始日期格式有誤" & vbCrLf
            If Errmsg = "" Then LSDate.Text = CDate(LSDate.Text).ToString("yyyy/MM/dd")
        Else
            Errmsg += "離退日期區間 起始日期 為必填" & vbCrLf
        End If

        If LEDate.Text <> "" Then LEDate.Text = Trim(LEDate.Text)
        If LEDate.Text <> "" Then
            'LEDate.Text = LEDate.Text.Trim
            If Not TIMS.IsDate1(LEDate.Text) Then Errmsg += "離退日期區間 迄止日期格式有誤" & vbCrLf
            If Errmsg = "" Then LEDate.Text = CDate(LEDate.Text).ToString("yyyy/MM/dd")
        Else
            Errmsg += "離退日期區間 迄止日期 為必填" & vbCrLf
        End If

        If Errmsg = "" Then
            If LSDate.Text.ToString <> "" AndAlso LEDate.Text.ToString <> "" Then
                If CDate(LSDate.Text) > CDate(LEDate.Text) Then
                    Errmsg += "【離退日期區間】的起日不得大於迄日!!" & vbCrLf
                End If
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        If OCIDValue.Value <> "" Then OCIDValue.Value = Trim(OCIDValue.Value)

        Dim dt As New DataTable
        'Dim dr As DataRow

        '津貼結果
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.IDNO" & vbCrLf
        sql += " ,a.Name" & vbCrLf
        sql += " ,a.Birthday" & vbCrLf
        sql += " ,a.TSDate" & vbCrLf
        sql += " ,a.TEDate" & vbCrLf
        sql += " ,a.OPayMoney" & vbCrLf
        sql += " ,a.IdentityID" & vbCrLf
        sql += " ,a.ClassName" & vbCrLf
        sql += " ,a.ApplyMonth" & vbCrLf
        sql += " ,k2.name  IdentityName" & vbCrLf
        sql += " ,k1.Reason+case when a.RTReason_O is not null and a.RTReason_O!=' ' then ','+a.RTReason_O  end  RTReason" & vbCrLf
        sql += " ,a.ApplyMoney" & vbCrLf
        sql += " ,b.orgname" & vbCrLf
        sql += " ,a.RtnMoney" & vbCrLf
        sql += " ,a.LDate" & vbCrLf
        sql += " ,a.OCID" & vbCrLf
        sql += " FROM Sub_SubSidyApply_All a" & vbCrLf
        sql += " LEFT JOIN org_orginfo b on a.orgid=b.orgid" & vbCrLf
        sql += " LEFT JOIN Key_RejectTReaSon k1 on k1.RTReasonId=a.RTReason" & vbCrLf
        sql += " LEFT JOIN Key_Identity k2 on k2.identityid=a.IdentityID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND a.fromtype='1' " & vbCrLf
        sql += " AND a.LFlag in('1','2')" & vbCrLf
        '離退日期區間
        If LSDate.Text <> "" Then LSDate.Text = Trim(LSDate.Text)
        If LSDate.Text <> "" Then
            sql += " AND a.LDate >= " & TIMS.To_date(LSDate.Text) & vbCrLf
        End If
        If LEDate.Text <> "" Then LEDate.Text = Trim(LEDate.Text)
        If LEDate.Text <> "" Then
            sql += " AND a.LDate <= " & TIMS.To_date(LEDate.Text) & vbCrLf
        End If
        '課程
        If OCIDValue.Value.ToString <> "" Then
            sql += " AND a.OCID in (" & OCIDValue.Value & ") "
        Else
            '單位
            If center.Text.ToString <> "" Then
                sql += " AND a.OCID in (SELECT OCID FROM Class_ClassInfo WHERE PlanId=" & sm.UserInfo.PlanID & " AND Years='" & Right(sm.UserInfo.Years.ToString, 2) & "'"
                If RIDValue.Value <> "" Then
                    sql += " AND RID LIKE '" & RIDValue.Value.ToString & "%'"
                Else
                    sql += " AND RID LIKE '" & sm.UserInfo.RID & "%'"
                End If
                sql += " )"
            End If
        End If
        sql += " ORDER BY A.OCID,a.IDNO, a.TSDate DESC"
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/* sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            DataGrid1.Visible = False
            msg.Text = "查無資料！"
            Button3.Visible = False
            Memo.Visible = False
            td1.InnerText = ""
            Td2.InnerText = ""
            Td3.InnerText = ""
            Exit Sub
        End Try


        DataGrid1.Visible = False
        msg.Text = "查無資料！"
        Button3.Visible = False
        Memo.Visible = False
        td1.InnerText = ""
        Td2.InnerText = ""
        Td3.InnerText = ""

        If dt.Rows.Count > 0 Then
            Dim objrow As DataRow

            If sm.UserInfo.PlanID <> 0 Then
                sql = "" & vbCrLf
                sql += " select b.name as UserName,d.name as UserRole" & vbCrLf
                sql += " ,c.Years+f.PlanName as UserPlan " & vbCrLf
                sql += " from Auth_AccRWPlan a " & vbCrLf
                sql += " join Auth_Account b on a.Account=b.Account " & vbCrLf
                sql += " join ID_Plan c on a.PlanID=c.PlanID " & vbCrLf
                sql += " join ID_Role d on b.RoleID=d.RoleID " & vbCrLf
                sql += " join ID_District e on c.DistID=e.DistID " & vbCrLf
                sql += " join Key_Plan f on c.TPlanID=f.TPlanID" & vbCrLf
            Else
                sql = "" & vbCrLf
                sql += " select b.name as UserName,c.name as UserRole " & vbCrLf
                sql += " from Auth_AccRWPlan a " & vbCrLf
                sql += " join Auth_Account b on a.Account=b.Account " & vbCrLf
                sql += " join ID_Role c on b.RoleID=c.RoleID" & vbCrLf
            End If

            sql += " where a.Account = '" & sm.UserInfo.UserID & "' and a.PlanID=" & sm.UserInfo.PlanID
            objrow = DbAccess.GetOneRow(sql, objconn)

            If sm.UserInfo.PlanID <> 0 Then
                td1.InnerText = objrow.Item("UserPlan")
            Else
                td1.InnerText = ""
            End If

            Td2.InnerText = center.Text
            Td3.InnerText = ""
            Td3.InnerText += CStr(CInt(CDate(LSDate.Text).ToString("yyyy")) - 1911) & "年度"
            Td3.InnerText += "繳回退訓學員訓練生活津貼補助清冊"

            msg.Text = ""
            Memo.Visible = True
            DataGrid1.Visible = True
            DataGrid1.DataSource = dt
            Button3.Visible = True
            DataGrid1.DataBind()

            PeopleNum.Text = dt.Rows.Count

            Dim url As String = ""
            url = "SD_08_015_1.aspx?ID=" & Request("ID")
            url += "&sad=" & LSDate.Text
            url += "&ead=" & LEDate.Text
            url += "&Rid=" & RIDValue.Value
            url += "&OCID=" & OCIDValue.Value
            url += "&OrgName=" & Server.UrlEncode(center.Text)
            url += "&PlanName=" & Server.UrlEncode(td1.InnerText)

            Button3.Attributes("onclick") = "javascript:openOrg('" + url + "'); return false;"
        End If
    End Sub

    '訓練機構快查
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim Flag As Boolean = False

        Dim dt As DataTable = TIMS.GetCookieTable(sm, objconn)

        For i As Integer = 1 To 5
            Dim s_find As String = String.Format("ItemName='SubsidyRID{0}' and ItemValue='{1}'", i, RIDValue.Value)
            If dt.Select(s_find).Length <> 0 Then
                Dim s_itemname As String = String.Format("SubsidyClass{0}", i)
                OCIDValue.Value = TIMS.GetCookieItemValue(dt, s_itemname)
                Button1_Click(sender, e)
                Flag = True
                Exit For
            End If
        Next
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim birth As Label = e.Item.FindControl("birth")
                Dim sd As Label = e.Item.FindControl("sd")
                Dim ed As Label = e.Item.FindControl("ed")
                Dim RtnData As Label = e.Item.FindControl("RtnData")
                Dim FinPayMoney As Label = e.Item.FindControl("FinPayMoney")

                e.Item.Cells(0).Text = e.Item.ItemIndex + 1

                birth.Text = Common.FormatDate2Roc(drv("Birthday")) '出生日期

                '受訓起迄日
                sd.Text = Common.FormatDate2Roc(drv("TSDate"))
                ed.Text = Common.FormatDate2Roc(drv("TEDate"))

                Dim iRtnMoney As Integer = 0
                If Convert.ToString(drv("RtnMoney")) <> "" Then
                    iRtnMoney = Convert.ToInt64(drv("RtnMoney"))
                End If

                RtnData.Text = Common.FormatDate2Roc(drv("LDate")) & "<br>" & drv("RTReason") '退訓日期、原因
                'FinPayMoney.Text = Convert.ToInt64(drv("OPayMoney")) - Convert.ToInt64(drv("RtnMoney")) '審核通過後實際可領取金額=原核發金額-本次退回金額
                'ttlPeopleMoney += Convert.ToInt64(drv("RtnMoney"))
                FinPayMoney.Text = Convert.ToInt64(drv("OPayMoney")) - iRtnMoney '審核通過後實際可領取金額=原核發金額-本次退回金額
                ttlPeopleMoney += iRtnMoney
        End Select

        PeopleMoney.Text = "&nbsp;&nbsp;" & FormatNumber(Convert.ToInt64(ttlPeopleMoney), 0).ToString & "&nbsp;&nbsp;"
    End Sub
End Class
