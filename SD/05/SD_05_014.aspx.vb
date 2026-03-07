Partial Class SD_05_014
    Inherits AuthBasePage

    'STUD_TURNOUT2 (產投出缺勤使用)
    Const cst_ttip04 As String = "此班級已結訓，且送出補助申請，不可再修改出缺勤作業。"
    Dim oTest_flag As Boolean = False '測試
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        'If TIMS.sUtl_ChkTest() Then oTest_flag = True '測試

        If Not IsPostBack Then
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            DataGridTable.Style("display") = "none"
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            If Not Session("_SearchStr") Is Nothing Then
                Dim strSearch1 As String = Session("_SearchStr")
                center.Text = TIMS.GetMyValue(strSearch1, "center")
                RIDValue.Value = TIMS.GetMyValue(strSearch1, "RIDValue")
                TMID1.Text = TIMS.GetMyValue(strSearch1, "TMID1")
                OCID1.Text = TIMS.GetMyValue(strSearch1, "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(strSearch1, "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(strSearch1, "OCIDValue1")

                PageControler1.PageIndex = 0
                'PageControler1.PageIndex = TIMS.GetMyValue(Session("SearchStr"), "PageIndex")
                Dim MyValue As String = TIMS.GetMyValue(strSearch1, "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If
                If TIMS.GetMyValue(strSearch1, "submit") = "1" Then
                    'Button1_Click(sender, e)
                    Call Search2()
                End If
                Session("_SearchStr") = Nothing
            End If
        End If

    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        If OCIDValue1.Value = "" Then
            Errmsg += "班別代碼有誤，請確認點選職類/班別" & vbCrLf
            'Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別")
            'Exit Function
        End If

        If start_date.Text <> "" Then
            If Not TIMS.IsDate1(start_date.Text) Then
                Errmsg += "未出席日期 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            'Errmsg += "未出席日期 起始日期 為必填" & vbCrLf
        End If

        If end_date.Text <> "" Then
            If Not TIMS.IsDate1(end_date.Text) Then
                Errmsg += "未出席日期 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            'Errmsg += "未出席日期 迄止日期 為必填" & vbCrLf
        End If

        If Errmsg = "" Then
            If start_date.Text.ToString <> "" AndAlso end_date.Text.ToString <> "" Then
                If CDate(start_date.Text) > CDate(end_date.Text) Then
                    Errmsg += "【未出席日期】的起日不得大於【未出席日期】的迄日!!" & vbCrLf
                    'Common.MessageBox(Me, "【未出席日期】的起日不得大於【未出席日期】的迄日!!")
                    'Exit Function
                End If
            End If
        End If

        'If Convert.ToString(OCID.SelectedValue) = "" _
        '    OrElse Not IsNumeric(OCID.SelectedValue) Then
        '    Errmsg += "統計對象 為必選" & vbCrLf
        'End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢
    Sub Search2()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "輸入班級有誤!!")
            Exit Sub
        End If

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'UPDATE STUD_SUBSIDYCOST SET APPLIEDSTATUSM=null where SOCID='" & Hid_SOCID.Value & "' 
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.STOID" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,c.OCID" & vbCrLf
        sql &= " ,d.OrgName" & vbCrLf
        sql &= " ,c.ClassCName" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(b.StudentID) StudentID" & vbCrLf
        sql &= " ,e.Name" & vbCrLf
        sql &= " ,d.OrgID" & vbCrLf
        sql &= " ,a.LeaveDate" & vbCrLf
        'sql += " ,a.HOURS" & vbCrLf
        sql &= " ,dbo.DECODE(a.LEAVEID,'05',0,a.HOURS) HOURS" & vbCrLf '(修正舊資料)
        'sql += " ,CONVERT(varchar, a.Hours)+dbo.DECODE(a.LEAVEID,'05','(喪假)') Hours" & vbCrLf '喪假(LEAVEID:05)。
        sql &= " ,c.AppliedResultM " & vbCrLf '經費審核確認
        sql &= " ,CASE WHEN f.SOCID IS NOT NULL THEN 'Y' END SSC" & vbCrLf '查看是否有提出申請。
        sql &= " ,CASE WHEN f.APPLIEDSTATUSM IS NOT NULL THEN 'Y' END SSC2" & vbCrLf '提出申請且有結果。

        sql &= " ,a.LEAVEID" & vbCrLf
        sql &= " ,dbo.DECODE(a.LEAVEID,'05',a.HOURS,a.NIHOURS) NIHOURS" & vbCrLf '(修正舊資料)
        'LEAVEID='05'=喪假,其餘顯示資訊
        sql &= " ,dbo.DECODE(a.LEAVEID,'05','喪假',a.NIREASONS) NIREASONS" & vbCrLf '(修正舊資料)
        sql &= " ,c.IsClosed" & vbCrLf

        sql &= " FROM STUD_TURNOUT2 a " & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON a.SOCID=b.SOCID " & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON b.OCID=c.OCID " & vbCrLf
        sql &= " JOIN VIEW_RIDNAME d ON c.RID=d.RID " & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO e ON b.SID=e.SID " & vbCrLf
        sql &= " LEFT JOIN STUD_SUBSIDYCOST f ON f.SOCID =b.SOCID " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and b.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        If RIDValue.Value <> "" Then
            sql &= " and d.RID='" & RIDValue.Value & "'" & vbCrLf
        End If
        If start_date.Text.ToString <> "" Then
            sql &= " and a.LeaveDate >= " & TIMS.To_date(start_date.Text) & vbCrLf
        End If
        If end_date.Text.ToString <> "" Then
            sql &= " and a.LeaveDate <= " & TIMS.To_date(end_date.Text) & vbCrLf
        End If
        If cjobValue.Value <> "" Then
            'sql &= " and c.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " and c.CJOB_UNKEY=" & cjobValue.Value & vbCrLf
        End If

        Select Case sm.UserInfo.LID
            Case 0
                sql &= " and d.TPLANID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                sql &= " and d.YEARS='" & sm.UserInfo.Years & "'" & vbCrLf
            Case Else
                sql &= " and c.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End Select

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGridTable.Style("display") = "none"
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Style("display") = ""
            msg.Text = ""
            'PageControler1.SqlPrimaryKeyDataCreate(sql, "STOID", "OrgID,OCID,StudentID,LeaveDate")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "STOID"
            PageControler1.Sort = "OrgID,OCID,StudentID,LeaveDate"
            PageControler1.ControlerLoad()
        End If

        Dim flag_Lock_Edit1 As Boolean = False
        Button5.Enabled = True '新增
        If Not oTest_flag Then
            Dim flag_ISCLOSED As Boolean = If(Convert.ToString(drCC("ISCLOSED")).Equals("Y"), True, False)
            Dim flag_AppliedResultM As Boolean = If(Convert.ToString(drCC("AppliedResultM")).Equals("Y"), True, False)
            Dim flag_ChkSUBSIDYCOST As Boolean = If(TIMS.sUtl_ChkSUBSIDYCOST(OCIDValue1.Value, objconn), True, False)
            If flag_ISCLOSED OrElse flag_ChkSUBSIDYCOST OrElse flag_AppliedResultM Then flag_Lock_Edit1 = True
            If Not flag_ISCLOSED AndAlso Not flag_AppliedResultM Then flag_Lock_Edit1 = False
            If flag_Lock_Edit1 Then
                Button5.Enabled = False '新增
                TIMS.Tooltip(Button5, cst_ttip04, True)
                Common.MessageBox(Me, cst_ttip04)
                Exit Sub
            End If
        End If

    End Sub

    '查詢按鈕
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        Call Search2()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Dim s_OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim s_SOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        If s_OCID = "" Then Exit Sub
        If s_SOCID = "" Then Exit Sub
        Dim sSTOID As String = TIMS.GetMyValue(sCmdArg, "STOID")

        Select Case e.CommandName
            Case "edit"
                Call KeepSearchStr()
                'Response.Redirect("SD_05_014_add.aspx?ID=" & Request("ID") & "&" & e.CommandArgument & "")
                Dim url1 As String = ""
                url1 = "SD_05_014_add.aspx?ID=" & Request("ID") & "&" & e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "del"
                'Dim sCmdArg As String = e.CommandArgument
                'Dim sSTOID As String = TIMS.GetMyValue(sCmdArg, "STOID")
                If sSTOID <> "" Then
                    Dim d_parms As New Hashtable
                    d_parms.Add("STOID", sSTOID)
                    d_parms.Add("SOCID", s_SOCID)
                    Dim sql As String
                    sql = "DELETE STUD_TURNOUT2 WHERE STOID=@STOID And SOCID=@SOCID"
                    DbAccess.ExecuteScalar(sql, objconn, d_parms)
                End If
                'Button1_Click(Button1, e)
                Call Search2()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"

            Case ListItemType.Item, ListItemType.AlternatingItem
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                Dim Button3 As Button = e.Item.FindControl("Button3") '修改
                Dim Button4 As Button = e.Item.FindControl("Button4") '刪除

                Dim labHours As Label = e.Item.FindControl("labHours") 'labHours
                Dim sMsg2 As String = ""
                If Convert.ToString(drv("Hours")) <> "" Then
                    If Val(drv("Hours")) > 0 Then sMsg2 &= Val(drv("Hours"))
                End If
                If Convert.ToString(drv("NIHOURS")) <> "" AndAlso Val(drv("NIHOURS")) > 0 Then
                    sMsg2 &= String.Format("{0}({1}))", Val(drv("NIHOURS")), Convert.ToString(drv("NIREASONS")))
                End If
                labHours.Text = sMsg2

                '班級已做結訓作業
                Dim flag_ISCLOSED As Boolean = If(Convert.ToString(drv("ISCLOSED")).Equals("Y"), True, False)
                'cc.AppliedResultM 班級-學員經費審核結果已經通過
                Dim flag_AppliedResultM_Y As Boolean = If(Convert.ToString(drv("AppliedResultM")).Equals("Y"), True, False)
                '學員經費-有提出申請
                Dim flag_SSC_Y As Boolean = If(Convert.ToString(drv("SSC")).Equals("Y"), True, False)
                '學員經費-提出申請且有結果
                Dim flag_SSC2_Y As Boolean = If(Convert.ToString(drv("SSC2")).Equals("Y"), True, False)

                Dim flag_Lock_Edit1 As Boolean = False 'true:不可修改刪除
                If flag_ISCLOSED OrElse flag_AppliedResultM_Y OrElse flag_SSC_Y Then flag_Lock_Edit1 = True
                If Not flag_ISCLOSED AndAlso Not flag_AppliedResultM_Y AndAlso Not flag_SSC_Y AndAlso Not flag_SSC2_Y Then flag_Lock_Edit1 = False

                Button3.Enabled = True '修改
                Button4.Enabled = True  '刪除
                'true:不可修改刪除
                If flag_Lock_Edit1 Then
                    Button3.Enabled = False '修改
                    TIMS.Tooltip(Button3, "班級已做結訓作業,學員經費審核結果已經通過，不可修改") '修改
                    Button4.Enabled = False '刪除
                    TIMS.Tooltip(Button4, "班級已做結訓作業,學員經費審核結果已經通過，不可刪除") '刪除
                    'If oTest_flag Then Button4.Enabled = True 'test
                End If

                'If Convert.ToString(drv("IsClosed")) = "Y" Then
                '    If Button3.Enabled Then
                '        Button3.Enabled = False
                '        TIMS.Tooltip(Button3, "班級已做結訓作業，不可修改")
                '    End If
                '    If Button4.Enabled Then
                '        Button4.Enabled = False
                '        TIMS.Tooltip(Button4, "班級已做結訓作業，不可刪除")
                '    End If
                'End If

                Dim sCmdArg As String = ""
                sCmdArg &= "&STOID=" & Convert.ToString(drv("STOID"))
                sCmdArg &= "&SOCID=" & Convert.ToString(drv("SOCID"))
                sCmdArg &= "&OCID=" & Convert.ToString(drv("OCID"))
                '修改
                If Button3.Enabled Then Button3.CommandArgument = "act=edit" & sCmdArg
                '刪除
                If Button4.Enabled Then Button4.CommandArgument = "act=del" & sCmdArg 'drv("STOID").ToString
                Button4.Attributes("onclick") = TIMS.cst_confirm_delmsg1
        End Select
    End Sub

    Private Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call KeepSearchStr()
        'Response.Redirect("SD_05_014_add.aspx?ID=" & Request("ID") & "")
        Dim url1 As String = ""
        url1 = "SD_05_014_add.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Sub KeepSearchStr()
        Dim strSearch1 As String = ""
        strSearch1 = "center=" & center.Text & ""
        strSearch1 += "&RIDValue=" & RIDValue.Value & ""
        strSearch1 += "&TMID1=" & TMID1.Text & ""
        strSearch1 += "&OCID1=" & OCID1.Text & ""
        strSearch1 += "&TMIDValue1=" & TMIDValue1.Value & ""
        strSearch1 += "&OCIDValue1=" & OCIDValue1.Value & ""
        strSearch1 += "&PageIndex=" & DataGrid1.CurrentPageIndex + 1 & ""
        strSearch1 &= If(DataGridTable.Style("display") = "", "&submit=1", "&submit=0")

        Session("_SearchStr") = strSearch1
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Style("display") = "none"
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Style("display") = "none"
    End Sub

    Private Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
