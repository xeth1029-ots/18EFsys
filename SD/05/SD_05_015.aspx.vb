Partial Class SD_05_015
    Inherits AuthBasePage

    '產業人才投資方案專用
    Const Cst_m1yes As String = "本班學員經費審核結果通過"
    Const Cst_m1no As String = "本班學員經費審核結果尚未通過"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

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

        Button1.Attributes("onclick") = "return CheckSearch();"

        If Not IsPostBack Then
            Button6.Visible = False
            msg.Text = ""
            DataGrid1.Visible = False
            Button3.Visible = False
            '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
            If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                Button6.Visible = True
                Call SD05015_GetSession()
                If Convert.ToString(Me.ViewState("Button1")) = "1" Then
                    Button2.Disabled = True
                    chooseclass.Disabled = True
                    Button1.Visible = False
                    TIMS.Tooltip(Button2, "單一查詢動作不提供此功能", True)
                    TIMS.Tooltip(chooseclass, "單一查詢動作不提供此功能", True)
                    TIMS.Tooltip(Button1, "單一查詢動作不提供此功能", True)

                    'Button1_Click(sender, e)
                    Call sSearch1()
                End If
                'rbCreditPoints.Style.Item("display") = "inline"
            Else
                'DataGridTable.Style("display") = "none"
                center.Text = sm.UserInfo.OrgName
                RIDValue.Value = sm.UserInfo.RID
                '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            End If

        End If

    End Sub

    Sub sSearch1()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If


        '排除離退訓學員輸入資料 by AMU 20090916
        Dim pms_s1 As New Hashtable From {{"OCID", OCIDValue1.Value}}
        Dim sql As String = ""
        sql &= " SELECT a.SOCID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(a.StudentID) StudentID" & vbCrLf
        sql &= " ,b.Name,a.CreditPoints,a.StudStatus" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a " & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b ON a.SID=b.SID" & vbCrLf
        sql &= " WHERE a.STUDSTATUS NOT IN (2,3)" & vbCrLf
        sql &= " AND a.OCID=@OCID" & vbCrLf
        sql &= " ORDER BY a.StudentID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms_s1)

        'DataGridTable.Style("display") = "none"
        msg.Text = "查無資料"
        DataGrid1.Visible = False
        Button3.Visible = False

        If dt.Rows.Count > 0 Then
            'sql = "SELECT * FROM Class_ClassInfo WHERE OCID='" & OCIDValue1.Value & "'"
            'dr = DbAccess.GetOneRow(sql, objconn)
            If drCC("AppliedResultM").ToString = "Y" Then
                Me.Button3.Enabled = False
                TIMS.Tooltip(Button3, Cst_m1yes)
            Else
                Me.Button3.Enabled = True
                TIMS.Tooltip(Button3, Cst_m1no)
            End If

            Me.ViewState("AppliedResultM") = drCC("AppliedResultM").ToString

            'DataGridTable.Style("display") = "inline"
            msg.Text = ""
            DataGrid1.Visible = True
            Button3.Visible = True

            DataGrid1.DataKeyField = "SOCID"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sSearch1()
    End Sub

    ''' <summary>是否取得結訓資格 </summary>
    ''' <param name="CreditPoints"></param>
    ''' <param name="oValue"></param>
    Sub SET_CreditPoints_SelectedIndex1(ByRef CreditPoints As DropDownList, ByRef oValue As Object)
        If oValue Is Nothing OrElse IsDBNull(oValue) OrElse Convert.ToString(oValue) = "" Then Return
        CreditPoints.SelectedIndex = If(Convert.ToInt32(oValue) = 1, 1, 2) '1:是／2:否
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim AppliedResultM As String = Me.ViewState("AppliedResultM")
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass = "SD_TD1"
                Dim SelectAll As DropDownList = e.Item.FindControl("SelectAll")
                SelectAll.Enabled = If(AppliedResultM = "Y", False, True)
                SelectAll.Attributes("onchange") = "ChangeAll(this.selectedIndex);"
            Case ListItemType.Item, ListItemType.AlternatingItem
                'If e.Item.ItemType = ListItemType.Item Then                    e.Item.CssClass = "SD_TD2"                End If
                Dim drv As DataRowView = e.Item.DataItem
                '是否取得結訓資格
                Dim CreditPoints As DropDownList = e.Item.FindControl("CreditPoints")
                SET_CreditPoints_SelectedIndex1(CreditPoints, drv("CreditPoints"))

                If AppliedResultM = "Y" Then
                    CreditPoints.Enabled = False
                    TIMS.Tooltip(CreditPoints, Cst_m1yes)
                Else
                    CreditPoints.Enabled = True
                    TIMS.Tooltip(CreditPoints, Cst_m1no)
                End If
        End Select
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Dim conn As SqlConnection
        ''2006/03/28 add conn by matt
        'conn = DbAccess.GetConnection
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        Try
            'Dim sql As String,'Dim dt As DataTable,'Dim dr As DataRow,'Dim da As SqlDataAdapter = Nothing,
            'sql = "SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCIDValue1.Value & "'" ' order by  StudentID"
            ''2006/03/28 add conn by matt
            'dt = DbAccess.GetDataTable(sql, da, objconn)
            Dim i As Integer = 0
            Dim OCID_V As String = TIMS.ClearSQM(OCIDValue1.Value)
            For Each item As DataGridItem In DataGrid1.Items
                Dim CreditPoints As DropDownList = item.FindControl("CreditPoints")
                Dim SOCID_V As String = TIMS.ClearSQM(DataGrid1.DataKeys(i))
                If OCID_V <> "" AndAlso SOCID_V <> "" Then
                    Dim pms_s1 As New Hashtable From {{"OCID", OCID_V}, {"SOCID", SOCID_V}}
                    Dim sSql As String = ""
                    sSql &= " SELECT 1 FROM CLASS_STUDENTSOFCLASS"
                    sSql &= " WHERE OCID=@OCID AND SOCID=@SOCID"
                    Dim dtS1 As DataTable = DbAccess.GetDataTable(sSql, objconn, pms_s1)
                    If dtS1.Rows.Count > 0 Then
                        '未選擇／請選擇
                        Dim oCreditPoints As Object = Convert.DBNull
                        Select Case CreditPoints.SelectedIndex
                            Case 1 '是 'dr("CreditPoints") = True '1
                                oCreditPoints = 1
                            Case 2 '否 'dr("CreditPoints") = False '0
                                oCreditPoints = 0
                        End Select
                        Dim pms_u1 As New Hashtable From {{"OCID", OCID_V}, {"SOCID", SOCID_V}}
                        pms_u1.Add("CreditPoints", oCreditPoints)
                        pms_u1.Add("ModifyAcct", sm.UserInfo.UserID)
                        Dim uSql As String = ""
                        uSql &= " UPDATE CLASS_STUDENTSOFCLASS"
                        uSql &= " SET CreditPoints=@CreditPoints,ModifyAcct=@ModifyAcct,ModifyDate=GETDATE()"
                        uSql &= " WHERE OCID=@OCID AND SOCID=@SOCID"
                        DbAccess.ExecuteNonQuery(uSql, objconn, pms_u1)
                    End If
                End If
                i += 1
            Next
            'DbAccess.UpdateDataTable(dt, da)

        Catch ex As Exception
            '取得錯誤資訊寫入
            Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf, TIMS.GetErrorMsg(Me), vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)

            strErrmsg = String.Concat("!!儲存失敗!!", vbCrLf, ex.Message, vbCrLf)
            Common.MessageBox(Me, strErrmsg)
            Exit Sub
        End Try

        Common.MessageBox(Me, "儲存成功")
        'Call Button1_Click(sender, e)
        Call sSearch1()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        'DataGridTable.Style("display") = "none"
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        'DataGridTable.Style("display") = "none"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Sub SD05015_GetSession()
        If Session("_SearchSD05015") IsNot Nothing Then
            center.Text = TIMS.GetMyValue(Session("_SearchSD05015"), "center")
            RIDValue.Value = TIMS.GetMyValue(Session("_SearchSD05015"), "RIDValue")
            TMID1.Text = TIMS.GetMyValue(Session("_SearchSD05015"), "TMID1")
            OCID1.Text = TIMS.GetMyValue(Session("_SearchSD05015"), "OCID1")
            TMIDValue1.Value = TIMS.GetMyValue(Session("_SearchSD05015"), "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(Session("_SearchSD05015"), "OCIDValue1")

            Me.ViewState("Button1") = TIMS.GetMyValue(Session("_SearchSD05015"), "Button1")
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TIMS.Utl_Redirect1(Me, "SD_05_005.aspx?ID=" & Request("ID"))
    End Sub
End Class
