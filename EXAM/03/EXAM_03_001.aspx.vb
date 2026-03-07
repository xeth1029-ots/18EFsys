Partial Class EXAM_03_001
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = dg_Sch

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            If Not Session("_SearchStr") Is Nothing Then
                Dim MyArray As Array
                Dim MyItem As String
                Dim MyValue As String
                MyArray = Split(Session("_SearchStr"), "&")
                For i As Integer = 0 To MyArray.Length - 1
                    MyItem = Split(MyArray(i), "=")(0)
                    MyValue = Split(MyArray(i), "=")(1)
                    Select Case MyItem
                        Case "PageIndex"
                            PageControler1.PageIndex = MyValue
                    End Select
                Next
                tab_view.Visible = False
                Session("_SearchStr") = Nothing
            End If

        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button5.Attributes("onclick") = "choose_class();"

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        'btn_dgcrt.Attributes("onclick") = "return dg_create();"
        'btn_save.Attributes("onclick") = "return check_asave();"
    End Sub

    Private Sub btn_sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_sch.Click
        Call search()
    End Sub

    Sub search()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch) '顯示列數不正確

        Dim sql As String = ""
        sql &= " select distinct  ec.ocid " & vbCrLf
        sql &= "  ,ec.isonline ,ec.avail" & vbCrLf
        'sql &= " ,substring(convert(varchar,ec.examsdate,108),1,5) examSTIME" & vbCrLf '★
        sql &= "  ,convert(varchar(5), ec.examsdate, 108) examSTIME" & vbCrLf
        'sql &= " ,substring(convert(varchar,ec.examedate,108),1,5) examETIME " & vbCrLf '★
        sql &= "  ,convert(varchar(5), ec.examedate, 108) examETIME " & vbCrLf
        sql &= "  ,cc.RID, ie.DistID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.classcname,cc.cycltype) classcname" & vbCrLf
        sql &= "  ,CONVERT(varchar, cc.examdate, 111) examDATE" & vbCrLf
        sql &= " from exam_classdata ec " & vbCrLf
        sql &= " join class_classinfo cc on cc.ocid=ec.ocid " & vbCrLf
        sql &= " join id_examtype ie on ie.etid=ec.etid " & vbCrLf
        sql &= " join auth_relship ar on ar.RID =cc.RID" & vbCrLf
        sql &= " where 1=1" & vbCrLf

        '不卡署(局)的權限
        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
        Else
            sql &= "    and cc.PlanID ='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        'sql &= "    and cc.PlanID ='" & sm.UserInfo.PlanID & "'" & vbCrLf

        If Me.RIDValue.Value <> "" Then
            sql &= " and cc.RID LIKE '" & Me.RIDValue.Value & "%'" & vbCrLf
        Else
            sql &= " and cc.RID LIKE '" & sm.UserInfo.RID & "%'" & vbCrLf
        End If

        '不卡署(局)的權限
        If sm.UserInfo.DistID <> "000" Then
            sql &= " and ie.DistID = '" & sm.UserInfo.DistID & "'" & vbCrLf
        End If

        If OCIDValue1.Value <> "" Then
            sql &= " and cc.ocid=" & OCIDValue1.Value & vbCrLf
        End If
        If txt_examdate.Text <> "" Then
            sql &= " and cc.examdate=convert(datetime, '" & txt_examdate.Text & "', 111)" & vbCrLf
        End If

        msg.Visible = True
        tab_view.Visible = False
        PageControler1.Visible = False

        If TIMS.Get_SQLRecordCount(sql, objconn) > 0 Then
            msg.Visible = False
            tab_view.Visible = True
            PageControler1.Visible = True

            PageControler1.SqlString = sql
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub dg_Sch_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            'Dim sql As String
            Dim drv As DataRowView = e.Item.DataItem
            Dim btnedit As Button = e.Item.FindControl("btn_edit")
            Dim btnopen As Button = e.Item.FindControl("btn_open")
            Dim btnclose As Button = e.Item.FindControl("btn_close")
            Dim btndel As Button = e.Item.FindControl("btn_del")
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1

            '甄試日期是否填寫
            If IsDBNull(drv("examDATE")) Then
                e.Item.Cells(4).Text = "未填寫"
            End If

            '是否為線上考試
            If drv("isonline") = "Y" Then
                e.Item.Cells(3).Text = "線上考試"
                If e.Item.Cells(5).Text = "1" Then
                    btnopen.Visible = False
                Else
                    btnclose.Visible = False
                End If
            Else
                e.Item.Cells(3).Text = "一般筆試"
                btnopen.Visible = False
                btnclose.Visible = False
            End If

            '結束時間為NULL
            If IsDBNull(drv("examETIME")) Then
                e.Item.Cells(6).Text = "－"
            Else
                e.Item.Cells(6).Text = drv("examDATE") & vbCrLf & drv("examSTIME") & "~" & drv("examETIME")
            End If

            btndel.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆資料?');"
            btnedit.CommandArgument = drv("ocid").ToString + "," + drv("isonline").ToString
            btnopen.CommandArgument = btnedit.CommandArgument 'drv("ocid").ToString + "," + drv("isonline").ToString
            btnclose.CommandArgument = btnedit.CommandArgument 'drv("ocid").ToString + "," + drv("isonline").ToString
            btndel.CommandArgument = btnedit.CommandArgument 'drv("ocid").ToString + "," + drv("isonline").ToString

        End If
    End Sub

    Private Sub dg_Sch_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        Dim sql As String = ""
        Dim arr() As String = Split(e.CommandArgument, ",")
        ViewState("OCID") = TIMS.ClearSQM(arr(0))
        ViewState("IsOnline") = TIMS.ClearSQM(arr(1))
        Dim MRqID As String = TIMS.ClearSQM(Request("ID"))

        Select Case e.CommandName
            Case "edit" '修改
                'Exam_Edit_Cmd(e)
                ViewState("Redirect") = $"exam_03_001_add.aspx?e=1&ID={MRqID}&un=edit&OCID={ViewState("OCID")}&IsOnline={ViewState("IsOnline")}"
                TIMS.Utl_Redirect1(Me, ViewState("Redirect"))

            Case "avail_open" '考試啟動
                sql = $"update exam_classdata set avail='1' where ocid={ViewState("OCID")} and isonline='{ViewState("IsOnline")}'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "考試已啟動!!")

                search()

            Case "avail_close" '考試關閉
                sql = $"update exam_classdata set avail='0' where ocid={ViewState("OCID")} and isonline='{ViewState("IsOnline")}'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "考試已關閉!!")

                search()

            Case "del" '刪除
                sql = "" & vbCrLf
                sql &= " delete exam_classquestion " & vbCrLf
                sql &= " where ecid in (" & vbCrLf
                sql &= " 	select ecid from exam_classdata " & vbCrLf
                sql &= $" 	where ocid={ViewState("OCID")} AND isonline='{ViewState("IsOnline")}'" & vbCrLf
                sql &= " )" & vbCrLf
                DbAccess.ExecuteNonQuery(sql, objconn)

                sql = $" delete exam_classdata where ocid={ViewState("OCID")} AND isonline='{ViewState("IsOnline")}'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "資料已刪除!!")

                search()
        End Select

    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        Dim MRqID As String = TIMS.ClearSQM(Request("ID"))

        ViewState("Redirect") = $"exam_03_001_add.aspx?e=1&ID={MRqID}&un=add"
        TIMS.Utl_Redirect1(Me, ViewState("Redirect"))

    End Sub
End Class
