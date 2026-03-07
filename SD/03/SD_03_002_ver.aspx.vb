Partial Class SD_03_002_ver
    Inherits AuthBasePage

    Dim iPlanKind As Integer = 0

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    End If
        'End If

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        '分頁設定 Start
        PageControler1.PageDataGrid = DG_ClassInfo
        PageControler2.PageDataGrid = DataGrid2
        '分頁設定 End
        'ProcessType=Request("ProcessType")

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "bt_search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '依sm.UserInfo.PlanID取得PlanKind
        iPlanKind = TIMS.Get_PlanKind(Me, objconn)
        'PlanKind=DbAccess.ExecuteScalar("SELECT PlanKind From ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'")
        If iPlanKind = 1 Then
            Button5.Attributes("onclick") = "choose_class(2);"
        Else
            Button5.Attributes("onclick") = "choose_class(1);"
        End If

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            If Not Session("_SearchStr") Is Nothing Then
                ViewState("Load") = TIMS.GetMyValue(Session("_SearchStr"), "Load")
                'GetSearchStr
                If ViewState("Load") = "SD_03_002_ver" Then
                    center.Text = TIMS.GetMyValue(Session("_SearchStr"), "center")
                    RIDValue.Value = TIMS.GetMyValue(Session("_SearchStr"), "RIDValue")
                    'TMID1.Text=TIMS.GetMyValue(Session("_SearchStr"), "TMID1")
                    'OCID1.Text=TIMS.GetMyValue(Session("_SearchStr"), "OCID1")
                    'TMIDValue1.Value=TIMS.GetMyValue(Session("_SearchStr"), "TMIDValue1")
                    'OCIDValue1.Value=TIMS.GetMyValue(Session("_SearchStr"), "OCIDValue1")
                    TMID1.Text = TIMS.GetMyValue(Session("_SearchStr"), "rTMID1")
                    OCID1.Text = TIMS.GetMyValue(Session("_SearchStr"), "rOCID1")
                    TMIDValue1.Value = TIMS.GetMyValue(Session("_SearchStr"), "rTMIDValue1")
                    OCIDValue1.Value = TIMS.GetMyValue(Session("_SearchStr"), "rOCIDValue1")
                    start_date.Text = TIMS.GetMyValue(Session("_SearchStr"), "start_date")
                    end_date.Text = TIMS.GetMyValue(Session("_SearchStr"), "end_date")
                    ViewState("NotOpen") = TIMS.GetMyValue(Session("_SearchStr"), "NotOpen")
                    If ViewState("NotOpen") <> "" Then Common.SetListItem(NotOpen, ViewState("NotOpen"))
                    TxtPageSize.Text = TIMS.GetMyValue(Session("_SearchStr"), "TxtPageSize")
                    If TIMS.GetMyValue(Session("_SearchStr"), "submit") = "1" Then
                        Session("_SearchStr") = Nothing
                        ShowGrid()
                    End If
                End If
            End If

            ''970505  Andy SD_03_002.aspx.vb 回上頁
            'If (Session("Review")="yes") Then Session("Review")="no"
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            PageControler1.Visible = False
            PageControler2.Visible = False
        End If
    End Sub

    '顯示查詢資料 SQL
    Sub ShowGrid()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DG_ClassInfo) '顯示列數不正確
        DataGrid2.PageSize = TxtPageSize.Text

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)

        Dim sqlstr As String = ""
        sqlstr = "
SELECT a.YEARS ,a.CYCLTYPE ,a.OCLASSID ,a.PLANID ,a.OCID ,a.COMIDNO ,a.SEQNO ,a.ORGNAME,a.APPLIEDRESULTR
,a.CLASSCNAME2,a.TRAINNAME ,a.TPROPERTYID ,a.HOURRANNAME ,a.STDATE ,a.FTDATE ,a.RID
FROM VIEW2 a
WHERE 0=0"

        If sm.UserInfo.RID = "A" Then
            sqlstr &= " AND a.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sqlstr &= " AND a.YEARS='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sqlstr &= " AND a.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If

        If RIDValue.Value <> "" Then
            sqlstr &= " AND a.RID LIKE '" & RIDValue.Value & "%'" & vbCrLf
        Else
            sqlstr &= " AND a.RID LIKE '" & sm.UserInfo.RID & "%'" & vbCrLf
        End If
        If OCIDValue1.Value <> "" Then sqlstr &= " AND a.OCID='" & OCIDValue1.Value & "'" & vbCrLf

        If start_date.Text <> "" Then sqlstr &= " AND a.STDate >= " & TIMS.To_date(start_date.Text) & vbCrLf
        If end_date.Text <> "" Then sqlstr &= " AND a.STDate <= " & TIMS.To_date(end_date.Text) & vbCrLf

        Panel.Visible = False '隱藏 未審核畫面
        DG_ClassInfo.Visible = False '隱藏 未審核畫面
        PageControler1.Visible = False

        DataGridTable2.Visible = False '隱藏 己審核畫面
        DataGrid2.Visible = False   '隱藏 己審核畫面
        PageControler2.Visible = False

        msg.Text = "查無資料!!(查無資料時，請確認「學員資料維護」功能 ，已做「學員資料審核」)"

        Select Case NotOpen.SelectedValue '選擇審核狀態
            Case "N" '未審核
                sqlstr &= " AND a.AppliedResultR='C'" & vbCrLf

            Case "Y" '已審核
                sqlstr &= " AND a.AppliedResultR='Y'" & vbCrLf

            Case Else '退件修正
                sqlstr &= " AND a.AppliedResultR='R'" & vbCrLf

        End Select

        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn)
        If dt.Rows.Count > 0 Then
            '寫入Log查詢 SubInsAccountLog1 (Auth_Accountlog)
            Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm查詢, 2, OCIDValue1.Value, "")
        End If
        If dt.Rows.Count = 0 Then Return

        Select Case NotOpen.SelectedValue '選擇審核狀態
            Case "N" '未審核
                Panel.Visible = True '顯示 未審核畫面
                DG_ClassInfo.Visible = True '顯示 未審核畫面
                PageControler1.Visible = True

            Case "Y" '已審核
                DataGridTable2.Visible = True '顯示 己審核畫面
                DataGrid2.Visible = True '顯示 己審核畫面
                PageControler2.Visible = True

                msg.Text = ""
                PageControler2.PageDataTable = dt 'sqlstr
                'PageControler2.PrimaryKey = "OCID"
                'PageControler2.Sort = "ClassID,CyclType"
                PageControler2.ControlerLoad()

                Return
            Case Else '退件修正
                Panel.Visible = True '顯示 未審核畫面
                DG_ClassInfo.Visible = True '顯示 未審核畫面
                PageControler1.Visible = True

        End Select

        msg.Text = ""
        PageControler1.PageDataTable = dt 'sqlstr
        'PageControler1.PrimaryKey = "OCID"
        'PageControler1.Sort = "ClassID,CyclType"
        PageControler1.ControlerLoad()
    End Sub

    '查詢鈕
    Private Sub Bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call ShowGrid()
    End Sub

    Private Sub DG_ClassInfo_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_ClassInfo.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        '學員資料審核
        Dim v_OCID As String = TIMS.GetMyValue(e.CommandArgument, "OCID")
        Dim v_MRqID As String = TIMS.Get_MRqID(Me)
        KeepSearchStr(v_OCID)
        Select Case e.CommandName
            Case "edit"
                'GetSearchStr() 'Response.Redirect("SD_03_002_classver.aspx?" & e.CommandArgument & "")
                Dim url1 As String = "SD_03_002_classver.aspx?ID=" & v_MRqID & "&OCID=" & v_OCID
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "view"
                Dim url1 As String = "SD_03_002.aspx?ID=" & v_MRqID & "&OCID=" & v_OCID 'e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_ClassInfo.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 

                'Dim s_ClassID As String = String.Concat(If(Len(drv("ClassID").ToString) < 4, "0", ""), drv("ClassID"))
                'Dim s_courName As String = String.Format("{0}{1}{2}", drv("Years"), s_ClassID, drv("CyclType"))
                'e.Item.Cells(2).Text = s_courName
                Dim date_str As String = String.Format("{0}<br>{1}", TIMS.Cdate3(drv("STDate")), TIMS.Cdate3(drv("FTDate")))
                e.Item.Cells(3).Text = date_str

                Dim but_edit As LinkButton = e.Item.FindControl("edit_but") '修改
                Dim but_view As LinkButton = e.Item.FindControl("view_but") '檢視
                Dim v_MRqID As String = TIMS.Get_MRqID(Me)
                but_edit.CommandArgument = String.Concat("OCID=", TIMS.ClearSQM(drv("OCID")), "&ID=", v_MRqID)
                but_view.CommandArgument = String.Concat("OCID=", TIMS.ClearSQM(drv("OCID")), "&ID=", v_MRqID)
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        '審核還原
        Dim v_OCID As String = TIMS.GetMyValue(e.CommandArgument, "OCID")
        Dim v_MRqID As String = TIMS.Get_MRqID(Me)
        KeepSearchStr(v_OCID)
        Select Case e.CommandName
            Case "edit2"
                Dim url1 As String = "SD_03_002_classver.aspx?act=R&ID=" & v_MRqID & "&OCID=" & v_OCID
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "view2"
                Dim url1 As String = "SD_03_002.aspx?ID=" & v_MRqID & "&OCID=" & v_OCID 'e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 

                'Dim s_ClassID As String = String.Concat(If(Len(drv("ClassID").ToString) < 4, "0", ""), drv("ClassID"))
                'Dim s_courName As String = String.Format("{0}{1}{2}", drv("Years"), s_ClassID, drv("CyclType"))
                'e.Item.Cells(2).Text = s_courName
                Dim date_str As String = String.Format("{0}<br>{1}", TIMS.Cdate3(drv("STDate")), TIMS.Cdate3(drv("FTDate")))
                e.Item.Cells(3).Text = date_str

                Dim but_edit As LinkButton = e.Item.FindControl("return_btn") '修改
                Dim but_view As LinkButton = e.Item.FindControl("view_btn") '檢視
                Dim v_MRqID As String = TIMS.Get_MRqID(Me)
                but_edit.CommandArgument = String.Concat("OCID=", TIMS.ClearSQM(drv("OCID")), "&ID=", v_MRqID)
                but_view.CommandArgument = String.Concat("OCID=", TIMS.ClearSQM(drv("OCID")), "&ID=", v_MRqID)
        End Select
    End Sub


    Sub KeepSearchStr(ByVal OCID1Value As String)
        OCID1Value = TIMS.ClearSQM(OCID1Value)
        center.Text = TIMS.ClearSQM(center.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        TMID1.Text = TIMS.ClearSQM(TMID1.Text)
        TMIDValue1.Value = TIMS.ClearSQM(TMIDValue1.Value)
        OCID1.Text = TIMS.ClearSQM(OCID1.Text)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)
        Dim v_NotOpen As String = TIMS.ClearSQM(NotOpen.SelectedValue)
        TxtPageSize.Text = TIMS.ClearSQM(TxtPageSize.Text)

        Dim SS_SearchStr As String = "Load=SD_03_002_ver"
        SS_SearchStr &= "&center=" & center.Text
        SS_SearchStr &= "&RIDValue=" & RIDValue.Value
        SS_SearchStr &= "&rTMID1=" & TMID1.Text 'TIMS.GET_OCIDInfo(OCID1Value, "TMID1")
        SS_SearchStr &= "&rTMIDValue1=" & TMIDValue1.Value 'TIMS.GET_OCIDInfo(OCID1Value, "TMIDValue1")
        SS_SearchStr &= "&rOCID1=" & OCID1.Text 'TIMS.GET_OCIDInfo(OCID1Value, "OCID1")
        SS_SearchStr &= "&rOCIDValue1=" & OCIDValue1.Value 'OCID1Value
        SS_SearchStr &= "&TMID1=" & TIMS.GET_OCIDInfo(OCID1Value, "TMID1", objconn)
        SS_SearchStr &= "&TMIDValue1=" & TIMS.GET_OCIDInfo(OCID1Value, "TMIDValue1", objconn)
        SS_SearchStr &= "&OCID1=" & TIMS.GET_OCIDInfo(OCID1Value, "OCID1", objconn)
        SS_SearchStr &= "&OCIDValue1=" & OCID1Value 'OCIDValue1.Value 'OCID1Value
        SS_SearchStr &= "&start_date=" & start_date.Text
        SS_SearchStr &= "&end_date=" & end_date.Text
        SS_SearchStr &= "&NotOpen=" & v_NotOpen 'NotOpen.SelectedValue
        SS_SearchStr &= "&TxtPageSize=" & TxtPageSize.Text
        SS_SearchStr &= "&submit=1"

        Session("_SearchStr") = SS_SearchStr
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable2.Visible = False
        DG_ClassInfo.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable2.Visible = False
        DG_ClassInfo.Visible = False
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class